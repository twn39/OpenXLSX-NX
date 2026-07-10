#ifndef OPENXLSX_XLXMLHELPERS_HPP
#define OPENXLSX_XLXMLHELPERS_HPP

/**
 * @file XLXmlHelpers.hpp
 * @brief Layer-2 OOXML DOM idioms built on the frozen XLXmlParser substrate.
 *
 * Design rules:
 *  - Depends only on XLXmlParser (no Styles / Document / Cell types).
 *  - Pure inline helpers; no second DOM model.
 *  - Prefer ensureChild / setAttr names for new code; appendAnd* aliases keep source compat.
 *  - XLXmlParser surface (child / attribute / append_*) stays stable; add domain semantics here.
 */

#include <initializer_list>
#include <string>
#include <string_view>
#include <type_traits>
#include <vector>

#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    // =========================================================================
    // Flags shared by setChildAttr / appendAndSetNodeAttribute
    // =========================================================================

    constexpr bool XLRemoveAttributes = true;
    constexpr bool XLKeepAttributes   = false;

    constexpr int SORT_INDEX_NOT_FOUND = -1;

    /**
     * @brief Index of @p nodeName in @p nodeOrder, or SORT_INDEX_NOT_FOUND.
     */
    inline int findStringInVector(std::string_view nodeName, std::vector<std::string_view> const& nodeOrder)
    {
        for (size_t i = 0; i < nodeOrder.size(); ++i)
            if (nodeName == nodeOrder[i]) return static_cast<int>(i);
        return SORT_INDEX_NOT_FOUND;
    }

    /**
     * @brief Copy leading pcdata (whitespace) siblings from @p fromNode before @p toNode.
     */
    inline void copyLeadingWhitespaces(XMLNode& parent, XMLNode fromNode, XMLNode toNode)
    {
        fromNode = fromNode.previous_sibling();
        while (fromNode.type() == node_pcdata) {
            toNode = parent.insert_child_before(node_pcdata, toNode);
            toNode.set_value(fromNode.value());
            fromNode = fromNode.previous_sibling();
        }
    }

    // =========================================================================
    // ensureChild — get-or-create child element (optional MS Office node order)
    // =========================================================================

    /**
     * @brief Ensure a child element named @p nodeName exists under @p parent and return it.
     * @param nodeOrder Optional MS Office–required element sequence for ordered insert.
     * @param force_ns  Forwarded to XMLNode insert/append (namespace forcing).
     */
    inline XMLNode ensureChild(XMLNode&                             parent,
                               const char*                          nodeName,
                               std::vector<std::string_view> const& nodeOrder = {},
                               bool                                 force_ns  = false)
    {
        if (parent.empty() || nodeName == nullptr || *nodeName == '\0') return XMLNode{};

        XMLNode nextNode = parent.first_child_of_type(node_element);
        if (nextNode.empty()) return parent.prepend_child(nodeName, force_ns);

        XMLNode node{};

        const int nodeSortIndex = (nodeOrder.size() > 1 ? findStringInVector(nodeName, nodeOrder) : SORT_INDEX_NOT_FOUND);
        if (nodeSortIndex != SORT_INDEX_NOT_FOUND) {
            while (not nextNode.empty() and findStringInVector(nextNode.raw_name(), nodeOrder) < nodeSortIndex)
                nextNode = nextNode.next_sibling_of_type(node_element);

            if (not nextNode.empty()) {
                if (std::string_view(nextNode.raw_name()) == std::string_view(nodeName))
                    node = nextNode;
                else {
                    node = parent.insert_child_before(nodeName, nextNode, force_ns);
                    copyLeadingWhitespaces(parent, node, nextNode);
                }
            }
        }
        else
            node = parent.child(nodeName);

        if (node.empty()) {
            nextNode = parent.last_child_of_type(node_element);
            node     = parent.insert_child_after(nodeName, nextNode, force_ns);
            copyLeadingWhitespaces(parent, nextNode, node);
        }

        return node;
    }

    inline XMLNode ensureChild(XMLNode&                             parent,
                               std::string const&                   nodeName,
                               std::vector<std::string_view> const& nodeOrder = {},
                               bool                                 force_ns  = false)
    {
        return ensureChild(parent, nodeName.c_str(), nodeOrder, force_ns);
    }

    /**
     * @brief Ensure a chain of simple (unordered) child elements exists.
     * @note Does not apply nodeOrder; use ensureChild per level when order matters.
     */
    inline XMLNode ensureChildPath(XMLNode& parent, std::initializer_list<const char*> path)
    {
        XMLNode current = parent;
        for (const char* name : path) {
            if (current.empty()) return XMLNode{};
            current = ensureChild(current, name);
        }
        return current;
    }

    // =========================================================================
    // ensureAttr / setAttr — attribute get-or-create / always-set
    // =========================================================================

    /**
     * @brief Ensure attribute exists; write @p defaultVal only when creating.
     */
    inline XMLAttribute ensureAttr(XMLNode& node, const char* attrName, const char* defaultVal = "")
    {
        if (node.empty() || attrName == nullptr || *attrName == '\0') return XMLAttribute{};
        XMLAttribute attr = node.attribute(attrName);
        if (attr.empty()) {
            attr = node.append_attribute(attrName);
            attr.set_value(defaultVal != nullptr ? defaultVal : "");
        }
        return attr;
    }

    inline XMLAttribute ensureAttr(XMLNode& node, std::string const& attrName, std::string const& defaultVal)
    {
        return ensureAttr(node, attrName.c_str(), defaultVal.c_str());
    }

    /**
     * @brief Ensure attribute exists and always set its value (string).
     */
    inline XMLAttribute setAttr(XMLNode& node, const char* attrName, const char* attrVal)
    {
        if (node.empty() || attrName == nullptr || *attrName == '\0') return XMLAttribute{};
        XMLAttribute attr = node.attribute(attrName);
        if (attr.empty()) attr = node.append_attribute(attrName);
        attr.set_value(attrVal != nullptr ? attrVal : "");
        return attr;
    }

    inline XMLAttribute setAttr(XMLNode& node, const char* attrName, std::string_view attrVal)
    {
        if (node.empty() || attrName == nullptr || *attrName == '\0') return XMLAttribute{};
        XMLAttribute attr = node.attribute(attrName);
        if (attr.empty()) attr = node.append_attribute(attrName);
        const std::string tmp(attrVal);
        attr.set_value(tmp.c_str());
        return attr;
    }

    inline XMLAttribute setAttr(XMLNode& node, const char* attrName, std::string const& attrVal)
    {
        return setAttr(node, attrName, attrVal.c_str());
    }

    inline XMLAttribute setAttr(XMLNode& node, std::string const& attrName, std::string const& attrVal)
    {
        return setAttr(node, attrName.c_str(), attrVal.c_str());
    }

    /**
     * @brief Ensure attribute exists and set a numeric / bool value via XMLAttribute::set_value overloads.
     * @note Constrained to arithmetic types so std::string rvalues bind to the string overloads above.
     */
    template<typename T, typename = std::enable_if_t<std::is_arithmetic_v<T>>>
    inline XMLAttribute setAttr(XMLNode& node, const char* attrName, T attrVal)
    {
        if (node.empty() || attrName == nullptr || *attrName == '\0') return XMLAttribute{};
        XMLAttribute attr = node.attribute(attrName);
        if (attr.empty()) attr = node.append_attribute(attrName);
        attr.set_value(attrVal);
        return attr;
    }

    /** @brief Excel-style boolean attribute as "1" / "0". */
    inline XMLAttribute setAttrBool01(XMLNode& node, const char* attrName, bool value)
    {
        return setAttr(node, attrName, value ? "1" : "0");
    }

    /** @brief Boolean attribute as "true" / "false". */
    inline XMLAttribute setAttrBoolTF(XMLNode& node, const char* attrName, bool value)
    {
        return setAttr(node, attrName, value ? "true" : "false");
    }

    // =========================================================================
    // setChildAttr / ensureChildAttr — child element + attribute
    // =========================================================================

    /**
     * @brief Ensure child @p nodeName exists and return its attribute @p attrName (create with default).
     */
    inline XMLAttribute ensureChildAttr(XMLNode&                             parent,
                                        const char*                          nodeName,
                                        const char*                          attrName,
                                        const char*                          attrDefaultVal,
                                        std::vector<std::string_view> const& nodeOrder = {})
    {
        if (parent.empty()) return XMLAttribute{};
        XMLNode node = ensureChild(parent, nodeName, nodeOrder);
        return ensureAttr(node, attrName, attrDefaultVal);
    }

    inline XMLAttribute ensureChildAttr(XMLNode&                             parent,
                                        std::string const&                   nodeName,
                                        std::string const&                   attrName,
                                        std::string const&                   attrDefaultVal,
                                        std::vector<std::string_view> const& nodeOrder = {})
    {
        return ensureChildAttr(parent, nodeName.c_str(), attrName.c_str(), attrDefaultVal.c_str(), nodeOrder);
    }

    /**
     * @brief Ensure child exists, optionally clear its attributes, then set @p attrName.
     */
    inline XMLAttribute setChildAttr(XMLNode&                             parent,
                                     const char*                          nodeName,
                                     const char*                          attrName,
                                     const char*                          attrVal,
                                     bool                                 removeAttributes = XLKeepAttributes,
                                     std::vector<std::string_view> const& nodeOrder        = {})
    {
        if (parent.empty()) return XMLAttribute{};
        XMLNode node = ensureChild(parent, nodeName, nodeOrder);
        if (removeAttributes) node.remove_attributes();
        return setAttr(node, attrName, attrVal);
    }

    inline XMLAttribute setChildAttr(XMLNode&                             parent,
                                     std::string const&                   nodeName,
                                     std::string const&                   attrName,
                                     std::string const&                   attrVal,
                                     bool                                 removeAttributes = XLKeepAttributes,
                                     std::vector<std::string_view> const& nodeOrder        = {})
    {
        return setChildAttr(parent, nodeName.c_str(), attrName.c_str(), attrVal.c_str(), removeAttributes, nodeOrder);
    }

    // =========================================================================
    // val-attribute & text helpers (spreadsheetml / chartml common patterns)
    // =========================================================================

    /** @brief Set the @val attribute on @p node. */
    inline XMLAttribute setVal(XMLNode& node, const char* value) { return setAttr(node, "val", value); }

    inline XMLAttribute setVal(XMLNode& node, std::string_view value) { return setAttr(node, "val", value); }

    /**
     * @brief Ensure child @p childName exists under @p parent and set its @val attribute.
     */
    inline XMLAttribute setChildVal(XMLNode&                             parent,
                                    const char*                          childName,
                                    const char*                          value,
                                    std::vector<std::string_view> const& nodeOrder = {})
    {
        return setChildAttr(parent, childName, "val", value, XLKeepAttributes, nodeOrder);
    }

    inline XMLAttribute setChildVal(XMLNode&                             parent,
                                    const char*                          childName,
                                    std::string_view                     value,
                                    std::vector<std::string_view> const& nodeOrder = {})
    {
        const std::string tmp(value);
        return setChildVal(parent, childName, tmp.c_str(), nodeOrder);
    }

    /**
     * @brief Ensure child exists (optional order) and set its text content.
     */
    inline XMLNode setChildText(XMLNode&                             parent,
                                const char*                          childName,
                                std::string_view                     text,
                                std::vector<std::string_view> const& nodeOrder = {})
    {
        XMLNode node = ensureChild(parent, childName, nodeOrder);
        if (not node.empty()) {
            const std::string tmp(text);
            node.text().set(tmp.c_str());
        }
        return node;
    }

    /**
     * @brief Bool tag with omitted @val meaning true (legacy OOXML quirk).
     * @note Creates and sets attr to "true" when omitted (preserves historical behavior).
     */
    inline bool getBoolAttributeWhenOmittedMeansTrue(XMLNode& parent, const char* tagName, const char* attrName = "val")
    {
        if (parent.empty() || tagName == nullptr) return false;
        XMLNode tagNode = parent.child(tagName);
        if (tagNode.empty()) return false;
        XMLAttribute valAttr = tagNode.attribute(attrName);
        if (valAttr.empty()) {
            setAttr(tagNode, attrName, "true");
            return true;
        }
        return valAttr.as_bool();
    }

    inline bool getBoolAttributeWhenOmittedMeansTrue(XMLNode& parent, std::string const& tagName, std::string const& attrName = "val")
    {
        return getBoolAttributeWhenOmittedMeansTrue(parent, tagName.c_str(), attrName.c_str());
    }

    // =========================================================================
    // Legacy names (appendAnd*) — thin aliases for source compatibility
    // =========================================================================

    inline XMLNode appendAndGetNode(XMLNode&                             parent,
                                    std::string const&                   nodeName,
                                    std::vector<std::string_view> const& nodeOrder = {},
                                    bool                                 force_ns  = false)
    {
        return ensureChild(parent, nodeName.c_str(), nodeOrder, force_ns);
    }

    inline XMLAttribute appendAndGetAttribute(XMLNode& node, std::string const& attrName, std::string const& attrDefaultVal)
    {
        return ensureAttr(node, attrName.c_str(), attrDefaultVal.c_str());
    }

    inline XMLAttribute appendAndSetAttribute(XMLNode& node, std::string const& attrName, std::string const& attrVal)
    {
        return setAttr(node, attrName.c_str(), attrVal.c_str());
    }

    inline XMLAttribute appendAndGetNodeAttribute(XMLNode&                             parent,
                                                  std::string const&                   nodeName,
                                                  std::string const&                   attrName,
                                                  std::string const&                   attrDefaultVal,
                                                  std::vector<std::string_view> const& nodeOrder = {})
    {
        return ensureChildAttr(parent, nodeName.c_str(), attrName.c_str(), attrDefaultVal.c_str(), nodeOrder);
    }

    inline XMLAttribute appendAndSetNodeAttribute(XMLNode&                             parent,
                                                  std::string const&                   nodeName,
                                                  std::string const&                   attrName,
                                                  std::string const&                   attrVal,
                                                  bool                                 removeAttributes = XLKeepAttributes,
                                                  std::vector<std::string_view> const& nodeOrder        = {})
    {
        return setChildAttr(parent, nodeName.c_str(), attrName.c_str(), attrVal.c_str(), removeAttributes, nodeOrder);
    }

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLXMLHELPERS_HPP

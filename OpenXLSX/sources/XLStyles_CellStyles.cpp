// ===== External Includes ===== //
#include <cstdint>
#include <fmt/format.h>
#include <gsl/gsl>
#include <memory>
#include <pugixml.hpp>
#include <stdexcept>
#include <string>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLColor.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLStyles.hpp"
#include "XLStyles_Internal.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

XLCellStyle::XLCellStyle() : m_cellStyleNode(std::make_unique<XMLNode>()) {}

XLCellStyle::XLCellStyle(const XMLNode& node) : m_cellStyleNode(std::make_unique<XMLNode>(node)) {}

XLCellStyle::~XLCellStyle() = default;

XLCellStyle::XLCellStyle(const XLCellStyle& other) : m_cellStyleNode(std::make_unique<XMLNode>(*other.m_cellStyleNode)) {}

XLCellStyle& XLCellStyle::operator=(const XLCellStyle& other)
{
    if (&other != this) *m_cellStyleNode = *other.m_cellStyleNode;
    return *this;
}

bool XLCellStyle::empty() const { return m_cellStyleNode->empty(); }

std::string  XLCellStyle::name() const { return m_cellStyleNode->attribute("name").value(); }
XLStyleIndex XLCellStyle::xfId() const { return m_cellStyleNode->attribute("xfId").as_uint(XLInvalidStyleIndex); }
uint32_t     XLCellStyle::builtinId() const { return m_cellStyleNode->attribute("builtinId").as_uint(XLInvalidUInt32); }
uint32_t     XLCellStyle::outlineStyle() const { return m_cellStyleNode->attribute("iLevel").as_uint(XLInvalidUInt32); }
bool         XLCellStyle::hidden() const { return m_cellStyleNode->attribute("hidden").as_bool(); }
bool         XLCellStyle::customBuiltin() const { return m_cellStyleNode->attribute("customBuiltin").as_bool(); }

bool XLCellStyle::setName(std::string_view newName)
{ return setAttr(*m_cellStyleNode, "name", std::string(newName)).empty() == false; }
bool XLCellStyle::setXfId(XLStyleIndex newXfId)
{ return setAttr(*m_cellStyleNode, "xfId", std::to_string(newXfId)).empty() == false; }
bool XLCellStyle::setBuiltinId(uint32_t newBuiltinId)
{ return setAttr(*m_cellStyleNode, "builtinId", std::to_string(newBuiltinId)).empty() == false; }
bool XLCellStyle::setOutlineStyle(uint32_t newOutlineStyle)
{ return setAttr(*m_cellStyleNode, "iLevel", std::to_string(newOutlineStyle)).empty() == false; }
bool XLCellStyle::setHidden(bool set)
{ return setAttr(*m_cellStyleNode, "hidden", (set ? "true" : "false")).empty() == false; }
bool XLCellStyle::setCustomBuiltin(bool set)
{ return setAttr(*m_cellStyleNode, "customBuiltin", (set ? "true" : "false")).empty() == false; }

/**
 * @brief Unsupported setter function
 */
bool XLCellStyle::setExtLst(XLUnsupportedElement const& newExtLst)
{
    OpenXLSX::ignore(newExtLst);
    return false;
}

std::string XLCellStyle::summary() const
{
    uint32_t iLevel = outlineStyle();
    return fmt::format("name={}, xfId={}, builtinId={}{}{}{}",
                       name(),
                       xfId(),
                       builtinId(),
                       iLevel != XLInvalidUInt32 ? fmt::format(", iLevel={}", outlineStyle()) : "",
                       hidden() ? ", hidden=true" : "",
                       customBuiltin() ? ", customBuiltin=true" : "");
}

// ===== XLCellStyles, parent of XLCellStyle

XLCellStyles::XLCellStyles() : m_cellStylesNode(std::make_unique<XMLNode>()) {}

XLCellStyles::XLCellStyles(const XMLNode& cellStyles) : m_cellStylesNode(std::make_unique<XMLNode>(cellStyles))
{
    // initialize XLCellStyles entries and m_cellStyles here
    XMLNode node = cellStyles.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string_view nodeName(node.name());
        if (nodeName == "cellStyle")
            m_cellStyles.push_back(XLCellStyle(node));
        else
            std::cerr << "WARNING: XLCellStyles constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLCellStyles::~XLCellStyles() { m_cellStyles.clear(); }

XLCellStyles::XLCellStyles(const XLCellStyles& other)
    : m_cellStylesNode(std::make_unique<XMLNode>(*other.m_cellStylesNode)),
      m_cellStyles(other.m_cellStyles)
{}

XLCellStyles::XLCellStyles(XLCellStyles&& other)
    : m_cellStylesNode(std::move(other.m_cellStylesNode)),
      m_cellStyles(std::move(other.m_cellStyles))
{}

XLCellStyles& XLCellStyles::operator=(const XLCellStyles& other)
{
    if (&other != this) {
        *m_cellStylesNode = *other.m_cellStylesNode;
        m_cellStyles.clear();
        m_cellStyles = other.m_cellStyles;
    }
    return *this;
}

size_t XLCellStyles::count() const { return m_cellStyles.size(); }

XLCellStyle XLCellStyles::cellStyleByIndex(XLStyleIndex index) const
{
    Expects(index < m_cellStyles.size());
    return m_cellStyles.at(index);
}

XLStyleIndex XLCellStyles::create(XLCellStyle copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();
    XMLNode      newNode{};

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_cellStylesNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty())
        newNode = m_cellStylesNode->prepend_child("cellStyle");
    else
        newNode = m_cellStylesNode->insert_child_after("cellStyle", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCellStyles::"s + __func__ + ": failed to append a new cellStyle node"s);
    }
    if (styleEntriesPrefix.length() > 0)
        m_cellStylesNode->insert_child_before(pugi::node_pcdata, newNode).set_value(std::string(styleEntriesPrefix).c_str());

    XLCellStyle newCellStyle(newNode);
    if (copyFrom.m_cellStyleNode->empty()) {
        // ===== Create a cell style with default values
        newCellStyle.setName("");
        newCellStyle.setXfId(0);
        newCellStyle.setHidden(false);
    }
    else
        copyXMLNode(newNode, *copyFrom.m_cellStyleNode);

    m_cellStyles.push_back(newCellStyle);
    setAttr(*m_cellStylesNode, "count", std::to_string(m_cellStyles.size()));
    return index;
}

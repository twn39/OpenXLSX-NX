// ===== External Includes ===== //
#include <cstring>    // strlen, memcpy, strcpy
#include <pugixml.hpp>
#include <string_view>

// // ===== OpenXLSX Includes ===== //
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    bool enable_xml_namespaces()
    {
        return true;
    }

    bool disable_xml_namespaces()
    {
        return true;
    }

    const pugi::char_t* XMLNode::name_without_namespace(const pugi::char_t* name_) const
    {
        if (!name_) return nullptr;
        const std::basic_string_view<pugi::char_t> nameView(name_);
        const auto                                 pos = nameView.find(':');
        if (pos == std::basic_string_view<pugi::char_t>::npos) return name_;
        return name_ + pos + 1;
    }

    

    NameProxy OpenXLSX_xml_node::namespaced_name_proxy(const pugi::char_t* name_, bool force_ns) const
    {
        if (!name_begin || force_ns) return NameProxy(name_);
        if (name_) {
            const char* colon = std::strchr(name_, ':');
            if (colon) return NameProxy(name_);
        }
        std::string namespaced_name_(xml_node::name(), name_begin);
        namespaced_name_.append(name_);
        return NameProxy(std::move(namespaced_name_));
    }

    XMLNode XMLNode::first_child_of_type(pugi::xml_node_type type_) const
    {
        if (_root) {
            XMLNode       x = first_child();
            const XMLNode l = last_child();
            while (x != l and x.type() != type_) x = x.next_sibling();
            if (x.type() == type_) return XMLNode(x);
        }
        return XMLNode();
    }

    XMLNode XMLNode::last_child_of_type(pugi::xml_node_type type_) const
    {
        if (_root) {
            const XMLNode f = first_child();
            XMLNode       x = last_child();
            while (x != f and x.type() != type_) x = x.previous_sibling();
            if (x.type() == type_) return XMLNode(x);
        }
        return XMLNode();
    }

    size_t XMLNode::child_count_of_type(pugi::xml_node_type type_) const
    {
        size_t counter = 0;
        if (_root) {
            XMLNode c = first_child_of_type(type_);
            while (!c.empty()) {
                ++counter;
                c = c.next_sibling_of_type(type_);
            }
        }
        return counter;
    }

    XMLNode XMLNode::next_sibling_of_type(pugi::xml_node_type type_) const
    {
        for (pugi::xml_node n = pugi::xml_node::next_sibling(); n; n = n.next_sibling()) {
            if (n.type() == type_) return XMLNode(n);
        }
        return XMLNode();
    }

    XMLNode XMLNode::previous_sibling_of_type(pugi::xml_node_type type_) const
    {
        for (pugi::xml_node n = pugi::xml_node::previous_sibling(); n; n = n.previous_sibling()) {
            if (n.type() == type_) return XMLNode(n);
        }
        return XMLNode();
    }

    XMLNode XMLNode::next_sibling_of_type(const pugi::char_t* name_, pugi::xml_node_type type_) const
    {
        const std::basic_string_view<pugi::char_t> targetName(name_);
        for (pugi::xml_node n = pugi::xml_node::next_sibling(); n; n = n.next_sibling()) {
            if (n.type() == type_ && std::basic_string_view<pugi::char_t>(n.name()) == targetName) return XMLNode(n);
        }
        return XMLNode();
    }

    XMLNode XMLNode::previous_sibling_of_type(const pugi::char_t* name_, pugi::xml_node_type type_) const
    {
        const std::basic_string_view<pugi::char_t> targetName(name_);
        for (pugi::xml_node n = pugi::xml_node::previous_sibling(); n; n = n.previous_sibling()) {
            if (n.type() == type_ && std::basic_string_view<pugi::char_t>(n.name()) == targetName) return XMLNode(n);
        }
        return XMLNode();
    }

}    // namespace OpenXLSX

// ===== External Includes ===== //
#include <cstring>    // strlen, memcpy, strcpy
#include <pugixml.hpp>
#include <string_view>

// ===== OpenXLSX Includes ===== //
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

    // ===== XMLNode implementation =====

    OpenXLSX_xml_node::OpenXLSX_xml_node() : m_node(nullptr), name_begin(0) {}

    OpenXLSX_xml_node::OpenXLSX_xml_node(pugi::xml_node_struct* node) : m_node(node), name_begin(0)
    {
        pugi::xml_node pugiNode(node);
        const char* name_ = pugiNode.name();
        if (name_ && name_[0]) {
            const char* colon = std::strchr(name_, ':');
            if (colon) {
                name_begin = (colon - name_) + 1;
            }
        }
    }

    OpenXLSX_xml_node::OpenXLSX_xml_node(const pugi::xml_node& node) : OpenXLSX_xml_node(node.internal_object()) {}

    bool OpenXLSX_xml_node::operator==(const OpenXLSX_xml_node& other) const { return m_node == other.m_node; }
    bool OpenXLSX_xml_node::operator!=(const OpenXLSX_xml_node& other) const { return m_node != other.m_node; }
    
    bool OpenXLSX_xml_node::operator<(const OpenXLSX_xml_node& other) const
    {
        return pugi::xml_node(m_node) < pugi::xml_node(other.m_node);
    }

    OpenXLSX_xml_node::operator bool() const { return m_node != nullptr; }
    bool OpenXLSX_xml_node::empty() const { return m_node == nullptr; }

    const char* OpenXLSX_xml_node::name() const
    {
        pugi::xml_node node(m_node);
        return &node.name()[name_begin];
    }

    const char* OpenXLSX_xml_node::raw_name() const
    {
        pugi::xml_node node(m_node);
        return node.name();
    }

    const char* OpenXLSX_xml_node::value() const
    {
        pugi::xml_node node(m_node);
        return node.value();
    }

    bool OpenXLSX_xml_node::set_name(const char* rhs, bool force_ns)
    {
        pugi::xml_node node(m_node);
        return node.set_name(NAMESPACED_NAME(rhs, force_ns));
    }

    bool OpenXLSX_xml_node::set_name(const char* rhs, size_t size, bool force_ns)
    {
        pugi::xml_node node(m_node);
        return node.set_name(NAMESPACED_NAME(rhs, force_ns), size + name_begin);
    }

    bool OpenXLSX_xml_node::set_value(const char* rhs)
    {
        pugi::xml_node node(m_node);
        return node.set_value(rhs);
    }

    XMLAttribute OpenXLSX_xml_node::attribute(const char* name_) const
    {
        pugi::xml_node node(m_node);
        return XMLAttribute(node.attribute(name_).internal_object());
    }

    XMLAttribute OpenXLSX_xml_node::first_attribute() const
    {
        pugi::xml_node node(m_node);
        return XMLAttribute(node.first_attribute().internal_object());
    }

    XMLAttribute OpenXLSX_xml_node::last_attribute() const
    {
        pugi::xml_node node(m_node);
        return XMLAttribute(node.last_attribute().internal_object());
    }

    XMLAttribute OpenXLSX_xml_node::append_attribute(const char* name_)
    {
        pugi::xml_node node(m_node);
        return XMLAttribute(node.append_attribute(name_).internal_object());
    }

    XMLAttribute OpenXLSX_xml_node::prepend_attribute(const char* name_)
    {
        pugi::xml_node node(m_node);
        return XMLAttribute(node.prepend_attribute(name_).internal_object());
    }

    bool OpenXLSX_xml_node::remove_attribute(const char* name_)
    {
        pugi::xml_node node(m_node);
        return node.remove_attribute(name_);
    }

    bool OpenXLSX_xml_node::remove_attribute(const XMLAttribute& a)
    {
        pugi::xml_node node(m_node);
        return node.remove_attribute(pugi::xml_attribute(a.internal_object()));
    }

    void OpenXLSX_xml_node::remove_attributes()
    {
        pugi::xml_node node(m_node);
        node.remove_attributes();
    }

    XMLNode OpenXLSX_xml_node::parent() const
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.parent().internal_object());
    }

    XMLNode OpenXLSX_xml_node::child(const char* name_) const
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.child(NAMESPACED_NAME(name_, false)).internal_object());
    }

    XMLNode OpenXLSX_xml_node::next_sibling(const char* name_) const
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.next_sibling(NAMESPACED_NAME(name_, false)).internal_object());
    }

    XMLNode OpenXLSX_xml_node::next_sibling() const
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.next_sibling().internal_object());
    }

    XMLNode OpenXLSX_xml_node::previous_sibling(const char* name_) const
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.previous_sibling(NAMESPACED_NAME(name_, false)).internal_object());
    }

    XMLNode OpenXLSX_xml_node::previous_sibling() const
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.previous_sibling().internal_object());
    }

    const char* OpenXLSX_xml_node::child_value() const
    {
        pugi::xml_node node(m_node);
        return node.child_value();
    }

    const char* OpenXLSX_xml_node::child_value(const char* name_) const
    {
        pugi::xml_node node(m_node);
        return node.child_value(NAMESPACED_NAME(name_, false));
    }

    XMLNode OpenXLSX_xml_node::first_child() const
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.first_child().internal_object());
    }

    XMLNode OpenXLSX_xml_node::last_child() const
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.last_child().internal_object());
    }

    XMLNode OpenXLSX_xml_node::append_child(const char* name_, bool force_ns)
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.append_child(NAMESPACED_NAME(name_, force_ns)).internal_object());
    }

    XMLNode OpenXLSX_xml_node::prepend_child(const char* name_, bool force_ns)
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.prepend_child(NAMESPACED_NAME(name_, force_ns)).internal_object());
    }

    XMLNode OpenXLSX_xml_node::insert_child_after(const char* name_, const XMLNode& refNode, bool force_ns)
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.insert_child_after(NAMESPACED_NAME(name_, force_ns), pugi::xml_node(refNode.internal_object())).internal_object());
    }

    XMLNode OpenXLSX_xml_node::insert_child_before(const char* name_, const XMLNode& refNode, bool force_ns)
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.insert_child_before(NAMESPACED_NAME(name_, force_ns), pugi::xml_node(refNode.internal_object())).internal_object());
    }

    bool OpenXLSX_xml_node::remove_child(const char* name_)
    {
        pugi::xml_node node(m_node);
        return node.remove_child(NAMESPACED_NAME(name_, false));
    }

    bool OpenXLSX_xml_node::remove_child(const XMLNode& n)
    {
        pugi::xml_node node(m_node);
        return node.remove_child(pugi::xml_node(n.internal_object()));
    }

    void OpenXLSX_xml_node::remove_children()
    {
        pugi::xml_node node(m_node);
        node.remove_children();
    }

    XMLNode OpenXLSX_xml_node::find_child_by_attribute(const char* name_, const char* attr_name, const char* attr_value) const
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.find_child_by_attribute(NAMESPACED_NAME(name_, false), attr_name, attr_value).internal_object());
    }

    XMLNode OpenXLSX_xml_node::find_child_by_attribute(const char* attr_name, const char* attr_value) const
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.find_child_by_attribute(attr_name, attr_value).internal_object());
    }

    XMLNode OpenXLSX_xml_node::find_child_impl(std::function<bool(pugi::xml_node_struct*)> pred) const
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.find_child([&pred](const pugi::xml_node& n) {
            return pred(n.internal_object());
        }).internal_object());
    }

    pugi::xpath_node OpenXLSX_xml_node::select_node(const char* query) const
    {
        pugi::xml_node node(m_node);
        return node.select_node(query);
    }

    pugi::xpath_node_set OpenXLSX_xml_node::select_nodes(const char* query) const
    {
        pugi::xml_node node(m_node);
        return node.select_nodes(query);
    }

    XMLText OpenXLSX_xml_node::text() const
    {
        return XMLText(m_node);
    }

    XMLNodeType OpenXLSX_xml_node::type() const
    {
        pugi::xml_node node(m_node);
        return static_cast<XMLNodeType>(node.type());
    }

    void OpenXLSX_xml_node::print(std::ostream& os, const char* indent, unsigned int flags) const
    {
        pugi::xml_node node(m_node);
        node.print(os, indent, flags);
    }

    const char* OpenXLSX_xml_node::name_without_namespace(const char* name_) const
    {
        if (!name_) return nullptr;
        const std::string_view nameView(name_);
        const auto             pos = nameView.find(':');
        if (pos == std::string_view::npos) return name_;
        return name_ + pos + 1;
    }

    NameProxy OpenXLSX_xml_node::namespaced_name_proxy(const char* name_, bool force_ns) const
    {
        if (!name_begin || force_ns) return NameProxy(name_);
        if (name_) {
            const char* colon = std::strchr(name_, ':');
            if (colon) return NameProxy(name_);
        }
        pugi::xml_node node(m_node);
        std::string namespaced_name_(node.name(), name_begin);
        namespaced_name_.append(name_);
        return NameProxy(std::move(namespaced_name_));
    }

    // Type-erased helper implementations

    XMLNode OpenXLSX_xml_node::append_child_impl(int type_)
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.append_child(static_cast<pugi::xml_node_type>(type_)).internal_object());
    }

    XMLNode OpenXLSX_xml_node::prepend_child_impl(int type_)
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.prepend_child(static_cast<pugi::xml_node_type>(type_)).internal_object());
    }

    XMLNode OpenXLSX_xml_node::insert_child_after_impl(int type_, const XMLNode& refNode)
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.insert_child_after(static_cast<pugi::xml_node_type>(type_), pugi::xml_node(refNode.internal_object())).internal_object());
    }

    XMLNode OpenXLSX_xml_node::insert_child_before_impl(int type_, const XMLNode& refNode)
    {
        pugi::xml_node node(m_node);
        return XMLNode(node.insert_child_before(static_cast<pugi::xml_node_type>(type_), pugi::xml_node(refNode.internal_object())).internal_object());
    }

    XMLNode OpenXLSX_xml_node::first_child_of_type_impl(int type_) const
    {
        pugi::xml_node node(m_node);
        if (node) {
            pugi::xml_node x = node.first_child();
            pugi::xml_node l = node.last_child();
            pugi::xml_node_type pugiType = static_cast<pugi::xml_node_type>(type_);
            while (x != l && x.type() != pugiType) x = x.next_sibling();
            if (x.type() == pugiType) return XMLNode(x.internal_object());
        }
        return XMLNode();
    }

    XMLNode OpenXLSX_xml_node::last_child_of_type_impl(int type_) const
    {
        pugi::xml_node node(m_node);
        if (node) {
            pugi::xml_node f = node.first_child();
            pugi::xml_node x = node.last_child();
            pugi::xml_node_type pugiType = static_cast<pugi::xml_node_type>(type_);
            while (x != f && x.type() != pugiType) x = x.previous_sibling();
            if (x.type() == pugiType) return XMLNode(x.internal_object());
        }
        return XMLNode();
    }

    size_t OpenXLSX_xml_node::child_count_of_type_impl(int type_) const
    {
        size_t counter = 0;
        pugi::xml_node node(m_node);
        if (node) {
            XMLNode c = first_child_of_type_impl(type_);
            while (!c.empty()) {
                ++counter;
                c = c.next_sibling_of_type_impl(type_);
            }
        }
        return counter;
    }

    XMLNode OpenXLSX_xml_node::next_sibling_of_type_impl(int type_) const
    {
        pugi::xml_node node(m_node);
        pugi::xml_node_type pugiType = static_cast<pugi::xml_node_type>(type_);
        for (pugi::xml_node n = node.next_sibling(); n; n = n.next_sibling()) {
            if (n.type() == pugiType) return XMLNode(n.internal_object());
        }
        return XMLNode();
    }

    XMLNode OpenXLSX_xml_node::previous_sibling_of_type_impl(int type_) const
    {
        pugi::xml_node node(m_node);
        pugi::xml_node_type pugiType = static_cast<pugi::xml_node_type>(type_);
        for (pugi::xml_node n = node.previous_sibling(); n; n = n.previous_sibling()) {
            if (n.type() == pugiType) return XMLNode(n.internal_object());
        }
        return XMLNode();
    }

    XMLNode OpenXLSX_xml_node::next_sibling_of_type_name_impl(const char* name_, int type_) const
    {
        pugi::xml_node node(m_node);
        pugi::xml_node_type pugiType = static_cast<pugi::xml_node_type>(type_);
        const std::string_view targetName(name_);
        for (pugi::xml_node n = node.next_sibling(); n; n = n.next_sibling()) {
            if (n.type() == pugiType && std::string_view(n.name()) == targetName) return XMLNode(n.internal_object());
        }
        return XMLNode();
    }

    XMLNode OpenXLSX_xml_node::previous_sibling_of_type_name_impl(const char* name_, int type_) const
    {
        pugi::xml_node node(m_node);
        pugi::xml_node_type pugiType = static_cast<pugi::xml_node_type>(type_);
        const std::string_view targetName(name_);
        for (pugi::xml_node n = node.previous_sibling(); n; n = n.previous_sibling()) {
            if (n.type() == pugiType && std::string_view(n.name()) == targetName) return XMLNode(n.internal_object());
        }
        return XMLNode();
    }

    // Copy and move methods

    XMLNode OpenXLSX_xml_node::append_copy(const XMLNode& node)
    {
        pugi::xml_node n(m_node);
        return XMLNode(n.append_copy(pugi::xml_node(node.internal_object())).internal_object());
    }

    XMLNode OpenXLSX_xml_node::prepend_copy(const XMLNode& node)
    {
        pugi::xml_node n(m_node);
        return XMLNode(n.prepend_copy(pugi::xml_node(node.internal_object())).internal_object());
    }

    XMLNode OpenXLSX_xml_node::insert_copy_after(const XMLNode& node, const XMLNode& ref)
    {
        pugi::xml_node n(m_node);
        return XMLNode(n.insert_copy_after(pugi::xml_node(node.internal_object()), pugi::xml_node(ref.internal_object())).internal_object());
    }

    XMLNode OpenXLSX_xml_node::insert_copy_before(const XMLNode& node, const XMLNode& ref)
    {
        pugi::xml_node n(m_node);
        return XMLNode(n.insert_copy_before(pugi::xml_node(node.internal_object()), pugi::xml_node(ref.internal_object())).internal_object());
    }

    XMLNode OpenXLSX_xml_node::append_move(const XMLNode& node)
    {
        pugi::xml_node n(m_node);
        return XMLNode(n.append_move(pugi::xml_node(node.internal_object())).internal_object());
    }

    XMLNode OpenXLSX_xml_node::prepend_move(const XMLNode& node)
    {
        pugi::xml_node n(m_node);
        return XMLNode(n.prepend_move(pugi::xml_node(node.internal_object())).internal_object());
    }

    XMLNode OpenXLSX_xml_node::insert_move_after(const XMLNode& node, const XMLNode& ref)
    {
        pugi::xml_node n(m_node);
        return XMLNode(n.insert_move_after(pugi::xml_node(node.internal_object()), pugi::xml_node(ref.internal_object())).internal_object());
    }

    XMLNode OpenXLSX_xml_node::insert_move_before(const XMLNode& node, const XMLNode& ref)
    {
        pugi::xml_node n(m_node);
        return XMLNode(n.insert_move_before(pugi::xml_node(node.internal_object()), pugi::xml_node(ref.internal_object())).internal_object());
    }

    XMLAttribute OpenXLSX_xml_node::append_copy(const XMLAttribute& attr)
    {
        pugi::xml_node n(m_node);
        return XMLAttribute(n.append_copy(pugi::xml_attribute(attr.internal_object())).internal_object());
    }

    OpenXLSX_xml_node::operator pugi::xml_node() const
    {
        return pugi::xml_node(m_node);
    }

    XMLNodeRange OpenXLSX_xml_node::children() const
    {
        pugi::xml_node node(m_node);
        return XMLNodeRange(node.first_child().internal_object());
    }

    XMLNodeRange OpenXLSX_xml_node::children(const char* name_) const
    {
        pugi::xml_node node(m_node);
        return XMLNodeRange(node.first_child().internal_object(), name_);
    }

    XMLAttributeRange OpenXLSX_xml_node::attributes() const
    {
        pugi::xml_node node(m_node);
        return XMLAttributeRange(node.first_attribute().internal_object());
    }

    // ===== XMLAttribute implementation =====

    XMLAttribute::XMLAttribute() : m_attr(nullptr) {}
    XMLAttribute::XMLAttribute(pugi::xml_attribute_struct* attr) : m_attr(attr) {}
    XMLAttribute::XMLAttribute(const pugi::xml_attribute& attr) : m_attr(attr.internal_object()) {}

    bool XMLAttribute::operator==(const XMLAttribute& other) const { return m_attr == other.m_attr; }
    bool XMLAttribute::operator!=(const XMLAttribute& other) const { return m_attr != other.m_attr; }
    XMLAttribute::operator bool() const { return m_attr != nullptr; }
    bool XMLAttribute::empty() const { return m_attr == nullptr; }

    const char* XMLAttribute::name() const
    {
        pugi::xml_attribute attr(m_attr);
        return attr.name();
    }

    const char* XMLAttribute::value() const
    {
        pugi::xml_attribute attr(m_attr);
        return attr.value();
    }

    bool XMLAttribute::set_name(const char* rhs)
    {
        pugi::xml_attribute attr(m_attr);
        return attr.set_name(rhs);
    }

    bool XMLAttribute::set_value(const char* rhs)
    {
        pugi::xml_attribute attr(m_attr);
        return attr.set_value(rhs);
    }

    bool XMLAttribute::set_value(int rhs)
    {
        pugi::xml_attribute attr(m_attr);
        return attr.set_value(rhs);
    }

    bool XMLAttribute::set_value(unsigned int rhs)
    {
        pugi::xml_attribute attr(m_attr);
        return attr.set_value(rhs);
    }

    bool XMLAttribute::set_value(long rhs)
    {
        pugi::xml_attribute attr(m_attr);
        return attr.set_value(rhs);
    }

    bool XMLAttribute::set_value(unsigned long rhs)
    {
        pugi::xml_attribute attr(m_attr);
        return attr.set_value(rhs);
    }

    bool XMLAttribute::set_value(double rhs)
    {
        pugi::xml_attribute attr(m_attr);
        return attr.set_value(rhs);
    }

    bool XMLAttribute::set_value(bool rhs)
    {
        pugi::xml_attribute attr(m_attr);
        return attr.set_value(rhs);
    }

    bool XMLAttribute::set_value(long long rhs)
    {
        pugi::xml_attribute attr(m_attr);
        return attr.set_value(rhs);
    }

    bool XMLAttribute::set_value(unsigned long long rhs)
    {
        pugi::xml_attribute attr(m_attr);
        return attr.set_value(rhs);
    }

    const char* XMLAttribute::as_string(const char* def) const
    {
        pugi::xml_attribute attr(m_attr);
        return attr.as_string(def);
    }

    int XMLAttribute::as_int(int def) const
    {
        pugi::xml_attribute attr(m_attr);
        return attr.as_int(def);
    }

    unsigned int XMLAttribute::as_uint(unsigned int def) const
    {
        pugi::xml_attribute attr(m_attr);
        return attr.as_uint(def);
    }

    double XMLAttribute::as_double(double def) const
    {
        pugi::xml_attribute attr(m_attr);
        return attr.as_double(def);
    }

    float XMLAttribute::as_float(float def) const
    {
        pugi::xml_attribute attr(m_attr);
        return attr.as_float(def);
    }

    bool XMLAttribute::as_bool(bool def) const
    {
        pugi::xml_attribute attr(m_attr);
        return attr.as_bool(def);
    }

    unsigned long long XMLAttribute::as_ullong(unsigned long long def) const
    {
        pugi::xml_attribute attr(m_attr);
        return attr.as_ullong(def);
    }

    long long XMLAttribute::as_llong(long long def) const
    {
        pugi::xml_attribute attr(m_attr);
        return attr.as_llong(def);
    }

    XMLAttribute XMLAttribute::next_attribute() const
    {
        pugi::xml_attribute attr(m_attr);
        return XMLAttribute(attr.next_attribute().internal_object());
    }

    XMLAttribute XMLAttribute::previous_attribute() const
    {
        pugi::xml_attribute attr(m_attr);
        return XMLAttribute(attr.previous_attribute().internal_object());
    }

    XMLAttribute::operator pugi::xml_attribute() const
    {
        return pugi::xml_attribute(m_attr);
    }

    // ===== XMLText implementation =====

    XMLText::XMLText() : m_node(nullptr) {}
    XMLText::XMLText(pugi::xml_node_struct* node) : m_node(node) {}
    bool XMLText::empty() const { return m_node == nullptr; }

    const char* XMLText::get() const
    {
        pugi::xml_node node(m_node);
        return node.text().get();
    }

    const char* XMLText::as_string(const char* def) const
    {
        pugi::xml_node node(m_node);
        return node.text().as_string(def);
    }

    int XMLText::as_int(int def) const
    {
        pugi::xml_node node(m_node);
        return node.text().as_int(def);
    }

    unsigned int XMLText::as_uint(unsigned int def) const
    {
        pugi::xml_node node(m_node);
        return node.text().as_uint(def);
    }

    double XMLText::as_double(double def) const
    {
        pugi::xml_node node(m_node);
        return node.text().as_double(def);
    }

    float XMLText::as_float(float def) const
    {
        pugi::xml_node node(m_node);
        return node.text().as_float(def);
    }

    bool XMLText::as_bool(bool def) const
    {
        pugi::xml_node node(m_node);
        return node.text().as_bool(def);
    }

    unsigned long long XMLText::as_ullong(unsigned long long def) const
    {
        pugi::xml_node node(m_node);
        return node.text().as_ullong(def);
    }

    long long XMLText::as_llong(long long def) const
    {
        pugi::xml_node node(m_node);
        return node.text().as_llong(def);
    }

    bool XMLText::set(const char* rhs)
    {
        pugi::xml_node node(m_node);
        return node.text().set(rhs);
    }

    bool XMLText::set(int rhs)
    {
        pugi::xml_node node(m_node);
        return node.text().set(rhs);
    }

    bool XMLText::set(unsigned int rhs)
    {
        pugi::xml_node node(m_node);
        return node.text().set(rhs);
    }

    bool XMLText::set(long rhs)
    {
        pugi::xml_node node(m_node);
        return node.text().set(rhs);
    }

    bool XMLText::set(unsigned long rhs)
    {
        pugi::xml_node node(m_node);
        return node.text().set(rhs);
    }

    bool XMLText::set(double rhs)
    {
        pugi::xml_node node(m_node);
        return node.text().set(rhs);
    }

    bool XMLText::set(bool rhs)
    {
        pugi::xml_node node(m_node);
        return node.text().set(rhs);
    }

    bool XMLText::set(long long rhs)
    {
        pugi::xml_node node(m_node);
        return node.text().set(rhs);
    }

    bool XMLText::set(unsigned long long rhs)
    {
        pugi::xml_node node(m_node);
        return node.text().set(rhs);
    }

    // ===== XMLNodeIterator implementation =====

    XMLNodeIterator::XMLNodeIterator() : m_node(nullptr), m_filter(nullptr), m_wrap(nullptr) {}
    XMLNodeIterator::XMLNodeIterator(pugi::xml_node_struct* node, const char* filter)
        : m_node(node), m_filter(filter), m_wrap(nullptr)
    {
        if (m_node && m_filter) {
            pugi::xml_node n(m_node);
            const std::string_view filterView(m_filter);
            while (n && std::string_view(n.name()) != filterView) {
                n = n.next_sibling();
            }
            m_node = n.internal_object();
        }
        m_wrap = XMLNode(m_node);
    }

    XMLNode& XMLNodeIterator::operator*() const { return m_wrap; }

    XMLNodeIterator& XMLNodeIterator::operator++()
    {
        if (m_node) {
            pugi::xml_node n(m_node);
            if (m_filter) {
                const std::string_view filterView(m_filter);
                n = n.next_sibling();
                while (n && std::string_view(n.name()) != filterView) {
                    n = n.next_sibling();
                }
                m_node = n.internal_object();
            } else {
                m_node = n.next_sibling().internal_object();
            }
        }
        m_wrap = XMLNode(m_node);
        return *this;
    }

    bool XMLNodeIterator::operator==(const XMLNodeIterator& other) const { return m_node == other.m_node; }
    bool XMLNodeIterator::operator!=(const XMLNodeIterator& other) const { return m_node != other.m_node; }

    XMLNodeRange::XMLNodeRange(pugi::xml_node_struct* first, const char* filter)
        : m_first(first), m_filter(filter) {}
    XMLNodeIterator XMLNodeRange::begin() const { return XMLNodeIterator(m_first, m_filter); }
    XMLNodeIterator XMLNodeRange::end() const { return XMLNodeIterator(nullptr, m_filter); }

    // ===== XMLAttributeIterator implementation =====

    XMLAttributeIterator::XMLAttributeIterator() : m_attr(nullptr), m_wrap(nullptr) {}
    XMLAttributeIterator::XMLAttributeIterator(pugi::xml_attribute_struct* attr)
        : m_attr(attr), m_wrap(nullptr)
    {
        m_wrap = XMLAttribute(m_attr);
    }

    XMLAttribute& XMLAttributeIterator::operator*() const { return m_wrap; }

    XMLAttributeIterator& XMLAttributeIterator::operator++()
    {
        if (m_attr) {
            pugi::xml_attribute attr(m_attr);
            m_attr = attr.next_attribute().internal_object();
        }
        m_wrap = XMLAttribute(m_attr);
        return *this;
    }

    bool XMLAttributeIterator::operator==(const XMLAttributeIterator& other) const { return m_attr == other.m_attr; }
    bool XMLAttributeIterator::operator!=(const XMLAttributeIterator& other) const { return m_attr != other.m_attr; }

    XMLAttributeRange::XMLAttributeRange(pugi::xml_attribute_struct* first) : m_first(first) {}
    XMLAttributeIterator XMLAttributeRange::begin() const { return XMLAttributeIterator(m_first); }
    XMLAttributeIterator XMLAttributeRange::end() const { return XMLAttributeIterator(nullptr); }

    // ===== XMLDocument implementation =====

    OpenXLSX_xml_document::OpenXLSX_xml_document()
        : XMLNode(),
          m_doc(std::make_unique<pugi::xml_document>())
    {
        m_node = m_doc->internal_object();
        name_begin = 0;
    }

    OpenXLSX_xml_document::~OpenXLSX_xml_document() = default;

    OpenXLSX_xml_document::OpenXLSX_xml_document(OpenXLSX_xml_document&& other) noexcept
        : XMLNode(std::move(other)),
          m_doc(std::move(other.m_doc))
    {}

    OpenXLSX_xml_document& OpenXLSX_xml_document::operator=(OpenXLSX_xml_document&& other) noexcept
    {
        if (this != &other) {
            XMLNode::operator=(std::move(other));
            m_doc = std::move(other.m_doc);
        }
        return *this;
    }

    XMLNode OpenXLSX_xml_document::document_element() const
    {
        return XMLNode(m_doc->document_element().internal_object());
    }

    void OpenXLSX_xml_document::print(std::ostream& ostr) const
    {
        m_doc->print(ostr);
    }

    XMLParseResult OpenXLSX_xml_document::load_string(const char* contents, unsigned int options)
    {
        pugi::xml_parse_result res = m_doc->load_string(contents, options);
        m_node = m_doc->internal_object();
        name_begin = 0;
        
        XMLParseResult ret;
        ret.status = static_cast<XMLParseStatus>(res.status);
        ret.m_description = res.description();
        ret.m_bool = res;
        return ret;
    }

    XMLParseResult OpenXLSX_xml_document::load_buffer(const void* contents, size_t size, unsigned int settings)
    {
        pugi::xml_parse_result res = m_doc->load_buffer(contents, size, settings);
        m_node = m_doc->internal_object();
        name_begin = 0;
        
        XMLParseResult ret;
        ret.status = static_cast<XMLParseStatus>(res.status);
        ret.m_description = res.description();
        ret.m_bool = res;
        return ret;
    }

    void OpenXLSX_xml_document::save(pugi::xml_writer& writer, const char* indent, unsigned int flags) const
    {
        m_doc->save(writer, indent, flags);
    }

    void OpenXLSX_xml_document::save(std::ostream& stream, const char* indent, unsigned int flags) const
    {
        m_doc->save(stream, indent, flags);
    }

    void OpenXLSX_xml_document::reset()
    {
        m_doc->reset();
        m_node = m_doc->internal_object();
        name_begin = 0;
    }

    void OpenXLSX_xml_document::reset(const XMLDocument& proto)
    {
        m_doc->reset(*proto.m_doc);
        m_node = m_doc->internal_object();
        name_begin = 0;
    }

    OpenXLSX_xml_document::operator pugi::xml_document&()
    {
        return *m_doc;
    }

    OpenXLSX_xml_document::operator const pugi::xml_document&() const
    {
        return *m_doc;
    }

} // namespace OpenXLSX

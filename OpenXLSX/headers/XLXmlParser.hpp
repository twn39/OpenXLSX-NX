#ifndef OPENXLSX_XLXMLPARSER_HPP
#define OPENXLSX_XLXMLPARSER_HPP

#include <memory>    // shared_ptr, unique_ptr
#include <cstring>   // std::strchr
#include <string>
#include <ostream>
#include <iterator>  // std::ptrdiff_t, std::forward_iterator_tag
#include <functional> // std::function

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"

// Forward declarations for pugixml structs and classes to avoid including pugixml.hpp
namespace pugi
{
    struct xml_node_struct;
    struct xml_attribute_struct;
    class xml_node;
    class xml_attribute;
    class xml_document;
    class xml_writer;
    class xpath_node;
    class xpath_node_set;
}

namespace OpenXLSX
{
#define ENABLE_XML_NAMESPACES 1

#ifdef ENABLE_XML_NAMESPACES
    #define NAMESPACED_NAME(name_, force_ns_) namespaced_name_proxy(name_, force_ns_).c_str()
#else
    #define NAMESPACED_NAME(name_, force_ns_) name_
#endif

    constexpr const bool XLForceNamespace = true;

    bool enable_xml_namespaces();
    bool disable_xml_namespaces();

    // Forward declarations
    class OpenXLSX_xml_node;
    class OpenXLSX_xml_document;
    class XMLAttribute;
    class XMLText;
    class XMLNodeIterator;
    class XMLNodeRange;
    class XMLAttributeIterator;
    class XMLAttributeRange;

    using XMLNode      = OpenXLSX_xml_node;
    using XMLDocument  = OpenXLSX_xml_document;

    // XMLNodeType as int with constants to prevent -Wenum-compare warnings when comparing with pugixml enums
    using XMLNodeType = int;
    constexpr XMLNodeType node_null = 0;
    constexpr XMLNodeType node_document = 1;
    constexpr XMLNodeType node_element = 2;
    constexpr XMLNodeType node_pcdata = 3;
    constexpr XMLNodeType node_cdata = 4;
    constexpr XMLNodeType node_comment = 5;
    constexpr XMLNodeType node_declaration = 6;
    constexpr XMLNodeType node_doctype = 7;
    constexpr XMLNodeType node_pi = 8;

    // XMLParseStatus mapping
    using XMLParseStatus = int;
    constexpr XMLParseStatus status_ok = 0;
    constexpr XMLParseStatus status_file_not_found = 1;
    constexpr XMLParseStatus status_io_error = 2;
    constexpr XMLParseStatus status_out_of_memory = 3;
    constexpr XMLParseStatus status_internal_error = 4;
    constexpr XMLParseStatus status_unrecognized_tag = 5;
    constexpr XMLParseStatus status_bad_pi = 6;
    constexpr XMLParseStatus status_bad_comment = 7;
    constexpr XMLParseStatus status_bad_cdata = 8;
    constexpr XMLParseStatus status_bad_doctype = 9;
    constexpr XMLParseStatus status_bad_pcdata = 10;
    constexpr XMLParseStatus status_bad_start_element = 11;
    constexpr XMLParseStatus status_bad_attribute = 12;
    constexpr XMLParseStatus status_bad_end_element = 13;
    constexpr XMLParseStatus status_end_element_mismatch = 14;
    constexpr XMLParseStatus status_append_invalid_root = 15;
    constexpr XMLParseStatus status_no_document_element = 16;

    struct OPENXLSX_EXPORT XMLParseResult {
        explicit operator bool() const { return m_bool; }
        const char* description() const { return m_description; }
        XMLParseStatus status{status_ok};
        
        bool m_bool{false};
        const char* m_description{""};
    };

    class NameProxy {
        std::string m_str;
        const char* m_ptr;
    public:
        NameProxy(const char* ptr) : m_str(), m_ptr(ptr) {}
        NameProxy(std::string str) : m_str(std::move(str)), m_ptr(m_str.c_str()) {}
        const char* c_str() const { return m_ptr; }
    };

    // ===== Custom OpenXLSX_xml_node wrapper class =====
    class OPENXLSX_EXPORT OpenXLSX_xml_node
    {
    public:
        OpenXLSX_xml_node();
        OpenXLSX_xml_node(pugi::xml_node_struct* node);
        OpenXLSX_xml_node(const pugi::xml_node& node); // Implicit conversion constructor

        bool operator==(const OpenXLSX_xml_node& other) const;
        bool operator!=(const OpenXLSX_xml_node& other) const;
        bool operator<(const OpenXLSX_xml_node& other) const;
        explicit operator bool() const;
        bool empty() const;

        const char* name() const;
        const char* raw_name() const;
        const char* value() const;
        bool set_name(const char* rhs, bool force_ns = false);
        bool set_name(const char* rhs, size_t size, bool force_ns = false);
        bool set_value(const char* rhs);

        XMLAttribute attribute(const char* name_) const;
        XMLAttribute first_attribute() const;
        XMLAttribute last_attribute() const;
        XMLAttribute append_attribute(const char* name_);
        XMLAttribute prepend_attribute(const char* name_);
        bool remove_attribute(const char* name_);
        bool remove_attribute(const XMLAttribute& a);
        void remove_attributes();

        XMLNode parent() const;
        XMLNode child(const char* name_) const;
        XMLNode next_sibling(const char* name_) const;
        XMLNode next_sibling() const;
        XMLNode previous_sibling(const char* name_) const;
        XMLNode previous_sibling() const;
        const char* child_value() const;
        const char* child_value(const char* name_) const;

        XMLNode first_child() const;
        XMLNode last_child() const;

        template<typename T = XMLNodeType>
        XMLNode append_child(T type_ = node_element) {
            return append_child_impl(static_cast<int>(type_));
        }
        template<typename T = XMLNodeType>
        XMLNode prepend_child(T type_ = node_element) {
            return prepend_child_impl(static_cast<int>(type_));
        }

        XMLNode append_child(const char* name_, bool force_ns = false);
        XMLNode prepend_child(const char* name_, bool force_ns = false);

        template<typename T = XMLNodeType>
        XMLNode insert_child_after(T type_, const XMLNode& node) {
            return insert_child_after_impl(static_cast<int>(type_), node);
        }
        template<typename T = XMLNodeType>
        XMLNode insert_child_before(T type_, const XMLNode& node) {
            return insert_child_before_impl(static_cast<int>(type_), node);
        }

        XMLNode insert_child_after(const char* name_, const XMLNode& node, bool force_ns = false);
        XMLNode insert_child_before(const char* name_, const XMLNode& node, bool force_ns = false);

        bool remove_child(const char* name_);
        bool remove_child(const XMLNode& n);
        void remove_children();

        XMLNode find_child_by_attribute(const char* name_, const char* attr_name, const char* attr_value) const;
        XMLNode find_child_by_attribute(const char* attr_name, const char* attr_value) const;

        template<typename Predicate>
        XMLNode find_child(Predicate pred) const {
            return find_child_impl([&pred](pugi::xml_node_struct* n) {
                return pred(XMLNode(n));
            });
        }

        pugi::xpath_node select_node(const char* query) const;
        pugi::xpath_node_set select_nodes(const char* query) const;

        XMLText text() const;
        XMLNodeType type() const;

        template<typename T = XMLNodeType>
        XMLNode first_child_of_type(T type_ = node_element) const {
            return first_child_of_type_impl(static_cast<int>(type_));
        }
        template<typename T = XMLNodeType>
        XMLNode last_child_of_type(T type_ = node_element) const {
            return last_child_of_type_impl(static_cast<int>(type_));
        }
        template<typename T = XMLNodeType>
        size_t child_count_of_type(T type_ = node_element) const {
            return child_count_of_type_impl(static_cast<int>(type_));
        }
        template<typename T = XMLNodeType>
        XMLNode next_sibling_of_type(T type_ = node_element) const {
            return next_sibling_of_type_impl(static_cast<int>(type_));
        }
        template<typename T = XMLNodeType>
        XMLNode previous_sibling_of_type(T type_ = node_element) const {
            return previous_sibling_of_type_impl(static_cast<int>(type_));
        }
        template<typename T = XMLNodeType>
        XMLNode next_sibling_of_type(const char* name_, T type_ = node_element) const {
            return next_sibling_of_type_name_impl(name_, static_cast<int>(type_));
        }
        template<typename T = XMLNodeType>
        XMLNode previous_sibling_of_type(const char* name_, T type_ = node_element) const {
            return previous_sibling_of_type_name_impl(name_, static_cast<int>(type_));
        }

        // Copy and move methods (analogous to pugixml)
        XMLNode append_copy(const XMLNode& node);
        XMLNode prepend_copy(const XMLNode& node);
        XMLNode insert_copy_after(const XMLNode& node, const XMLNode& ref);
        XMLNode insert_copy_before(const XMLNode& node, const XMLNode& ref);

        XMLNode append_move(const XMLNode& node);
        XMLNode prepend_move(const XMLNode& node);
        XMLNode insert_move_after(const XMLNode& node, const XMLNode& ref);
        XMLNode insert_move_before(const XMLNode& node, const XMLNode& ref);

        XMLAttribute append_copy(const XMLAttribute& attr);

        // Conversion operator to pugi::xml_node (implemented in cpp)
        operator pugi::xml_node() const;

        // Iteration ranges
        XMLNodeRange children() const;
        XMLNodeRange children(const char* name_) const;
        XMLAttributeRange attributes() const;

        void print(std::ostream& os, const char* indent = "\t", unsigned int flags = 1 /* format_default */) const;

        const char* name_without_namespace(const char* name_) const;
        NameProxy namespaced_name_proxy(const char* name_, bool force_ns) const;

        pugi::xml_node_struct* internal_object() const { return m_node; }

    private:
        XMLNode append_child_impl(int type_);
        XMLNode prepend_child_impl(int type_);
        XMLNode insert_child_after_impl(int type_, const XMLNode& node);
        XMLNode insert_child_before_impl(int type_, const XMLNode& node);
        XMLNode first_child_of_type_impl(int type_) const;
        XMLNode last_child_of_type_impl(int type_) const;
        size_t child_count_of_type_impl(int type_) const;
        XMLNode next_sibling_of_type_impl(int type_) const;
        XMLNode previous_sibling_of_type_impl(int type_) const;
        XMLNode next_sibling_of_type_name_impl(const char* name_, int type_) const;
        XMLNode previous_sibling_of_type_name_impl(const char* name_, int type_) const;

        // Non-templated helper for find_child
        XMLNode find_child_impl(std::function<bool(pugi::xml_node_struct*)> pred) const;

    protected:
        pugi::xml_node_struct* m_node{nullptr};
        size_t name_begin{0};
    };

    // ===== XMLAttribute wrapper class =====
    class OPENXLSX_EXPORT XMLAttribute
    {
    public:
        XMLAttribute();
        XMLAttribute(pugi::xml_attribute_struct* attr);
        XMLAttribute(const pugi::xml_attribute& attr); // Implicit conversion constructor

        bool operator==(const XMLAttribute& other) const;
        bool operator!=(const XMLAttribute& other) const;
        explicit operator bool() const;
        bool empty() const;

        const char* name() const;
        const char* value() const;
        bool set_name(const char* rhs);
        bool set_value(const char* rhs);
        bool set_value(int rhs);
        bool set_value(unsigned int rhs);
        bool set_value(long rhs);
        bool set_value(unsigned long rhs);
        bool set_value(double rhs);
        bool set_value(bool rhs);
        bool set_value(long long rhs);
        bool set_value(unsigned long long rhs);

        // Assignment overloads
        XMLAttribute& operator=(const char* rhs) { set_value(rhs); return *this; }
        XMLAttribute& operator=(int rhs) { set_value(rhs); return *this; }
        XMLAttribute& operator=(unsigned int rhs) { set_value(rhs); return *this; }
        XMLAttribute& operator=(long rhs) { set_value(rhs); return *this; }
        XMLAttribute& operator=(unsigned long rhs) { set_value(rhs); return *this; }
        XMLAttribute& operator=(double rhs) { set_value(rhs); return *this; }
        XMLAttribute& operator=(bool rhs) { set_value(rhs); return *this; }
        XMLAttribute& operator=(long long rhs) { set_value(rhs); return *this; }
        XMLAttribute& operator=(unsigned long long rhs) { set_value(rhs); return *this; }

        const char* as_string(const char* def = "") const;
        int as_int(int def = 0) const;
        unsigned int as_uint(unsigned int def = 0) const;
        double as_double(double def = 0) const;
        float as_float(float def = 0.0f) const;
        bool as_bool(bool def = false) const;
        unsigned long long as_ullong(unsigned long long def = 0) const;
        long long as_llong(long long def = 0) const;

        XMLAttribute next_attribute() const;
        XMLAttribute previous_attribute() const;

        operator pugi::xml_attribute() const; // Conversion operator

        pugi::xml_attribute_struct* internal_object() const { return m_attr; }

    private:
        pugi::xml_attribute_struct* m_attr{nullptr};
    };

    // ===== XMLText wrapper class =====
    class OPENXLSX_EXPORT XMLText
    {
    public:
        XMLText();
        XMLText(pugi::xml_node_struct* node);

        bool empty() const;
        const char* get() const;

        const char* as_string(const char* def = "") const;
        int as_int(int def = 0) const;
        unsigned int as_uint(unsigned int def = 0) const;
        double as_double(double def = 0) const;
        float as_float(float def = 0.0f) const;
        bool as_bool(bool def = false) const;
        unsigned long long as_ullong(unsigned long long def = 0) const;
        long long as_llong(long long def = 0) const;

        bool set(const char* rhs);
        bool set(int rhs);
        bool set(unsigned int rhs);
        bool set(long rhs);
        bool set(unsigned long rhs);
        bool set(double rhs);
        bool set(bool rhs);
        bool set(long long rhs);
        bool set(unsigned long long rhs);

    private:
        pugi::xml_node_struct* m_node{nullptr};
    };

    // ===== XMLNodeIterator for children range loop =====
    class OPENXLSX_EXPORT XMLNodeIterator
    {
    public:
        using difference_type = std::ptrdiff_t;
        using value_type = XMLNode;
        using pointer = XMLNode*;
        using reference = XMLNode&; // Returns reference to internal wrap
        using iterator_category = std::forward_iterator_tag;

        XMLNodeIterator();
        XMLNodeIterator(pugi::xml_node_struct* node, pugi::xml_node_struct* parent, const char* filter = nullptr);

        XMLNode& operator*() const;
        XMLNode* operator->() const { return &m_wrap; }
        XMLNodeIterator& operator++();
        bool operator==(const XMLNodeIterator& other) const;
        bool operator!=(const XMLNodeIterator& other) const;

    private:
        pugi::xml_node_struct* m_node{nullptr};
        pugi::xml_node_struct* m_parent{nullptr};
        const char* m_filter{nullptr};
        mutable XMLNode m_wrap;
    };

    class OPENXLSX_EXPORT XMLNodeRange
    {
    public:
        XMLNodeRange(pugi::xml_node_struct* first, pugi::xml_node_struct* parent, const char* filter = nullptr);
        XMLNodeIterator begin() const;
        XMLNodeIterator end() const;
    private:
        pugi::xml_node_struct* m_first{nullptr};
        pugi::xml_node_struct* m_parent{nullptr};
        const char* m_filter{nullptr};
    };

    // ===== XMLAttributeIterator for attributes range loop =====
    class OPENXLSX_EXPORT XMLAttributeIterator
    {
    public:
        using difference_type = std::ptrdiff_t;
        using value_type = XMLAttribute;
        using pointer = XMLAttribute*;
        using reference = XMLAttribute&; // Returns reference to internal wrap
        using iterator_category = std::forward_iterator_tag;

        XMLAttributeIterator();
        XMLAttributeIterator(pugi::xml_attribute_struct* attr, pugi::xml_node_struct* parent);

        XMLAttribute& operator*() const;
        XMLAttribute* operator->() const { return &m_wrap; }
        XMLAttributeIterator& operator++();
        bool operator==(const XMLAttributeIterator& other) const;
        bool operator!=(const XMLAttributeIterator& other) const;

    private:
        pugi::xml_attribute_struct* m_attr{nullptr};
        pugi::xml_node_struct* m_parent{nullptr};
        mutable XMLAttribute m_wrap;
    };

    class OPENXLSX_EXPORT XMLAttributeRange
    {
    public:
        XMLAttributeRange(pugi::xml_attribute_struct* first, pugi::xml_node_struct* parent);
        XMLAttributeIterator begin() const;
        XMLAttributeIterator end() const;
    private:
        pugi::xml_attribute_struct* m_first{nullptr};
        pugi::xml_node_struct* m_parent{nullptr};
    };

    // ===== Custom OpenXLSX_xml_document wrapper class =====
    class OPENXLSX_EXPORT OpenXLSX_xml_document : public XMLNode
    {
    public:
        OpenXLSX_xml_document();
        ~OpenXLSX_xml_document();

        OpenXLSX_xml_document(const OpenXLSX_xml_document&) = delete;
        OpenXLSX_xml_document& operator=(const OpenXLSX_xml_document&) = delete;

        OpenXLSX_xml_document(OpenXLSX_xml_document&& other) noexcept;
        OpenXLSX_xml_document& operator=(OpenXLSX_xml_document&& other) noexcept;

        XMLNode document_element() const;
        void print(std::ostream& ostr) const;

        XMLParseResult load_string(const char* contents, unsigned int options = 0x0016 /* parse_default */);
        XMLParseResult load_buffer(const void* contents, size_t size, unsigned int settings = 0x0016 /* parse_default */);

        void save(pugi::xml_writer& writer, const char* indent = "\t", unsigned int flags = 1 /* format_default */) const;
        void save(std::ostream& stream, const char* indent = "\t", unsigned int flags = 1 /* format_default */) const;

        void reset();
        void reset(const XMLDocument& proto);

        operator pugi::xml_document&();
        operator const pugi::xml_document&() const;

        pugi::xml_document& internal_document() { return *m_doc; }
        const pugi::xml_document& internal_document() const { return *m_doc; }

    private:
        std::unique_ptr<pugi::xml_document> m_doc;
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLXMLPARSER_HPP

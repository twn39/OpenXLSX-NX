// ===== External Includes ===== //
#include <cstdlib>
#include <cstring>
#include <pugixml.hpp>
#include <sstream>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLXmlData.hpp"

namespace
{
    // Fast string writer to eliminate string stream intermediate allocations
    struct StringXmlWriter : public pugi::xml_writer {
        std::string* result;
        StringXmlWriter(std::string* str) : result(str) {}
        void write(const void* data, size_t size) override {
            result->append(static_cast<const char*>(data), size);
        }
    };

    // Zero-copy malloc writer for libzip
    struct MallocXmlWriter : public pugi::xml_writer {
        void*  buffer   = nullptr;
        size_t size     = 0;
        size_t capacity = 0;

        MallocXmlWriter() {
            capacity = 1048576; // Start with 1MB to avoid early fragmentation
            buffer = std::malloc(capacity);
            if (!buffer) throw std::bad_alloc();
        }

        ~MallocXmlWriter() {
            if (buffer) std::free(buffer);
        }

        void write(const void* data, size_t len) override {
            if (size + len > capacity) {
                size_t new_capacity = capacity;
                while (new_capacity < size + len) new_capacity *= 2;
                void* new_buffer = std::realloc(buffer, new_capacity);
                if (!new_buffer) throw std::bad_alloc();
                buffer = new_buffer;
                capacity = new_capacity;
            }
            std::memcpy(static_cast<char*>(buffer) + size, data, len);
            size += len;
        }

        void* release() {
            void* ptr = buffer;
            buffer = nullptr;
            return ptr;
        }
    };
}

using namespace OpenXLSX;

/**
 * @details
 */
XLXmlData::XLXmlData(XLDocument* parentDoc, const std::string& xmlPath, const std::string& xmlId, XLContentType xmlType)
    : m_parentDoc(parentDoc),
      m_xmlPath(xmlPath),
      m_xmlID(xmlId),
      m_xmlType(xmlType),
      m_xmlDoc(std::make_unique<XMLDocument>())
{ m_xmlDoc->reset(); }

/**
 * @details
 */
XLXmlData::~XLXmlData() = default;

/**
 * @details
 */
void XLXmlData::setRawData(const std::string& data)    // NOLINT
{
    auto result = m_xmlDoc->load_buffer(data.data(), data.size(), pugi_parse_settings);
    if (!result && result.status != pugi::status_no_document_element) {
        throw XLException("Failed to parse raw XML data. Error: " + std::string(result.description()));
    }
}

/**
 * @details
 * @note Default encoding for pugixml xml_document::save is pugi::encoding_auto, becomes pugi::encoding_utf8
 */
std::string XLXmlData::getRawData(XLXmlSavingDeclaration savingDeclaration) const
{
    XMLDocument* doc = const_cast<XMLDocument*>(getXmlDocument());

    // ===== 2024-07-08: ensure that the default encoding UTF-8 is explicitly written to the XML document with a custom saving declaration
    XMLNode saveDeclaration;
    for (auto node : doc->children()) {
        if (node.type() == pugi::node_declaration) {
            saveDeclaration = node;
            break;
        }
    }

    if (saveDeclaration.empty()) {                                       // if saving declaration node does not exist
        doc->prepend_child(pugi::node_pcdata).set_value("\n");           // prepend a line break
        saveDeclaration = doc->prepend_child(pugi::node_declaration);    // prepend a saving declaration
    }

    // ===== If a node_declaration could be fetched or created
    if (not saveDeclaration.empty()) {
        // ===== Fetch or create saving declaration attributes
        XMLAttribute attrVersion = saveDeclaration.attribute("version");
        if (attrVersion.empty()) attrVersion = saveDeclaration.append_attribute("version");
        XMLAttribute attrEncoding = saveDeclaration.attribute("encoding");
        if (attrEncoding.empty()) attrEncoding = saveDeclaration.append_attribute("encoding");
        XMLAttribute attrStandalone = saveDeclaration.attribute("standalone");
        if (savingDeclaration.standalone_as_bool()) {
            if (attrStandalone.empty()) attrStandalone = saveDeclaration.append_attribute("standalone");
            attrStandalone.set_value("yes");
        }
        else if (!attrStandalone.empty()) {
            saveDeclaration.remove_attribute(attrStandalone);
        }

        // ===== Set saving declaration attribute values (potentially overwriting existing values)
        attrVersion  = savingDeclaration.version().c_str();     // version="1.0" is XML default
        attrEncoding = savingDeclaration.encoding().c_str();    // encoding="UTF-8" is XML default
    }

    std::string str;
    str.reserve(8192); // Pre-allocate some capacity
    StringXmlWriter writer(&str);
    doc->save(writer, "", pugi::format_default);
    return str;
}

/**
 * @details
 */
XLAllocatedMemory XLXmlData::getRawAllocatedData(XLXmlSavingDeclaration savingDeclaration) const
{
    XMLDocument* doc = const_cast<XMLDocument*>(getXmlDocument());

    // ===== 2024-07-08: ensure that the default encoding UTF-8 is explicitly written to the XML document with a custom saving declaration
    XMLNode saveDeclaration;
    for (auto node : doc->children()) {
        if (node.type() == pugi::node_declaration) {
            saveDeclaration = node;
            break;
        }
    }

    if (saveDeclaration.empty()) {                                       // if saving declaration node does not exist
        doc->prepend_child(pugi::node_pcdata).set_value("\n");           // prepend a line break
        saveDeclaration = doc->prepend_child(pugi::node_declaration);    // prepend a saving declaration
    }

    // ===== If a node_declaration could be fetched or created
    if (not saveDeclaration.empty()) {
        // ===== Fetch or create saving declaration attributes
        XMLAttribute attrVersion = saveDeclaration.attribute("version");
        if (attrVersion.empty()) attrVersion = saveDeclaration.append_attribute("version");
        XMLAttribute attrEncoding = saveDeclaration.attribute("encoding");
        if (attrEncoding.empty()) attrEncoding = saveDeclaration.append_attribute("encoding");
        XMLAttribute attrStandalone = saveDeclaration.attribute("standalone");
        if (savingDeclaration.standalone_as_bool()) {
            if (attrStandalone.empty()) attrStandalone = saveDeclaration.append_attribute("standalone");
            attrStandalone.set_value("yes");
        }
        else if (!attrStandalone.empty()) {
            saveDeclaration.remove_attribute(attrStandalone);
        }

        // ===== Set saving declaration attribute values (potentially overwriting existing values)
        attrVersion  = savingDeclaration.version().c_str();     // version="1.0" is XML default
        attrEncoding = savingDeclaration.encoding().c_str();    // encoding="UTF-8" is XML default
    }

    MallocXmlWriter writer;
    doc->save(writer, "", pugi::format_default);
    
    XLAllocatedMemory mem;
    mem.data = writer.release();
    mem.size = writer.size;
    return mem;
}

/**
 * @details
 */
XLDocument* XLXmlData::getParentDoc() { return m_parentDoc; }

/**
 * @details
 */
const XLDocument* XLXmlData::getParentDoc() const { return m_parentDoc; }

/**
 * @details
 */
std::string XLXmlData::getXmlPath() const { return m_xmlPath; }

/**
 * @details
 */
std::string XLXmlData::getXmlID() const { return m_xmlID; }

/**
 * @details
 */
XLContentType XLXmlData::getXmlType() const { return m_xmlType; }

/**
 * @details
 */
XMLDocument* XLXmlData::getXmlDocument()
{
    if (!m_xmlDoc->document_element()) {
        std::string xmlContent = m_parentDoc->extractXmlFromArchive(m_xmlPath);
        if (!xmlContent.empty()) {
            auto result = m_xmlDoc->load_buffer(xmlContent.data(), xmlContent.size(), pugi_parse_settings);
            if (!result && result.status != pugi::status_no_document_element) {
                throw XLException("Failed to parse XML: " + m_xmlPath + ", Error: " + result.description());
            }
        }
    }

    return m_xmlDoc.get();
}

/**
 * @details
 */
const XMLDocument* XLXmlData::getXmlDocument() const
{
    if (!m_xmlDoc->document_element()) {
        std::string xmlContent = m_parentDoc->extractXmlFromArchive(m_xmlPath);
        if (!xmlContent.empty()) {
            auto result = m_xmlDoc->load_buffer(xmlContent.data(), xmlContent.size(), pugi_parse_settings);
            if (!result && result.status != pugi::status_no_document_element) {
                throw XLException("Failed to parse XML: " + m_xmlPath + ", Error: " + result.description());
            }
        }
    }

    return m_xmlDoc.get();
}

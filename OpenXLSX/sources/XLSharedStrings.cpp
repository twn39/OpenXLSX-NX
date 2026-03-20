// ===== External Includes ===== //
#include <algorithm>
#include <gsl/gsl>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLSharedStrings.hpp"

#include "XLException.hpp"

namespace OpenXLSX
{
    const XLSharedStrings XLSharedStringsDefaulted{};
}    // namespace OpenXLSX

using namespace OpenXLSX;

/**
 * @details Constructs a new XLSharedStrings object. Only one (common) object is allowed per XLDocument instance.
 * A filepath to the underlying XML file must be provided.
 */
XLSharedStrings::XLSharedStrings(XLXmlData*                                xmlData,
                                 std::deque<std::string>*                  stringCache,
                                 std::unordered_map<std::string, int32_t>* stringIndex)
    : XLXmlFile(xmlData),
      m_stringCache(stringCache),
      m_stringIndex(stringIndex)
{
    XMLDocument& doc = xmlDocument();
    if (doc.document_element().empty())    // handle a bad (no document element) xl/sharedStrings.xml
        doc.load_string(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"    // pull request #192 -> remove count &
                                                                                             // uniqueCount as they are optional
            // 2024-09-03: removed empty string entry on creation, as it appears to just waste a string index that will never be used
            // "  <si>\n"
            // "    <t/>\n"
            // "  </si>\n"
            "</sst>",
            pugi_parse_settings);

    // Build the hash index from the string cache for O(1) lookup
    if (m_stringIndex and m_stringCache) {
        m_stringIndex->clear();
        m_stringIndex->reserve(m_stringCache->size());
        int32_t idx = 0;
        for (const auto& str : *m_stringCache) { m_stringIndex->emplace(str, idx++); }
    }
}

/**
 * @details
 */
XLSharedStrings::~XLSharedStrings() = default;

/**
 * @details Look up a string index by the string content. If the string does not exist, the returned index is -1.
 * Optimized to use O(1) hash lookup when available.
 */
int32_t XLSharedStrings::getStringIndex(const std::string& str) const
{
    Expects(m_stringIndex != nullptr || m_stringCache != nullptr);
    // Use O(1) hash lookup if available
    if (m_stringIndex) {
        auto it = m_stringIndex->find(str);
        return it != m_stringIndex->end() ? it->second : -1;
    }

    // Fallback to linear search (legacy behavior)
    const auto iter = std::find_if(m_stringCache->begin(), m_stringCache->end(), [&](const std::string& s) { return str == s; });
    return iter == m_stringCache->end() ? -1 : static_cast<int32_t>(std::distance(m_stringCache->begin(), iter));
}

/**
 * @details Check if a string exists in the shared strings table. O(1) with hash index.
 */
bool XLSharedStrings::stringExists(const std::string& str) const
{
    // Use O(1) hash lookup if available
    if (m_stringIndex) { return m_stringIndex->find(str) != m_stringIndex->end(); }
    return getStringIndex(str) >= 0;
}

/**
 * @details
 */
const char* XLSharedStrings::getString(int32_t index) const
{
    Expects(m_stringCache != nullptr);
    if (index < 0 or static_cast<size_t>(index) >= m_stringCache->size()) {    // 2024-04-30: added range check
        using namespace std::literals::string_literals;
        throw XLInternalError("XLSharedStrings::"s + __func__ + ": index "s + std::to_string(index) + " is out of range"s);
    }
    return (*m_stringCache)[static_cast<size_t>(index)].c_str();
}

/**
 * @details Append a string by creating a new node in the XML file and adding the string to it. The index to the
 * shared string is returned
 */
int32_t XLSharedStrings::appendString(const std::string& str) const
{
    Expects(m_stringCache != nullptr);
    size_t stringCacheSize = m_stringCache->size();    // 2024-05-31: analogous with already added range check in getString
    if (stringCacheSize >= XLMaxSharedStrings) {       // 2024-05-31: added range check
        using namespace std::literals::string_literals;
        throw XLInternalError("XLSharedStrings::"s + __func__ + ": exceeded max strings count "s + std::to_string(XLMaxSharedStrings));
    }
    auto textNode = xmlDocument().document_element().append_child("si").append_child("t");
    if ((!str.empty()) and (str.front() == ' ' or str.back() == ' '))
        textNode.append_attribute("xml:space").set_value("preserve");    // pull request #161
    textNode.text().set(str.c_str());
    m_stringCache->emplace_back(textNode.text().get());    // index of this element = previous stringCacheSize

    // Update the hash index for O(1) lookup
    if (m_stringIndex) { m_stringIndex->emplace(str, static_cast<int32_t>(stringCacheSize)); }

    return static_cast<int32_t>(stringCacheSize);
}

/**
 * @details Get or create a string index in O(1) time. This is the optimized path for setting cell values.
 * It avoids the separate stringExists() + getStringIndex()/appendString() pattern.
 */
int32_t XLSharedStrings::getOrCreateStringIndex(const std::string& str) const
{
    // Fast path: O(1) lookup using hash index
    if (m_stringIndex) {
        auto it = m_stringIndex->find(str);
        if (it != m_stringIndex->end()) {
            return it->second;    // String already exists
        }
    }

    // String doesn't exist, append it
    return appendString(str);
}

/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLSharedStrings::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print(ostr); }

/**
 * @details Clear the string at the given index. This will affect the entire spreadsheet; everywhere the shared string
 * is used, it will be erased.
 * @note: 2024-05-31 DONE: index now int32_t everywhere, 2 billion shared strings should be plenty
 */
void XLSharedStrings::clearString(int32_t index) const    // 2024-04-30: whitespace support
{
    Expects(m_stringCache != nullptr);
    if (index < 0 or static_cast<size_t>(index) >= m_stringCache->size()) {    // 2024-04-30: added range check
        using namespace std::literals::string_literals;
        throw XLInternalError("XLSharedStrings::"s + __func__ + ": index "s + std::to_string(index) + " is out of range"s);
    }

    // Keep m_stringIndex in sync to prevent dangling references
    if (m_stringIndex) {
        auto oldStr = (*m_stringCache)[static_cast<size_t>(index)];
        if (!oldStr.empty()) {
            m_stringIndex->erase(oldStr);
        }
    }

    (*m_stringCache)[static_cast<size_t>(index)] = "";
    // auto iter            = xmlDocument().document_element().children().begin();
    // std::advance(iter, index);
    // iter->text().set(""); // 2024-04-30: BUGFIX: this was never going to work, <si> entries can be plenty that need to be cleared,
    // including formatting

    /* 2024-04-30 CAUTION: performance critical - with whitespace support, the function can no longer know the exact iterator position of
     *   the shared string to be cleared - TBD what to do instead?
     * Potential solution: store the XML child position with each entry in m_stringCache in a std::deque<struct entry>
     *   with struct entry { std::string s; uint64_t xmlChildIndex; };
     */
    XMLNode sharedStringNode = xmlDocument().document_element().first_child_of_type(pugi::node_element);
    int32_t sharedStringPos  = 0;
    while (sharedStringPos < index and not sharedStringNode.empty()) {
        sharedStringNode = sharedStringNode.next_sibling_of_type(pugi::node_element);
        ++sharedStringPos;
    }
    if (not sharedStringNode.empty()) {        // index was found
        sharedStringNode.remove_children();    // clear all data and formatting
        sharedStringNode.append_child("t");    // append an empty text node
    }
}

/**
 * @details
 */
int32_t XLSharedStrings::rewriteXmlFromCache()
{
    Expects(m_stringCache != nullptr);
    int32_t writtenStrings = 0;
    xmlDocument().document_element().remove_children();    // clear all existing XML
    for (std::string& s : *m_stringCache) {
        XMLNode textNode = xmlDocument().document_element().append_child("si").append_child("t");
        if ((!s.empty()) and (s.front() == ' ' or s.back() == ' '))
            textNode.append_attribute("xml:space").set_value("preserve");    // preserve spaces at begin/end of string
        textNode.text().set(s.c_str());
        ++writtenStrings;
    }
    return writtenStrings;
}

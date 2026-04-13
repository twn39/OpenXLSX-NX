// ===== External Includes ===== //
#include <algorithm>
#include <gsl/gsl>
#include <mutex>
#include <pugixml.hpp>
#include <shared_mutex>

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
XLSharedStrings::XLSharedStrings(XLXmlData*            xmlData,
                                 XLSharedStringsState* state)
    : XLXmlFile(xmlData),
      m_state(state)
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
    if (m_state) {
        m_state->index.clear();
        m_state->index.reserve(m_state->cache.size());
        int32_t idx = 0;
        for (const auto& str : m_state->cache) { m_state->index.emplace(str, idx++); }
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
int32_t XLSharedStrings::getStringIndex(std::string_view str) const
{
    Expects(m_state != nullptr);

    std::shared_lock<std::shared_mutex> lock;
    if (m_state->mutex) lock = std::shared_lock<std::shared_mutex>(*m_state->mutex);

    // Use O(1) hash lookup
    auto it = m_state->index.find(str);
    return it != m_state->index.end() ? it->second : -1;
}

/**
 * @details Check if a string exists in the shared strings table. O(1) with hash index.
 */
bool XLSharedStrings::stringExists(std::string_view str) const
{
    if (!m_state) return false;

    std::shared_lock<std::shared_mutex> lock;
    if (m_state->mutex) lock = std::shared_lock<std::shared_mutex>(*m_state->mutex);

    // Use O(1) hash lookup
    return m_state->index.find(str) != m_state->index.end();
}

/**
 * @details
 */
const char* XLSharedStrings::getString(int32_t index) const
{
    Expects(m_state != nullptr);

    std::shared_lock<std::shared_mutex> lock;
    if (m_state->mutex) lock = std::shared_lock<std::shared_mutex>(*m_state->mutex);

    if (index < 0 or static_cast<size_t>(index) >= m_state->cache.size()) {    // 2024-04-30: added range check
        using namespace std::literals::string_literals;
        throw XLInternalError("XLSharedStrings::"s + __func__ + ": index "s + std::to_string(index) + " is out of range"s);
    }
    return m_state->cache[static_cast<size_t>(index)].data();
}

/**
 * @details Append a string by creating a new node in the XML file and adding the string to it. The index to the
 * shared string is returned
 */
int32_t XLSharedStrings::appendString(std::string_view str) const
{
    if (!isCleanXmlString(str)) {
        return appendString(sanitizeXmlString(str));
    }

    Expects(m_state != nullptr);

    std::unique_lock<std::shared_mutex> lock;
    if (m_state->mutex) {
        lock = std::unique_lock<std::shared_mutex>(*m_state->mutex);
        // Double check locking
        auto it = m_state->index.find(str);
        if (it != m_state->index.end()) { return it->second; }
    }

    size_t stringCacheSize = m_state->cache.size();    // 2024-05-31: analogous with already added range check in getString
    if (stringCacheSize >= XLMaxSharedStrings) {       // 2024-05-31: added range check
        using namespace std::literals::string_literals;
        throw XLInternalError("XLSharedStrings::"s + __func__ + ": exceeded max strings count "s + std::to_string(XLMaxSharedStrings));
    }
    // Lazy DOM path: store the string in the arena and cache only, then mark the DOM
    // as out-of-sync.  The full pugi DOM is rebuilt from the cache in
    // rewriteXmlFromCache() at save time.  Skipping DOM mutation here prevents a
    // potentially 10M-node tree from growing in RAM during large write sessions.
    std::string_view persistentView = m_state->arena.store(str);
    m_state->cache.emplace_back(persistentView);

    m_state->index.emplace(persistentView, static_cast<int32_t>(stringCacheSize));

    // Signal that the pugi DOM no longer reflects the full cache
    m_domDirty = true;

    return static_cast<int32_t>(stringCacheSize);
}

/**
 * @details Get or create a string index in O(1) time. This is the optimized path for setting cell values.
 * It avoids the separate stringExists() + getStringIndex()/appendString() pattern.
 */
int32_t XLSharedStrings::getOrCreateStringIndex(std::string_view str) const
{
    std::string cleanStr;
    if (!isCleanXmlString(str)) {
        cleanStr = sanitizeXmlString(str);
        str = cleanStr; // Point the view to the clean string
    }

    if (m_state && m_state->mutex) {
        std::shared_lock<std::shared_mutex> lock(*m_state->mutex);
        auto it = m_state->index.find(str);
        if (it != m_state->index.end()) {
            return it->second;    // String already exists
        }
    }
    else if (m_state) {
        // Fast path: O(1) lookup using hash index
        auto it = m_state->index.find(str);
        if (it != m_state->index.end()) {
            return it->second;    // String already exists
        }
    }

    return appendString(str);
}

/**
 * @details Pre-reserve capacity in the string cache vector and hash index for n strings.
 * Avoids repeated incremental reallocations when bulk-inserting many strings.
 */
void XLSharedStrings::reserveStrings(size_t n) const
{
    if (m_state) {
        m_state->cache.reserve(n);
        m_state->index.reserve(n);
    }
}

/**
 * @details Returns the approximate heap memory used by the string cache vector and
 * the hash index buckets.  Does not include the arena itself.
 */
size_t XLSharedStrings::memoryUsageBytes() const noexcept
{
    size_t total = 0;
    if (m_state) {
        total += m_state->cache.capacity() * sizeof(std::string_view);
        total += m_state->index.bucket_count() * (sizeof(void*) + sizeof(std::pair<std::string_view, int32_t>));
    }
    return total;
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
    Expects(m_state != nullptr);

    std::unique_lock<std::shared_mutex> lock;
    if (m_state->mutex) lock = std::unique_lock<std::shared_mutex>(*m_state->mutex);

    if (index < 0 or static_cast<size_t>(index) >= m_state->cache.size()) {    // 2024-04-30: added range check
        using namespace std::literals::string_literals;
        throw XLInternalError("XLSharedStrings::"s + __func__ + ": index "s + std::to_string(index) + " is out of range"s);
    }

    // Keep m_state->index in sync to prevent dangling references
    auto oldStr = m_state->cache[static_cast<size_t>(index)];
    if (!oldStr.empty()) { m_state->index.erase(oldStr); }

    m_state->cache[static_cast<size_t>(index)] = "";
    // auto iter            = xmlDocument().document_element().children().begin();
    // std::advance(iter, index);
    // iter->text().set(""); // 2024-04-30: BUGFIX: this was never going to work, <si> entries can be plenty that need to be cleared,
    // including formatting

    // Clear the node directly by iterating to its position
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
    Expects(m_state != nullptr);
    int32_t writtenStrings = 0;
    xmlDocument().document_element().remove_children();    // clear all existing XML
    for (std::string_view s : m_state->cache) {
        XMLNode textNode = xmlDocument().document_element().append_child("si").append_child("t");
        if ((!s.empty()) and (s.front() == ' ' or s.back() == ' '))
            textNode.append_attribute("xml:space").set_value("preserve");    // preserve spaces at begin/end of string
        textNode.text().set(s.data());                                       // s is guaranteed to be null-terminated by the Arena
        ++writtenStrings;
    }
    return writtenStrings;
}

void XLSharedStrings::rebuild(const std::vector<int32_t>& indexMap, int32_t newStringCount)
{
    Expects(m_state != nullptr);
    const size_t oldStringCount = m_state->cache.size();

    XLStringArena                 newArena;
    std::vector<std::string_view> newStringCache(static_cast<size_t>(newStringCount));
    newStringCache[0] = newArena.store("");
    for (size_t oldIdx = 0; oldIdx < oldStringCount; ++oldIdx) {
        if (const int32_t newIdx = indexMap[oldIdx]; newIdx > 0)
            newStringCache[static_cast<size_t>(newIdx)] = newArena.store(m_state->cache[oldIdx]);
    }

    m_state->arena = std::move(newArena);
    m_state->cache = std::move(newStringCache);
    
    // rebuilding index
    m_state->index.clear();
    for (size_t i = 0; i < m_state->cache.size(); ++i) {
        m_state->index.emplace(m_state->cache[i], static_cast<int32_t>(i));
    }

    if (static_cast<int32_t>(m_state->cache.size()) != rewriteXmlFromCache())
        throw XLInternalError("XLSharedStrings::rebuild: failed to rewrite shared string table - document would be corrupted");
}

// ===== External Includes ===== //
#include <algorithm>
#include <iostream>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCellReference.hpp"
#include "XLException.hpp"
#include "XLMergeCells.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

namespace
{
    /**
     * @details
     */
    XLMergeCells::XLRect parseRange(std::string_view reference)
    {
        using namespace std::literals::string_literals;

        size_t pos = reference.find_first_of(':');
        if (pos == std::string_view::npos or pos < 2 or pos + 2 >= reference.length())
            throw XLInputError("XLMergeCells: not a valid range reference: \""s + std::string(reference) + "\""s);

        XLCellReference refTL(std::string(reference.substr(0, pos)));
        XLCellReference refBR(std::string(reference.substr(pos + 1)));

        uint32_t refTopRow    = refTL.row();
        uint16_t refFirstCol  = refTL.column();
        uint32_t refBottomRow = refBR.row();
        uint16_t refLastCol   = refBR.column();

        if (refBottomRow < refTopRow or refLastCol < refFirstCol or (refBottomRow == refTopRow and refLastCol == refFirstCol))
            throw XLInputError("XLMergeCells: not a valid range reference: \""s + std::string(reference) + "\""s);

        return {refTopRow, refFirstCol, refBottomRow, refLastCol};
    }
}    // namespace

/**
 * @details
 */
XLMergeCells::XLMergeCells() = default;

/**
 * @details
 */
XLMergeCells::XLMergeCells(const XMLNode& rootNode, std::vector<std::string_view> const& nodeOrder)
    : m_rootNode(rootNode),
      m_nodeOrder(nodeOrder),
      m_mergeCellsNode()
{
    if (m_rootNode.empty()) throw XLInternalError("XLMergeCells constructor: can not construct with an empty XML root node");

    m_mergeCellsNode  = m_rootNode.child("mergeCells");
    XMLNode mergeNode = m_mergeCellsNode.first_child_of_type(pugi::node_element);
    while (not mergeNode.empty()) {
        bool invalidNode = true;

        if (std::string(mergeNode.name()) == "mergeCell") {
            std::string_view ref = mergeNode.attribute("ref").value();
            if (not ref.empty()) {
                try {
                    // Populate cache during initialization to avoid repeated XML parsing.
                    m_mergeCache.push_back({std::string(ref), parseRange(ref)});
                    invalidNode = false;
                }
                catch (const XLInputError&) {
                    // Skip invalid nodes.
                }
            }
        }

        XMLNode nextNode = mergeNode.next_sibling_of_type(pugi::node_element);

        if (invalidNode) {
            std::cerr << "XLMergeCells constructor: invalid child element, either name is not mergeCell or reference is invalid:"
                      << std::endl;
            mergeNode.print(std::cerr);
            if (not nextNode.empty()) {
                // Cleanup whitespace before removing node to keep XML clean.
                while (mergeNode.next_sibling() != nextNode) m_mergeCellsNode.remove_child(mergeNode.next_sibling());
            }
            m_mergeCellsNode.remove_child(mergeNode);
        }

        mergeNode = nextNode;
    }

    if (not m_mergeCache.empty()) {
        XMLAttribute attr = m_mergeCellsNode.attribute("count");
        if (attr.empty()) attr = m_mergeCellsNode.append_attribute("count");
        attr.set_value(static_cast<unsigned long long>(m_mergeCache.size()));
    }
    else
        deleteAll();
}

/**
 * @details
 */
bool XLMergeCells::valid() const { return (not m_rootNode.empty()); }

/**
 * @details
 */
XLMergeIndex XLMergeCells::findMerge(std::string_view reference) const
{
    const auto iter =
        std::find_if(m_mergeCache.begin(), m_mergeCache.end(), [&](const XLMergeEntry& entry) { return reference == entry.reference; });

    return iter == m_mergeCache.end() ? XLMergeNotFound : static_cast<XLMergeIndex>(std::distance(m_mergeCache.begin(), iter));
}

/**
 * @details
 */
bool XLMergeCells::mergeExists(std::string_view reference) const { return findMerge(reference) >= 0; }

/**
 * @details
 */
XLMergeIndex XLMergeCells::findMergeByCell(std::string_view cellRef) const
{ return findMergeByCell(XLCellReference(std::string(cellRef))); }

XLMergeIndex XLMergeCells::findMergeByCell(XLCellReference cellRef) const
{
    const uint32_t row = cellRef.row();
    const uint16_t col = cellRef.column();

    // Use pre-parsed numerical ranges for high performance.
    const auto iter = std::find_if(m_mergeCache.begin(), m_mergeCache.end(), [row, col](const XLMergeEntry& entry) {
        return entry.rect.contains(row, col);
    });

    return iter == m_mergeCache.end() ? XLMergeNotFound : static_cast<XLMergeIndex>(std::distance(m_mergeCache.begin(), iter));
}

/**
 * @details
 */
size_t XLMergeCells::count() const { return m_mergeCache.size(); }

/**
 * @details
 */
const char* XLMergeCells::merge(XLMergeIndex index) const
{
    if (index < 0 or static_cast<size_t>(index) >= m_mergeCache.size()) {
        using namespace std::literals::string_literals;
        throw XLInputError("XLMergeCells::merge: index "s + std::to_string(index) + " is out of range"s);
    }
    return m_mergeCache[static_cast<size_t>(index)].reference.c_str();
}

/**
 * @details
 */
XLMergeIndex XLMergeCells::appendMerge(const std::string& reference)
{
    using namespace std::literals::string_literals;

    if (m_mergeCache.size() >= XLMaxMergeCells)
        throw XLInputError("XLMergeCells::appendMerge: exceeded max merge cells count "s + std::to_string(XLMaxMergeCells));

    const XLRect newRect = parseRange(reference);

    // Guard against overlaps using numerical range comparison.
    for (const auto& entry : m_mergeCache) {
        if (entry.rect.overlaps(newRect))
            throw XLInputError("XLMergeCells::appendMerge: reference \""s + reference + "\" overlaps with existing reference \""s +
                               entry.reference + "\""s);
    }

    if (m_mergeCellsNode.empty())
        m_mergeCellsNode = appendAndGetNode(m_rootNode, "mergeCells", m_nodeOrder);

    XMLNode insertAfter = m_mergeCellsNode.last_child_of_type(pugi::node_element);
    XMLNode newMerge{};
    if (insertAfter.empty())
        newMerge = m_mergeCellsNode.prepend_child("mergeCell");
    else
        newMerge = m_mergeCellsNode.insert_child_after("mergeCell", insertAfter);

    if (newMerge.empty()) throw XLInternalError("XLMergeCells::appendMerge: failed to insert reference: \""s + reference + "\""s);
    newMerge.append_attribute("ref").set_value(reference.c_str());

    m_mergeCache.push_back({reference, newRect});

    XMLAttribute attr = m_mergeCellsNode.attribute("count");
    if (attr.empty()) attr = m_mergeCellsNode.append_attribute("count");
    attr.set_value(static_cast<unsigned long long>(m_mergeCache.size()));

    return static_cast<XLMergeIndex>(m_mergeCache.size() - 1);
}

/**
 * @details
 */
void XLMergeCells::deleteMerge(XLMergeIndex index)
{
    using namespace std::literals::string_literals;

    if (index < 0 or static_cast<size_t>(index) >= m_mergeCache.size())
        throw XLInputError("XLMergeCells::deleteMerge: index "s + std::to_string(index) + " is out of range"s);

    XLMergeIndex curIndex = 0;
    XMLNode      node     = m_mergeCellsNode.first_child_of_type(pugi::node_element);
    while (curIndex < index and not node.empty()) {
        node = node.next_sibling_of_type(pugi::node_element);
        ++curIndex;
    }

    if (node.empty()) throw XLInternalError("XLMergeCells::deleteMerge: mismatch between size of mergeCells XML node and internal cache");

    while (node.previous_sibling().type() == pugi::node_pcdata) m_mergeCellsNode.remove_child(node.previous_sibling());
    m_mergeCellsNode.remove_child(node);

    m_mergeCache.erase(m_mergeCache.begin() + index);

    if (not m_mergeCache.empty()) {
        XMLAttribute attr = m_mergeCellsNode.attribute("count");
        if (attr.empty()) attr = m_mergeCellsNode.append_attribute("count");
        attr.set_value(static_cast<unsigned long long>(m_mergeCache.size()));
    }
    else
        deleteAll();
}

/**
 * @details
 */
void XLMergeCells::deleteAll()
{
    m_mergeCache.clear();
    m_rootNode.remove_child(m_mergeCellsNode);
    m_mergeCellsNode = XMLNode();
}

/**
 * @details Shift merge regions vertically.
 * - Regions entirely above fromRow  → unchanged.
 * - Regions entirely at/below fromRow → both top and bottom slide by delta.
 * - Regions straddling fromRow (delta < 0 / delete):
 *     top unchanged, bottom shrinks; if this makes height < 1 → deleteMerge.
 * - Regions straddling fromRow (delta > 0 / insert): not possible because
 *     insert never touches an existing row – we simply push everything down.
 */
void XLMergeCells::shiftRows(int32_t delta, uint32_t fromRow)
{
    if (delta == 0 || m_mergeCache.empty()) return;

    // Iterate in reverse so that deleteMerge() index invalidation is safe
    for (auto i = static_cast<int32_t>(m_mergeCache.size()) - 1; i >= 0; --i) {
        auto idx  = static_cast<XLMergeIndex>(i);
        auto& entry = m_mergeCache[static_cast<size_t>(i)];
        XLRect& r = entry.rect;

        if (r.bottom < fromRow) continue;  // entirely above → no change

        if (delta < 0) {
            // Deletion band: [fromRow, fromRow + |delta|)
            uint32_t delEnd = fromRow + static_cast<uint32_t>(-delta) - 1;

            if (r.top >= fromRow && r.bottom <= delEnd) {
                // fully inside deleted band → remove
                deleteMerge(idx);
                continue;
            }
            if (r.top >= fromRow) {
                // top is inside band, bottom is below → pull top to fromRow
                r.top = fromRow;
            }
            // clip bottom if it falls inside the deleted band
            if (r.bottom <= delEnd) {
                r.bottom = fromRow;  // will be validated below
            } else {
                r.bottom = static_cast<uint32_t>(static_cast<int32_t>(r.bottom) + delta);
            }
            if (r.top >= r.bottom && !(r.top == r.bottom)) {
                deleteMerge(idx);
                continue;
            }
            // If top was below fromRow but not in band, slide it
            if (r.top > fromRow && r.top > delEnd) {
                r.top = static_cast<uint32_t>(static_cast<int32_t>(r.top) + delta);
            }
        } else {
            // Insertion: push everything at/below fromRow down
            if (r.top >= fromRow)
                r.top  = static_cast<uint32_t>(static_cast<int32_t>(r.top)  + delta);
            r.bottom = static_cast<uint32_t>(static_cast<int32_t>(r.bottom) + delta);
        }

        // Rebuild reference string and sync XML
        entry.reference = XLCellReference(r.top, r.left).address() + ":" +
                          XLCellReference(r.bottom, r.right).address();

        // Walk XML to find and update the corresponding mergeCell node
        XMLNode node = m_mergeCellsNode.first_child_of_type(pugi::node_element);
        for (XLMergeIndex k = 0; k < idx && not node.empty(); ++k)
            node = node.next_sibling_of_type(pugi::node_element);
        if (not node.empty())
            node.attribute("ref").set_value(entry.reference.c_str());
    }
}

/**
 * @details Mirror of shiftRows but operating on column (left/right) coordinates.
 */
void XLMergeCells::shiftCols(int32_t delta, uint16_t fromCol)
{
    if (delta == 0 || m_mergeCache.empty()) return;

    for (auto i = static_cast<int32_t>(m_mergeCache.size()) - 1; i >= 0; --i) {
        auto idx  = static_cast<XLMergeIndex>(i);
        auto& entry = m_mergeCache[static_cast<size_t>(i)];
        XLRect& r = entry.rect;

        if (r.right < fromCol) continue;  // entirely left of affected zone

        if (delta < 0) {
            uint16_t delEnd = static_cast<uint16_t>(fromCol + static_cast<uint16_t>(-delta) - 1);

            if (r.left >= fromCol && r.right <= delEnd) {
                deleteMerge(idx);
                continue;
            }
            if (r.left >= fromCol && r.left <= delEnd)
                r.left = fromCol;
            if (r.right <= delEnd) {
                r.right = static_cast<uint16_t>(fromCol - 1);
            } else {
                r.right = static_cast<uint16_t>(static_cast<int16_t>(r.right) + static_cast<int16_t>(delta));
            }
            if (r.right < r.left) {
                deleteMerge(idx);
                continue;
            }
            if (r.left > fromCol && r.left > delEnd)
                r.left = static_cast<uint16_t>(static_cast<int16_t>(r.left) + static_cast<int16_t>(delta));
        } else {
            if (r.left >= fromCol)
                r.left  = static_cast<uint16_t>(static_cast<int16_t>(r.left)  + static_cast<int16_t>(delta));
            r.right = static_cast<uint16_t>(static_cast<int16_t>(r.right) + static_cast<int16_t>(delta));
        }

        entry.reference = XLCellReference(r.top, r.left).address() + ":" +
                          XLCellReference(r.bottom, r.right).address();

        XMLNode node = m_mergeCellsNode.first_child_of_type(pugi::node_element);
        for (XLMergeIndex k = 0; k < idx && not node.empty(); ++k)
            node = node.next_sibling_of_type(pugi::node_element);
        if (not node.empty())
            node.attribute("ref").set_value(entry.reference.c_str());
    }
}

/**
 * @details
 */
void XLMergeCells::print(std::basic_ostream<char>& ostr) const { m_mergeCellsNode.print(ostr); }

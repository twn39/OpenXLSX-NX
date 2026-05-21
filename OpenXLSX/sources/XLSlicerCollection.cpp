#include "XLSlicerCollection.hpp"
#include "XLSlicerCache.hpp"
#include "XLWorksheet.hpp"
#include "XLDocument.hpp"
#include "XLRelationships.hpp"
#include "XLUtilities.hpp"
#include "XLException.hpp"
#include "XLCellReference.hpp"
#include "XLContentTypes.hpp"
#include "fmt/format.h"
#include <algorithm>
#include <stdexcept>

using namespace OpenXLSX;

// ─────────────────────────────────────────────────────────────────────────────
// Helper: locate the drawing anchor whose sle:slicer name matches
// ─────────────────────────────────────────────────────────────────────────────
namespace {

/// Find xdr:twoCellAnchor or xdr:oneCellAnchor containing a slicer with the
/// given name attribute. Returns empty node if not found.
XMLNode findSlicerAnchor(const XMLNode& drawingRoot, std::string_view slicerName)
{
    for (auto anchor : drawingRoot.children()) {
        // Look inside mc:AlternateContent → mc:Choice → xdr:graphicFrame → a:graphic →
        // a:graphicData → sle:slicer [name]
        auto alt = anchor.child("mc:AlternateContent");
        if (!alt) continue;
        for (auto choice : alt.children("mc:Choice")) {
            auto gf   = choice.child("xdr:graphicFrame");
            if (!gf) continue;
            auto sl   = gf.child("a:graphic").child("a:graphicData").child("sle:slicer");
            if (!sl)
                sl    = gf.child("a:graphic").child("a:graphicData").child("sle15:slicer");
            if (sl && std::string_view(sl.attribute("name").value()) == slicerName)
                return anchor;
        }
        // Also check xdr:graphicFrame directly (no AlternateContent)
        auto gf   = anchor.child("xdr:graphicFrame");
        if (gf) {
            auto sl = gf.child("a:graphic").child("a:graphicData").child("sle:slicer");
            if (!sl)
                sl  = gf.child("a:graphic").child("a:graphicData").child("sle15:slicer");
            if (sl && std::string_view(sl.attribute("name").value()) == slicerName)
                return anchor;
        }
    }
    return XMLNode{};
}

/// Find the slicerCacheN.xml XLXmlData for a given cache name.
XLXmlData* findCacheXml(XLDocument& doc, std::string_view cacheName)
{
    // Iterate through all content items looking for SlicerCache content type
    for (const auto& item : doc.contentTypes().getContentItems()) {
        if (item.type() != XLContentType::SlicerCache) continue;
        std::string relPath = item.path();
        if (!relPath.empty() && relPath[0] == '/') relPath = relPath.substr(1);
        XLXmlData* xmlData = doc.getXmlData(XLInternalAccess{}, relPath, /*doNotThrow=*/true);
        if (!xmlData) continue;
        auto root = xmlData->getXmlDocument()->document_element();
        if (std::string_view(root.attribute("name").value()) == cacheName)
            return xmlData;
    }
    return nullptr;
}

} // anonymous namespace

// ─────────────────────────────────────────────────────────────────────────────
// XLSlicerCollection
// ─────────────────────────────────────────────────────────────────────────────

XLSlicerCollection::XLSlicerCollection(XLWorksheet* worksheet)
    : m_worksheet(worksheet)
{}

void XLSlicerCollection::load() const
{
    if (m_loaded || !m_worksheet) return;
    m_loaded = true;
    m_slicers.clear();

    // ── Gather all slicer XML paths from worksheet relationships ──────────────
    auto& ws  = *const_cast<XLWorksheet*>(m_worksheet);
    auto& rels = ws.relationships();  // friend access
    for (const auto& rel : rels.relationships()) {
        if (rel.type() != XLRelationshipType::Slicer) continue;

        std::string target  = rel.target();
        std::string absPath = eliminateDotAndDotDotFromPath("xl/worksheets/" + target);

        auto& doc = const_cast<XLDocument&>(m_worksheet->parentDoc());

        // Get or load the slicer XML (slicer parts are lazily parsed from the ZIP)
        XLXmlData* slicerXml = doc.getXmlData(XLInternalAccess{}, absPath, /*doNotThrow=*/true);
        if (!slicerXml) {
            // Not yet parsed — load it from the archive
            slicerXml = doc.addXmlData(XLInternalAccess{}, absPath, "", XLContentType::Slicer);
        }
        if (!slicerXml) continue;

        // Each slicerN.xml may contain multiple <slicer> children
        XMLNode slicersRoot = slicerXml->getXmlDocument()->document_element();
        for (auto slicerNode : slicersRoot.children("slicer")) {
            std::string slicerName = slicerNode.attribute("name").value();
            std::string cacheName  = slicerNode.attribute("cache").value();

            // ── Find cache XML ───────────────────────────────────────────────
            XLXmlData* cacheXml = findCacheXml(doc, cacheName);

            // ── Find drawing anchor ──────────────────────────────────────────
            XMLNode anchorNode;
            if (m_worksheet->hasDrawing()) {
                auto& drw    = const_cast<XLWorksheet*>(m_worksheet)->drawing();
                auto  drwRoot = drw.xmlDocument().document_element();
                anchorNode   = findSlicerAnchor(drwRoot, slicerName);
            }

            m_slicers.emplace_back(slicerXml, slicerNode, cacheXml, anchorNode,
                                   const_cast<XLWorksheet*>(m_worksheet));
        }
    }
}


size_t XLSlicerCollection::count() const
{
    load();
    return m_slicers.size();
}

bool XLSlicerCollection::empty() const { return count() == 0; }

bool XLSlicerCollection::contains(std::string_view name) const
{
    load();
    for (const auto& s : m_slicers) {
        if (s.name() == name) return true;
    }
    return false;
}

XLSlicer XLSlicerCollection::find(std::string_view name) const
{
    load();
    for (const auto& s : m_slicers) {
        if (s.name() == name) return s;
    }
    return XLSlicer{};  // invalid
}

XLSlicer& XLSlicerCollection::operator[](size_t index)
{
    load();
    if (index >= m_slicers.size())
        throw std::out_of_range("XLSlicerCollection: index out of range");
    return m_slicers[index];
}

XLSlicer& XLSlicerCollection::operator[](std::string_view name)
{
    load();
    for (auto& s : m_slicers) {
        if (s.name() == name) return s;
    }
    throw std::out_of_range("XLSlicerCollection: slicer '" + std::string(name) + "' not found");
}

const XLSlicer& XLSlicerCollection::operator[](size_t index) const
{
    load();
    if (index >= m_slicers.size())
        throw std::out_of_range("XLSlicerCollection: index out of range");
    return m_slicers[index];
}

XLSlicerCollection::Iterator XLSlicerCollection::begin()
{
    load();
    return Iterator(&m_slicers, 0);
}

XLSlicerCollection::Iterator XLSlicerCollection::end()
{
    load();
    return Iterator(&m_slicers, m_slicers.size());
}

XLSlicerCollection::ConstIterator XLSlicerCollection::begin() const
{
    load();
    return ConstIterator(&m_slicers, 0);
}

XLSlicerCollection::ConstIterator XLSlicerCollection::end() const
{
    load();
    return ConstIterator(&m_slicers, m_slicers.size());
}

XLSlicerBuilder XLSlicerCollection::add(std::string_view    cellRef,
                                        const XLTable&      table,
                                        std::string_view    columnName)
{
    return XLSlicerBuilder(m_worksheet, std::string(cellRef),
                           &table, nullptr, std::string(columnName));
}

XLSlicerBuilder XLSlicerCollection::add(std::string_view     cellRef,
                                        const XLPivotTable&  pivotTable,
                                        std::string_view     fieldName)
{
    return XLSlicerBuilder(m_worksheet, std::string(cellRef),
                           nullptr, &pivotTable, std::string(fieldName));
}

void XLSlicerCollection::remove(std::string_view name)
{
    if (m_worksheet)
        const_cast<XLWorksheet*>(m_worksheet)->deleteSlicer(std::string(name));
    reload();
}

void XLSlicerCollection::remove(size_t index)
{
    load();
    if (index >= m_slicers.size())
        throw std::out_of_range("XLSlicerCollection: index out of range");
    std::string name = m_slicers[index].name();
    remove(name);
}

void XLSlicerCollection::reload()
{
    m_loaded = false;
    m_slicers.clear();
}

// ─────────────────────────────────────────────────────────────────────────────
// XLSlicerBuilder
// ─────────────────────────────────────────────────────────────────────────────

XLSlicerBuilder::XLSlicerBuilder(XLWorksheet*        worksheet,
                                 std::string         cellRef,
                                 const XLTable*      table,
                                 const XLPivotTable* pivot,
                                 std::string         columnName)
    : m_worksheet(worksheet)
    , m_table(table)
    , m_pivot(pivot)
    , m_cellRef(std::move(cellRef))
    , m_columnName(std::move(columnName))
{
    // Defaults: name and caption from column name
    m_state.name    = m_columnName;
    m_state.caption = m_columnName;
}

XLSlicerBuilder::XLSlicerBuilder(XLSlicerBuilder&& other) noexcept
    : m_worksheet(other.m_worksheet)
    , m_table(other.m_table)
    , m_pivot(other.m_pivot)
    , m_cellRef(std::move(other.m_cellRef))
    , m_columnName(std::move(other.m_columnName))
    , m_state(std::move(other.m_state))
    , m_committed(other.m_committed)
    , m_result(std::move(other.m_result))
{
    other.m_committed = true;  // prevent double-commit
}

XLSlicerBuilder::~XLSlicerBuilder()
{
    if (!m_committed) {
        try { commit(); } catch (...) {}
    }
}

XLSlicerBuilder& XLSlicerBuilder::name(std::string_view n)
{
    m_state.name = n;
    return *this;
}

XLSlicerBuilder& XLSlicerBuilder::caption(std::string_view c)
{
    m_state.caption = c;
    return *this;
}

XLSlicerBuilder& XLSlicerBuilder::style(XLSlicerStyle s)
{
    m_state.styleRaw = xlSlicerStyleToString(s);
    return *this;
}

XLSlicerBuilder& XLSlicerBuilder::styleRaw(std::string_view rawName)
{
    m_state.styleRaw = rawName;
    return *this;
}

XLSlicerBuilder& XLSlicerBuilder::size(uint32_t widthPx, uint32_t heightPx)
{
    m_state.widthPx  = widthPx;
    m_state.heightPx = heightPx;
    return *this;
}

XLSlicerBuilder& XLSlicerBuilder::showOnly(const std::vector<std::string>& items)
{
    m_state.selectedItems = items;
    return *this;
}

XLSlicerBuilder& XLSlicerBuilder::columnCount(int cols)
{
    m_state.columnCount = cols;
    return *this;
}

XLSlicerBuilder& XLSlicerBuilder::sortDescending(bool desc)
{
    m_state.sortDesc = desc;
    return *this;
}

XLSlicerBuilder& XLSlicerBuilder::lockedPosition(bool locked)
{
    m_state.locked = locked;
    return *this;
}

XLSlicerBuilder& XLSlicerBuilder::offset(int32_t dx, int32_t dy)
{
    m_state.offsetX = dx;
    m_state.offsetY = dy;
    return *this;
}

void XLSlicerBuilder::commit()
{
    if (m_committed || !m_worksheet) { m_committed = true; return; }
    m_committed = true;

    // Build XLSlicerOptions for the underlying impl
    XLSlicerOptions opts;
    opts.name          = m_state.name;
    opts.caption       = m_state.caption;
    opts.slicerStyle   = m_state.styleRaw;
    opts.width         = m_state.widthPx;
    opts.height        = m_state.heightPx;
    opts.offsetX       = m_state.offsetX;
    opts.offsetY       = m_state.offsetY;
    opts.selectedItems = m_state.selectedItems;
    // columnCount, sortDesc, locked applied post-creation below

    if (m_table)
        m_worksheet->addTableSlicer(m_cellRef, *m_table, m_columnName, opts);
    else if (m_pivot)
        m_worksheet->addPivotSlicer(m_cellRef, *m_pivot, m_columnName, opts);

    // Post-creation: apply extra attributes not yet in the legacy opts
    // (columnCount, sortDesc, locked)
    auto& coll = m_worksheet->slicers();
    coll.reload();
    if (coll.contains(m_state.name)) {
        auto& s = coll[m_state.name];
        if (m_state.columnCount > 0) s.setColumnCount(m_state.columnCount);
        if (m_state.sortDesc)        s.setSortDescending(true);
        if (m_state.locked)          s.setLockedPosition(true);
        m_result = s;
    }
}

XLSlicer XLSlicerBuilder::build()
{
    if (!m_committed) commit();
    return m_result;
}

XLSlicerBuilder::operator XLSlicer()
{
    return build();
}

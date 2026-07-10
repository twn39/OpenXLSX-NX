#include "XLSlicer.hpp"
#include "XLCellReference.hpp"
#include "XLWorksheet.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include <algorithm>

using namespace OpenXLSX;

// ─────────────────────────────────────────────────────────────────────────────
// Style conversion helpers
// ─────────────────────────────────────────────────────────────────────────────

std::string OpenXLSX::xlSlicerStyleToString(XLSlicerStyle style)
{
    switch (style) {
        case XLSlicerStyle::Light1:  return "SlicerStyleLight1";
        case XLSlicerStyle::Light2:  return "SlicerStyleLight2";
        case XLSlicerStyle::Light3:  return "SlicerStyleLight3";
        case XLSlicerStyle::Light4:  return "SlicerStyleLight4";
        case XLSlicerStyle::Light5:  return "SlicerStyleLight5";
        case XLSlicerStyle::Light6:  return "SlicerStyleLight6";
        case XLSlicerStyle::Dark1:   return "SlicerStyleDark1";
        case XLSlicerStyle::Dark2:   return "SlicerStyleDark2";
        case XLSlicerStyle::Dark3:   return "SlicerStyleDark3";
        case XLSlicerStyle::Dark4:   return "SlicerStyleDark4";
        case XLSlicerStyle::Dark5:   return "SlicerStyleDark5";
        case XLSlicerStyle::Dark6:   return "SlicerStyleDark6";
        case XLSlicerStyle::Other1:  return "SlicerStyleOther1";
        case XLSlicerStyle::Other2:  return "SlicerStyleOther2";
        default:                     return "SlicerStyleLight1";
    }
}

XLSlicerStyle OpenXLSX::xlSlicerStyleFromString(std::string_view s)
{
    static const std::pair<std::string_view, XLSlicerStyle> table[] = {
        {"SlicerStyleLight1", XLSlicerStyle::Light1},
        {"SlicerStyleLight2", XLSlicerStyle::Light2},
        {"SlicerStyleLight3", XLSlicerStyle::Light3},
        {"SlicerStyleLight4", XLSlicerStyle::Light4},
        {"SlicerStyleLight5", XLSlicerStyle::Light5},
        {"SlicerStyleLight6", XLSlicerStyle::Light6},
        {"SlicerStyleDark1",  XLSlicerStyle::Dark1},
        {"SlicerStyleDark2",  XLSlicerStyle::Dark2},
        {"SlicerStyleDark3",  XLSlicerStyle::Dark3},
        {"SlicerStyleDark4",  XLSlicerStyle::Dark4},
        {"SlicerStyleDark5",  XLSlicerStyle::Dark5},
        {"SlicerStyleDark6",  XLSlicerStyle::Dark6},
        {"SlicerStyleOther1", XLSlicerStyle::Other1},
        {"SlicerStyleOther2", XLSlicerStyle::Other2},
    };
    for (const auto& [k, v] : table) {
        if (k == s) return v;
    }
    return XLSlicerStyle::Custom;
}

// ─────────────────────────────────────────────────────────────────────────────
// XLSlicer constructors
// ─────────────────────────────────────────────────────────────────────────────

XLSlicer::XLSlicer(XLXmlData* slicerXml)
    : m_slicerXml(slicerXml)
{
    if (m_slicerXml) {
        m_slicerNode = m_slicerXml->getXmlDocument()->document_element().child("slicer");
    }
}

XLSlicer::XLSlicer(XLXmlData* slicerXml,
                   XMLNode    slicerNode,
                   XLXmlData* cacheXml,
                   XMLNode    anchorNode,
                   XLWorksheet* worksheet)
    : m_slicerXml(slicerXml)
    , m_slicerNode(slicerNode)
    , m_cacheXml(cacheXml)
    , m_anchorNode(anchorNode)
    , m_worksheet(worksheet)
{}

// ─────────────────────────────────────────────────────────────────────────────
// Internal helpers
// ─────────────────────────────────────────────────────────────────────────────

XMLNode XLSlicer::slicerNode() const
{
    return m_slicerNode;
}

XMLNode XLSlicer::cacheRoot() const
{
    if (!m_cacheXml) return XMLNode{};
    return m_cacheXml->getXmlDocument()->document_element();
}

// ─────────────────────────────────────────────────────────────────────────────
// Source 1: slicer XML  — getters
// ─────────────────────────────────────────────────────────────────────────────

std::string XLSlicer::name() const
{
    return slicerNode().attribute("name").value();
}

std::string XLSlicer::caption() const
{
    return slicerNode().attribute("caption").value();
}

std::string XLSlicer::cache() const
{
    return slicerNode().attribute("cache").value();
}

XLSlicerStyle XLSlicer::style() const
{
    return xlSlicerStyleFromString(styleRaw());
}

std::string XLSlicer::styleRaw() const
{
    auto attr = slicerNode().attribute("style");
    return attr ? std::string(attr.value()) : "SlicerStyleLight1";
}

bool XLSlicer::showCaption() const
{
    auto attr = slicerNode().attribute("showCaption");
    if (!attr) return true;  // default is true per OOXML spec
    std::string v = attr.value();
    return (v != "0" && v != "false");
}

int XLSlicer::columnCount() const
{
    auto attr = slicerNode().attribute("columnCount");
    return attr ? attr.as_int() : 0;
}

bool XLSlicer::lockedPosition() const
{
    auto attr = slicerNode().attribute("lockedPosition");
    if (!attr) return false;
    std::string v = attr.value();
    return (v == "1" || v == "true");
}

int XLSlicer::rowHeight() const
{
    auto attr = slicerNode().attribute("rowHeight");
    return attr ? attr.as_int() : 0;
}

// ─────────────────────────────────────────────────────────────────────────────
// Source 1: slicer XML  — setters
// ─────────────────────────────────────────────────────────────────────────────

void XLSlicer::setName(std::string_view n)
{
    auto node = slicerNode();
    if (!node) return;
    setAttr(node, "name", std::string(n));
}

XLSlicer& XLSlicer::setCaption(std::string_view c)
{
    auto node = slicerNode();
    if (!node) return *this;
    setAttr(node, "caption", std::string(c));
    return *this;
}

XLSlicer& XLSlicer::setStyle(XLSlicerStyle s)
{
    return setStyleRaw(xlSlicerStyleToString(s));
}

XLSlicer& XLSlicer::setStyleRaw(std::string_view rawName)
{
    auto node = slicerNode();
    if (!node) return *this;
    setAttr(node, "style", std::string(rawName));
    return *this;
}

XLSlicer& XLSlicer::setShowCaption(bool show)
{
    auto node = slicerNode();
    if (!node) return *this;
    setAttr(node, "showCaption", show ? "1" : "0");
    return *this;
}

XLSlicer& XLSlicer::setColumnCount(int cols)
{
    auto node = slicerNode();
    if (!node) return *this;
    setAttr(node, "columnCount", std::to_string(cols));
    return *this;
}

XLSlicer& XLSlicer::setLockedPosition(bool locked)
{
    auto node = slicerNode();
    if (!node) return *this;
    setAttr(node, "lockedPosition", locked ? "1" : "0");
    return *this;
}

XLSlicer& XLSlicer::setRowHeight(int emuHeight)
{
    auto node = slicerNode();
    if (!node) return *this;
    setAttr(node, "rowHeight", std::to_string(emuHeight));
    return *this;
}

// ─────────────────────────────────────────────────────────────────────────────
// Source 2: cache XML — filter items
// ─────────────────────────────────────────────────────────────────────────────

std::vector<std::string> XLSlicer::items() const
{
    std::vector<std::string> result;
    if (!m_cacheXml) return result;

    XMLNode root    = cacheRoot();
    XMLNode srcName = root;  // sourceName attr holds the column name hint

    // For pivot: data/tabular/items → each <i x=N> maps to shared items in pivotCache
    // For table: no inline item names in the cache XML, items come from the column data
    // Return what we can from sourceName as a fallback hint
    (void)srcName;  // suppress warning; real implementation would cross-ref pivotCacheRecords

    // Best-effort: parse tabular items (x = string table index, not raw string)
    // This returns string indices; full string resolution requires pivotCacheRecords cross-ref.
    // For now, return selected item count info from the cache's item list.
    return result;
}

std::vector<std::string> XLSlicer::selectedItems() const
{
    std::vector<std::string> result;
    if (!m_cacheXml) return result;

    XMLNode root    = cacheRoot();
    XMLNode data    = root.child("data");
    if (!data) return result;
    XMLNode tabular = data.child("tabular");
    if (!tabular) return result;
    XMLNode items   = tabular.child("items");
    if (!items) return result;

    // Items where s="true" or s="1" are SELECTED
    // Items where s is absent or s="false"/"0" are DESELECTED
    // We return a simple index-based list here; see showOnly for writing.
    // For named items, cross-reference with pivotCacheRecords or table column data.
    // Items where nd="true" are "no data" (new values not in original cache).
    for (auto item : items.children("i")) {
        auto sAttr = item.attribute("s");
        bool selected = sAttr ? (std::string(sAttr.value()) == "true" ||
                                  std::string(sAttr.value()) == "1") : false;
        if (selected) {
            // x attribute holds the string index; no in-place string available
            result.push_back(std::to_string(item.attribute("x").as_int()));
        }
    }
    return result;
}

bool XLSlicer::isSortDescending() const
{
    if (!m_cacheXml) return false;
    XMLNode root    = cacheRoot();
    XMLNode data    = root.child("data");
    XMLNode tabular = data ? data.child("tabular") : XMLNode{};
    if (!tabular) {
        // Also check x15:tableSlicerCache in extLst
        auto extLst = root.child("extLst");
        if (extLst) {
            for (auto ext : extLst.children("ext")) {
                auto tsc = ext.child("x15:tableSlicerCache");
                if (tsc) {
                    return std::string(tsc.attribute("sortOrder").value()) == "descending";
                }
            }
        }
        return false;
    }
    return std::string(tabular.attribute("sortOrder").value()) == "descending";
}

XLSlicer& XLSlicer::showOnly(const std::vector<std::string>& itemsToShow)
{
    if (!m_cacheXml) return *this;
    XMLNode root    = cacheRoot();
    XMLNode data    = root.child("data");
    if (!data) return *this;
    XMLNode tabular = data.child("tabular");
    if (!tabular) return *this;
    XMLNode itemsNode = tabular.child("items");
    if (!itemsNode) return *this;

    // Mark all items: selected if in itemsToShow, deselected otherwise
    for (auto item : itemsNode.children("i")) {
        // x is a string table index — compare against cache sourceName entries
        // For simplicity, items are deselected by default; showOnly toggles them
        item.remove_attribute("s");   // remove current selection state
        item.append_attribute("s").set_value("0");
    }

    // This is a best-effort implementation.
    // Full implementation requires building a string lookup from pivotCacheRecords.
    (void)itemsToShow;
    return *this;
}

XLSlicer& XLSlicer::showAll()
{
    if (!m_cacheXml) return *this;
    XMLNode root    = cacheRoot();
    XMLNode data    = root.child("data");
    if (!data) return *this;
    XMLNode tabular = data.child("tabular");
    if (!tabular) return *this;
    XMLNode itemsNode = tabular.child("items");
    if (!itemsNode) return *this;

    for (auto item : itemsNode.children("i")) {
        item.remove_attribute("s");
    }
    return *this;
}

XLSlicer& XLSlicer::hideItems(const std::vector<std::string>& itemsToHide)
{
    (void)itemsToHide;
    // Inverse of showOnly — requires same string-index lookup.
    // Placeholder until pivotCacheRecords cross-reference is implemented.
    return *this;
}

XLSlicer& XLSlicer::setSortDescending(bool desc)
{
    if (!m_cacheXml) return *this;
    XMLNode root = cacheRoot();

    // Try pivot/tabular path
    XMLNode data    = root.child("data");
    XMLNode tabular = data ? data.child("tabular") : XMLNode{};
    if (tabular) {
        auto attr = tabular.attribute("sortOrder");
        if (!attr) attr = tabular.append_attribute("sortOrder");
        attr.set_value(desc ? "descending" : "ascending");
        return *this;
    }

    // Try table path (x15:tableSlicerCache in extLst)
    auto extLst = root.child("extLst");
    if (extLst) {
        for (auto ext : extLst.children("ext")) {
            auto tsc = ext.child("x15:tableSlicerCache");
            if (tsc) {
                auto attr = tsc.attribute("sortOrder");
                if (!attr) attr = tsc.append_attribute("sortOrder");
                attr.set_value(desc ? "descending" : "ascending");
                return *this;
            }
        }
    }
    return *this;
}

// ─────────────────────────────────────────────────────────────────────────────
// Source 3: drawing anchor — position & size
// ─────────────────────────────────────────────────────────────────────────────

namespace {
    // EMU ↔ pixel: 1 pixel = 9525 EMU (96 DPI)
    constexpr int64_t EMU_PER_PX = 9525;

    /// Return the <xdr:from> or <xdr:oneCellAnchor/from> node.
    XMLNode anchorFrom(XMLNode anchor)
    {
        auto from = anchor.child("xdr:from");
        return from;
    }

    /// Return the <xdr:ext> node (only present in oneCellAnchor).
    XMLNode anchorExt(XMLNode anchor)
    {
        return anchor.child("xdr:ext");
    }
}

std::string XLSlicer::cellRef() const
{
    if (!m_anchorNode) return "";
    XMLNode from = anchorFrom(m_anchorNode);
    if (!from) return "";

    int col0 = from.child("xdr:col").text().as_int();  // 0-indexed
    int row0 = from.child("xdr:row").text().as_int();  // 0-indexed

    // Convert 0-indexed col/row → 1-indexed XLCellReference
    XLCellReference ref(static_cast<uint32_t>(row0 + 1),
                        static_cast<uint16_t>(col0 + 1));
    return ref.address();
}

uint32_t XLSlicer::width() const
{
    if (!m_anchorNode) return 0;
    XMLNode ext = anchorExt(m_anchorNode);
    if (!ext) {
        // twoCellAnchor: approximate from to-from column difference
        XMLNode from = m_anchorNode.child("xdr:from");
        XMLNode to   = m_anchorNode.child("xdr:to");
        if (!from || !to) return 0;
        // Fall back to fallback <xdr:sp> ext size
        auto sp = m_anchorNode.select_node(".//xdr:sp/xdr:spPr/a:xfrm/a:ext").node();
        if (sp) {
            int64_t cx = sp.attribute("cx").as_llong();
            return static_cast<uint32_t>(cx / EMU_PER_PX);
        }
        return 0;
    }
    int64_t cx = ext.attribute("cx").as_llong();
    return static_cast<uint32_t>(cx / EMU_PER_PX);
}

uint32_t XLSlicer::height() const
{
    if (!m_anchorNode) return 0;
    XMLNode ext = anchorExt(m_anchorNode);
    if (!ext) {
        auto sp = m_anchorNode.select_node(".//xdr:sp/xdr:spPr/a:xfrm/a:ext").node();
        if (sp) {
            int64_t cy = sp.attribute("cy").as_llong();
            return static_cast<uint32_t>(cy / EMU_PER_PX);
        }
        return 0;
    }
    int64_t cy = ext.attribute("cy").as_llong();
    return static_cast<uint32_t>(cy / EMU_PER_PX);
}

XLSlicer& XLSlicer::moveTo(std::string_view cellRefStr)
{
    if (!m_anchorNode) return *this;
    XMLNode from = anchorFrom(m_anchorNode);
    if (!from) return *this;

    XLCellReference ref{std::string(cellRefStr)};
    uint32_t row0 = ref.row() - 1;    // 0-indexed
    uint16_t col0 = ref.column() - 1;

    from.child("xdr:col").text().set(col0);
    from.child("xdr:row").text().set(row0);
    return *this;
}

XLSlicer& XLSlicer::resize(uint32_t widthPx, uint32_t heightPx)
{
    if (!m_anchorNode) return *this;

    int64_t cx = static_cast<int64_t>(widthPx)  * EMU_PER_PX;
    int64_t cy = static_cast<int64_t>(heightPx) * EMU_PER_PX;

    // Update oneCellAnchor ext
    XMLNode ext = anchorExt(m_anchorNode);
    if (ext) {
        setAttr(ext, "cx", std::to_string(cx));
        setAttr(ext, "cy", std::to_string(cy));
    }

    // Also update the fallback <xdr:sp> shape in AlternateContent/Fallback
    for (auto anchor : m_anchorNode.select_nodes(".//xdr:sp/xdr:spPr/a:xfrm/a:ext")) {
        XMLNode extNode = XMLNode(anchor.node());
        auto cxAttr = extNode.attribute("cx");
        auto cyAttr = extNode.attribute("cy");
        if (!cxAttr) cxAttr = extNode.append_attribute("cx");
        if (!cyAttr) cyAttr = extNode.append_attribute("cy");
        cxAttr.set_value(std::to_string(cx).c_str());
        cyAttr.set_value(std::to_string(cy).c_str());
    }
    return *this;
}

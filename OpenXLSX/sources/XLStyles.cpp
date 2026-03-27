// ===== External Includes ===== //
#include <cstdint>
#include <gsl/gsl>
#include <string_view>
#include <fmt/format.h>
#include <memory>
#include <pugixml.hpp>
#include <stdexcept>
#include <string>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLColor.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLStyle.hpp"
#include "XLStyles.hpp"
#include "XLUtilities.hpp"
#include "XLStyles_Internal.hpp"

using namespace OpenXLSX;

namespace {
    enum XLStylesEntryType : uint8_t {
        XLStylesNumberFormats    = 0,
        XLStylesFonts            = 1,
        XLStylesFills            = 2,
        XLStylesBorders          = 3,
        XLStylesCellStyleFormats = 4,
        XLStylesCellFormats      = 5,
        XLStylesCellStyles       = 6,
        XLStylesColors           = 7,
        XLStylesDiffCellFormats  = 8,
        XLStylesTableStyles      = 9,
        XLStylesExtLst           = 10,
        XLStylesInvalid          = 255
    };

    constexpr EnumStringMap<XLStylesEntryType> XLStylesEntryTypeMap[] = {
        {"numFmts", XLStylesNumberFormats},
        {"fonts", XLStylesFonts},
        {"fills", XLStylesFills},
        {"borders", XLStylesBorders},
        {"cellStyleXfs", XLStylesCellStyleFormats},
        {"cellXfs", XLStylesCellFormats},
        {"cellStyles", XLStylesCellStyles},
        {"colors", XLStylesColors},
        {"dxfs", XLStylesDiffCellFormats},
        {"tableStyles", XLStylesTableStyles},
        {"extLst", XLStylesExtLst},
    };

    XLStylesEntryType XLStylesEntryTypeFromString(std::string_view str)
    {
        return EnumFromString(str, XLStylesEntryTypeMap, XLStylesEntryType::XLStylesInvalid);
    }

    const char* XLStylesEntryTypeToString(XLStylesEntryType val)
    {
        return EnumToString(val, XLStylesEntryTypeMap, "(invalid)");
    }

    void wrapNode(XMLNode parentNode, const XMLNode& node, std::string_view prefix)
    {
        if (not node.empty() and prefix.length() > 0) {
            parentNode.insert_child_before(pugi::node_pcdata, node).set_value(std::string(prefix).c_str());    // insert prefix before node opening tag
            XMLNode n = node;
            n.append_child(pugi::node_pcdata).set_value(std::string(prefix).c_str());    // insert prefix before node closing tag (within node)
        }
    }
}

// ===== XLStyles, master class

XLStyles::XLStyles() {}

XLStyles::XLStyles(gsl::not_null<XLXmlData*> xmlData, bool suppressWarnings, std::string_view stylesPrefix)
    : XLXmlFile(xmlData),
      m_suppressWarnings(suppressWarnings)
{
    XMLDocument& doc = xmlDocument();
    if (doc.document_element().empty())    // handle a bad (no document element) xl/styles.xml
        doc.load_string("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
                        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
                        "</styleSheet>",
                        pugi_parse_settings);

    XMLNode node = doc.document_element().first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        XLStylesEntryType e = XLStylesEntryTypeFromString(node.name());
        switch (e) {
            case XLStylesNumberFormats:
                m_numberFormats = std::make_unique<XLNumberFormats>(node);
                break;
            case XLStylesFonts:
                m_fonts = std::make_unique<XLFonts>(node);
                break;
            case XLStylesFills:
                m_fills = std::make_unique<XLFills>(node);
                break;
            case XLStylesBorders:
                m_borders = std::make_unique<XLBorders>(node);
                break;
            case XLStylesCellStyleFormats:
                m_cellStyleFormats = std::make_unique<XLCellFormats>(node);
                break;
            case XLStylesCellFormats:
                m_cellFormats = std::make_unique<XLCellFormats>(node, XLPermitXfID);
                break;
            case XLStylesCellStyles:
                m_cellStyles = std::make_unique<XLCellStyles>(node);
                break;
            case XLStylesDiffCellFormats:
                m_dxfs = std::make_unique<XLDxfs>(node);
                break;
            case XLStylesColors:
                [[fallthrough]];
            case XLStylesTableStyles:
                [[fallthrough]];
            case XLStylesExtLst:
                if (!m_suppressWarnings)
                    std::cout << "XLStyles: Ignoring currently unsupported <" << std::string(XLStylesEntryTypeToString(e)) + "> node" << std::endl;
                break;
            case XLStylesInvalid:
                [[fallthrough]];
            default:
                if (!m_suppressWarnings)
                    std::cerr << "WARNING: XLStyles constructor: invalid styles subnode \"" << node.name() << "\"" << std::endl;
                break;
        }
        node = node.next_sibling_of_type(pugi::node_element);
    }

    // ===== Fallbacks: create root style nodes (in reverse order, using prepend_child)
    if (!m_dxfs) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesDiffCellFormats));
        wrapNode(doc.document_element(), node, stylesPrefix);
        m_dxfs = std::make_unique<XLDxfs>(node);
    }
    if (!m_cellStyles) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesCellStyles));
        wrapNode(doc.document_element(), node, stylesPrefix);
        m_cellStyles = std::make_unique<XLCellStyles>(node);
    }
    if (!m_cellFormats) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesCellFormats));
        wrapNode(doc.document_element(), node, stylesPrefix);
        m_cellFormats = std::make_unique<XLCellFormats>(node, XLPermitXfID);
    }
    if (m_cellFormats->count() == 0) {    // if the cell formats array is empty
        // ===== Create a default empty cell format with ID 0 (== XLDefaultCellFormat) because when XLDefaultCellFormat
        //        is assigned to an XLRow, the intention is interpreted as "set the cell back to default formatting",
        //        which does not trigger setting the attribute customFormat="true".
        //       To avoid confusing the user when the first style created does not work for rows, and setting a cell's
        //        format back to XLDefaultCellFormat would cause an actual formatting (if assigned format ID 0), this
        //        initial entry with no properties is created and henceforth ignored
        m_cellFormats->create();
    }

    if (!m_cellStyleFormats) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesCellStyleFormats));
        wrapNode(doc.document_element(), node, stylesPrefix);
        m_cellStyleFormats = std::make_unique<XLCellFormats>(node);
    }
    if (!m_borders) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesBorders));
        wrapNode(doc.document_element(), node, stylesPrefix);
        m_borders = std::make_unique<XLBorders>(node);
    }
    if (!m_fills) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesFills));
        wrapNode(doc.document_element(), node, stylesPrefix);
        m_fills = std::make_unique<XLFills>(node);
    }
    if (!m_fonts) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesFonts));
        wrapNode(doc.document_element(), node, stylesPrefix);
        m_fonts = std::make_unique<XLFonts>(node);
    }
    if (!m_numberFormats) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesNumberFormats));
        wrapNode(doc.document_element(), node, stylesPrefix);
        m_numberFormats = std::make_unique<XLNumberFormats>(node);
    }
}

XLStyles::~XLStyles() {}

XLStyles::XLStyles(XLStyles&& other) noexcept
    : XLXmlFile(other),
      m_suppressWarnings(other.m_suppressWarnings),
      m_numberFormats(std::move(other.m_numberFormats)),
      m_fonts(std::move(other.m_fonts)),
      m_fills(std::move(other.m_fills)),
      m_borders(std::move(other.m_borders)),
      m_cellStyleFormats(std::move(other.m_cellStyleFormats)),
      m_cellFormats(std::move(other.m_cellFormats)),
      m_cellStyles(std::move(other.m_cellStyles)),
      m_dxfs(std::move(other.m_dxfs))
{}

XLStyles::XLStyles(const XLStyles& other)
    : XLXmlFile(other),
      m_suppressWarnings(other.m_suppressWarnings),
      m_numberFormats(std::make_unique<XLNumberFormats>(*other.m_numberFormats)),
      m_fonts(std::make_unique<XLFonts>(*other.m_fonts)),
      m_fills(std::make_unique<XLFills>(*other.m_fills)),
      m_borders(std::make_unique<XLBorders>(*other.m_borders)),
      m_cellStyleFormats(std::make_unique<XLCellFormats>(*other.m_cellStyleFormats)),
      m_cellFormats(std::make_unique<XLCellFormats>(*other.m_cellFormats)),
      m_cellStyles(std::make_unique<XLCellStyles>(*other.m_cellStyles)),
      m_dxfs(std::make_unique<XLDxfs>(*other.m_dxfs))
{}

XLStyles& XLStyles::operator=(XLStyles&& other) noexcept
{
    if (&other != this) {
        XLXmlFile::operator=(std::move(other));
        m_suppressWarnings = other.m_suppressWarnings;
        m_numberFormats    = std::move(other.m_numberFormats);
        m_fonts            = std::move(other.m_fonts);
        m_fills            = std::move(other.m_fills);
        m_borders          = std::move(other.m_borders);
        m_cellStyleFormats = std::move(other.m_cellStyleFormats);
        m_cellFormats      = std::move(other.m_cellFormats);
        m_cellStyles       = std::move(other.m_cellStyles);
        m_dxfs             = std::move(other.m_dxfs);
    }
    return *this;
}

XLStyles& XLStyles::operator=(const XLStyles& other)
{
    if (&other != this) {
        XLStyles temp = other;              // copy-construct
        *this         = std::move(temp);    // move-assign & invalidate temp
    }
    return *this;
}

uint32_t XLStyles::createNumberFormat(std::string_view formatCode)
{
    return m_numberFormats->createNumberFormat(formatCode);
}

XLNumberFormats& XLStyles::numberFormats() const { return *m_numberFormats; }

XLFonts& XLStyles::fonts() const { return *m_fonts; }

XLFills& XLStyles::fills() const { return *m_fills; }

XLBorders& XLStyles::borders() const { return *m_borders; }

XLCellFormats& XLStyles::cellStyleFormats() const { return *m_cellStyleFormats; }

XLCellFormats& XLStyles::cellFormats() const { return *m_cellFormats; }

XLCellStyles& XLStyles::cellStyles() const { return *m_cellStyles; }

XLDxfs& XLStyles::dxfs() const { return *m_dxfs; }

XLStyleIndex XLStyles::addDxf(const XLDxf& dxf) { return m_dxfs->create(dxf); }

XLStyleIndex XLStyles::addNamedStyle(std::string_view name,
                                     std::optional<XLStyleIndex> fontId,
                                     std::optional<XLStyleIndex> fillId,
                                     std::optional<XLStyleIndex> borderId,
                                     std::optional<XLStyleIndex> numFmtId)
{
    // 1. Create format definition in <cellStyleXfs>
    XLStyleIndex styleXfIdx = cellStyleFormats().create();
    auto styleXf = cellStyleFormats()[styleXfIdx];
    
    if (fontId) {
        styleXf.setFontIndex(*fontId);
        styleXf.setApplyFont(true);
    }
    if (fillId) {
        styleXf.setFillIndex(*fillId);
        styleXf.setApplyFill(true);
    }
    if (borderId) {
        styleXf.setBorderIndex(*borderId);
        styleXf.setApplyBorder(true);
    }
    if (numFmtId) {
        styleXf.setNumberFormatId(*numFmtId);
        styleXf.setApplyNumberFormat(true);
    }

    // 2. Register the name in <cellStyles>
    XLStyleIndex cellStyleIdx = cellStyles().create();
    auto cellStyle = cellStyles()[cellStyleIdx];
    cellStyle.setName(name);
    cellStyle.setXfId(styleXfIdx);

    // 3. Create a ready-to-use cell format in <cellXfs> that inherits this style
    XLStyleIndex cellXfIdx = cellFormats().create();
    auto cellXf = cellFormats()[cellXfIdx];
    
    cellXf.setXfId(styleXfIdx); // Bind to the named style format
    
    // Excel strictly requires these to be mirrored in the cellXfs node
    if (fontId) {
        cellXf.setFontIndex(*fontId);
        cellXf.setApplyFont(true);
    }
    if (fillId) {
        cellXf.setFillIndex(*fillId);
        cellXf.setApplyFill(true);
    }
    if (borderId) {
        cellXf.setBorderIndex(*borderId);
        cellXf.setApplyBorder(true);
    }
    if (numFmtId) {
        cellXf.setNumberFormatId(*numFmtId);
        cellXf.setApplyNumberFormat(true);
    }

    return cellXfIdx;
}

XLStyleIndex XLStyles::namedStyle(std::string_view name) const
{
    // Find the cellStyle by name to get its xfId
    XLStyleIndex targetStyleXfId = XLInvalidStyleIndex;
    for (size_t i = 0; i < cellStyles().count(); ++i) {
        if (cellStyles()[i].name() == name) {
            targetStyleXfId = cellStyles()[i].xfId();
            break;
        }
    }
    
    if (targetStyleXfId == XLInvalidStyleIndex) return XLInvalidStyleIndex;

    // Find a corresponding cellXfs entry that uses this xfId
    for (size_t i = 0; i < cellFormats().count(); ++i) {
        if (cellFormats()[i].xfId() == targetStyleXfId) {
            return gsl::narrow_cast<XLStyleIndex>(i);
        }
    }
    
    return XLInvalidStyleIndex;
}

XLStyleIndex XLStyles::findOrCreateStyle(const XLStyle& style)
{
    // ===== Step 1: Resolve font index (deduplicated) ========================
    // Pattern: create a tentative entry, configure it, then call findOrCreate.
    // findOrCreate returns the canonical index (possibly for a pre-existing entry).
    // The tentative entry that was created lives in the pool but is never referenced
    // if it was a duplicate — the dedup cache ensures all future lookups return the
    // canonical index.  Pool bloat from tentative entries is bounded by the number of
    // distinct style.findOrCreateStyle calls, which is typically very small.
    XLStyleIndex fontIdx = 0;    // default font (index 0)
    if (style.font.name || style.font.size || style.font.color ||
        style.font.bold || style.font.italic || style.font.underline || style.font.strikethrough)
    {
        // Create a writable copy of the default font, then configure it
        XLStyleIndex tentativeFontIdx = fonts().create(fonts().fontByIndex(0));
        XLFont f = fonts().fontByIndex(tentativeFontIdx);
        if (style.font.name)          f.setFontName(*style.font.name);
        if (style.font.size)          f.setFontSize(*style.font.size);
        if (style.font.color)         f.setFontColor(*style.font.color);
        if (style.font.bold)          f.setBold(*style.font.bold);
        if (style.font.italic)        f.setItalic(*style.font.italic);
        if (style.font.underline)     f.setUnderline(*style.font.underline ? XLUnderlineSingle : XLUnderlineNone);
        if (style.font.strikethrough) f.setStrikethrough(*style.font.strikethrough);
        // findOrCreate: if an identical font is already in the pool it returns that index;
        // otherwise it creates a new entry (which happens to be the tentative one).
        fontIdx = fonts().findOrCreate(f);
    }

    // ===== Step 2: Resolve fill index (deduplicated) ========================
    XLStyleIndex fillIdx = 0;    // default fill
    if (style.fill.pattern || style.fill.fgColor || style.fill.bgColor) {
        XLStyleIndex tentativeFillIdx = fills().create();
        XLFill fill = fills().fillByIndex(tentativeFillIdx);
        if (style.fill.pattern)  fill.setPatternType(*style.fill.pattern);
        if (style.fill.fgColor)  fill.setColor(*style.fill.fgColor);
        if (style.fill.bgColor)  fill.setBackgroundColor(*style.fill.bgColor);
        fillIdx = fills().findOrCreate(fill);
    }

    // ===== Step 3: Resolve border index (deduplicated) ======================
    XLStyleIndex borderIdx = 0;    // default border
    if (style.border.left.style || style.border.left.color ||
        style.border.right.style || style.border.right.color ||
        style.border.top.style || style.border.top.color ||
        style.border.bottom.style || style.border.bottom.color)
    {
        XLStyleIndex tentativeBorderIdx = borders().create();
        XLBorder border = borders().borderByIndex(tentativeBorderIdx);
        if (style.border.left.style)
            border.setLeft(*style.border.left.style, style.border.left.color.value_or(XLColor("FF000000")));
        if (style.border.right.style)
            border.setRight(*style.border.right.style, style.border.right.color.value_or(XLColor("FF000000")));
        if (style.border.top.style)
            border.setTop(*style.border.top.style, style.border.top.color.value_or(XLColor("FF000000")));
        if (style.border.bottom.style)
            border.setBottom(*style.border.bottom.style, style.border.bottom.color.value_or(XLColor("FF000000")));
        borderIdx = borders().findOrCreate(border);
        ignore(tentativeBorderIdx);    // tentative stays; dedup cache handles future lookups
    }

    // ===== Step 4: Resolve number format ID =================================
    uint32_t numFmtId = 0;    // default (General)
    if (style.numberFormat)
        numFmtId = createNumberFormat(*style.numberFormat);

    // ===== Step 5: Assemble the XLCellFormat and deduplicate ================
    XLStyleIndex tentativeXfIdx = cellFormats().create();
    XLCellFormat xf = cellFormats().cellFormatByIndex(tentativeXfIdx);
    xf.setFontIndex(fontIdx);
    xf.setFillIndex(fillIdx);
    xf.setBorderIndex(borderIdx);
    if (numFmtId > 0) {
        xf.setNumberFormatId(numFmtId);
        xf.setApplyNumberFormat(true);
    }
    if (fontIdx > 0)   xf.setApplyFont(true);
    if (fillIdx > 0)   xf.setApplyFill(true);
    if (borderIdx > 0) xf.setApplyBorder(true);
    if (style.alignment.horizontal || style.alignment.vertical || style.alignment.wrapText) {
        auto align = xf.alignment(XLCreateIfMissing);
        if (style.alignment.horizontal) align.setHorizontal(*style.alignment.horizontal);
        if (style.alignment.vertical)   align.setVertical(*style.alignment.vertical);
        if (style.alignment.wrapText)   align.setWrapText(*style.alignment.wrapText);
        xf.setApplyAlignment(true);
    }

    // Dedup the assembled cellXf against the full pool
    return cellFormats().findOrCreate(xf);
}

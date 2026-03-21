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

    void wrapNode(XMLNode parentNode, const XMLNode& node, std::string const& prefix)
    {
        if (not node.empty() and prefix.length() > 0) {
            parentNode.insert_child_before(pugi::node_pcdata, node).set_value(prefix.c_str());    // insert prefix before node opening tag
            XMLNode n = node;
            n.append_child(pugi::node_pcdata).set_value(prefix.c_str());    // insert prefix before node closing tag (within node)
        }
    }
}

// ===== XLStyles, master class

/**
 * @details Default constructor
 */
XLStyles::XLStyles() {}    // TBD if defaulting this constructor again would reintroduce issue #310

/**
 * @details Creates an XLStyles object, which will initialize from the given xmlData
 */
XLStyles::XLStyles(XLXmlData* xmlData, bool suppressWarnings, std::string stylesPrefix)
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
        // std::cout << "node.name() is " << node.name() << ", resulting XLStylesEntryType is " << std::to_string(e) << std::endl;
        switch (e) {
            case XLStylesNumberFormats:
                // std::cout << "found XLStylesNumberFormats, node name is " << node.name() << std::endl;
                m_numberFormats = std::make_unique<XLNumberFormats>(node);
                break;
            case XLStylesFonts:
                // std::cout << "found XLStylesFonts, node name is " << node.name() << std::endl;
                m_fonts = std::make_unique<XLFonts>(node);
                break;
            case XLStylesFills:
                // std::cout << "found XLStylesFills, node name is " << node.name() << std::endl;
                m_fills = std::make_unique<XLFills>(node);
                break;
            case XLStylesBorders:
                // std::cout << "found XLStylesBorders, node name is " << node.name() << std::endl;
                m_borders = std::make_unique<XLBorders>(node);
                break;
            case XLStylesCellStyleFormats:
                // std::cout << "found XLStylesCellStyleFormats, node name is " << node.name() << std::endl;
                m_cellStyleFormats = std::make_unique<XLCellFormats>(node);
                break;
            case XLStylesCellFormats:
                // std::cout << "found XLStylesCellFormats, node name is " << node.name() << std::endl;
                m_cellFormats = std::make_unique<XLCellFormats>(node, XLPermitXfID);
                break;
            case XLStylesCellStyles:
                // std::cout << "found XLStylesCellStyles, node name is " << node.name() << std::endl;
                m_cellStyles = std::make_unique<XLCellStyles>(node);
                break;
            case XLStylesDiffCellFormats:
                // std::cout << "found XLStylesDiffCellFormats, node name is " << node.name() << std::endl;
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

/**
 * @details move-construct an XLStyles object
 */
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

/**
 * @details copy-construct an XLStyles object
 */
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

/**
 * @details move-assign an XLStyles object
 */
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

/**
 * @details copy-assign an XLStyles object
 */
XLStyles& XLStyles::operator=(const XLStyles& other)
{
    if (&other != this) {
        XLStyles temp = other;              // copy-construct
        *this         = std::move(temp);    // move-assign & invalidate temp
    }
    return *this;
}

/**
 * @details Create a new custom number format with a unique ID and format code
 */
uint32_t XLStyles::createNumberFormat(const std::string& formatCode)
{
    return m_numberFormats->createNumberFormat(formatCode);
}

/**
 * @details return a handle to the underlying number formats
 */
XLNumberFormats& XLStyles::numberFormats() const { return *m_numberFormats; }

/**
 * @details return a handle to the underlying fonts
 */
XLFonts& XLStyles::fonts() const { return *m_fonts; }

/**
 * @details return a handle to the underlying fills
 */
XLFills& XLStyles::fills() const { return *m_fills; }

/**
 * @details return a handle to the underlying borders
 */
XLBorders& XLStyles::borders() const { return *m_borders; }

/**
 * @details return a handle to the underlying cell style formats
 */
XLCellFormats& XLStyles::cellStyleFormats() const { return *m_cellStyleFormats; }

/**
 * @details return a handle to the underlying cell formats
 */
XLCellFormats& XLStyles::cellFormats() const { return *m_cellFormats; }

/**
 * @details return a handle to the underlying cell styles
 */
XLCellStyles& XLStyles::cellStyles() const { return *m_cellStyles; }

/**
 * @details return a handle to the underlying differential cell formats
 */
XLDxfs& XLStyles::dxfs() const { return *m_dxfs; }

/**
 * @details Add a differential cell format (DXF) to the styles and return its index
 */
XLStyleIndex XLStyles::addDxf(const XLDxf& dxf) { return m_dxfs->create(dxf); }

XLStyleIndex XLStyles::addNamedStyle(const std::string& name,
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

XLStyleIndex XLStyles::namedStyle(const std::string& name) const
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
            return static_cast<XLStyleIndex>(i);
        }
    }
    
    return XLInvalidStyleIndex;
}

/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLXMLSCHEMA_HPP
#define OPENXLSX_XLXMLSCHEMA_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLXmlHelpers.hpp"
#include "XLXmlParser.hpp"

#include <algorithm>
#include <cstdint>
#include <string>
#include <string_view>

namespace OpenXLSX
{
    namespace XmlNS
    {
        inline constexpr std::string_view SpreadsheetML        = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        inline constexpr std::string_view Relationships        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        inline constexpr std::string_view PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        inline constexpr std::string_view DrawingML            = "http://schemas.openxmlformats.org/drawingml/2006/main";
        inline constexpr std::string_view SpreadsheetDrawing   = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        inline constexpr std::string_view VML                  = "urn:schemas-microsoft-com:vml";
        inline constexpr std::string_view OfficeVML            = "urn:schemas-microsoft-com:office:office";
        inline constexpr std::string_view ExcelVML             = "urn:schemas-microsoft-com:office:excel";
        inline constexpr std::string_view ContentTypes         = "http://schemas.openxmlformats.org/package/2006/content-types";
        inline constexpr std::string_view CoreProperties       = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        inline constexpr std::string_view ExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        inline constexpr std::string_view CustomProperties   = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    }    // namespace XmlNS

    namespace XmlTag
    {
        inline constexpr const char* Cell            = "c";
        inline constexpr const char* Row             = "row";
        inline constexpr const char* Value           = "v";
        inline constexpr const char* Formula         = "f";
        inline constexpr const char* InlineStr       = "is";
        inline constexpr const char* Text            = "t";
        inline constexpr const char* SheetData       = "sheetData";
        inline constexpr const char* Dimension       = "dimension";
        inline constexpr const char* MergeCells      = "mergeCells";
        inline constexpr const char* MergeCell       = "mergeCell";
        inline constexpr const char* DataValidations = "dataValidations";
        inline constexpr const char* DataValidation  = "dataValidation";
        inline constexpr const char* AutoFilter      = "autoFilter";
        inline constexpr const char* SheetView       = "sheetView";
        inline constexpr const char* SheetViews      = "sheetViews";
        inline constexpr const char* Pane            = "pane";
        inline constexpr const char* Selection       = "selection";
        inline constexpr const char* PageSetup       = "pageSetup";
        inline constexpr const char* PageMargins     = "pageMargins";
        inline constexpr const char* HeaderFooter    = "headerFooter";
        inline constexpr const char* SheetPr         = "sheetPr";
        inline constexpr const char* TabColor        = "tabColor";
        inline constexpr const char* TableParts      = "tableParts";
        inline constexpr const char* TablePart       = "tablePart";
        inline constexpr const char* Drawing         = "drawing";
        inline constexpr const char* LegacyDrawing   = "legacyDrawing";
        inline constexpr const char* Table           = "table";
        inline constexpr const char* TableColumns    = "tableColumns";
        inline constexpr const char* TableColumn     = "tableColumn";
        inline constexpr const char* TableStyleInfo  = "tableStyleInfo";
        inline constexpr const char* FilterColumn    = "filterColumn";
        inline constexpr const char* CustomFilters   = "customFilters";
        inline constexpr const char* CustomFilter    = "customFilter";
        inline constexpr const char* Filters         = "filters";
        inline constexpr const char* Filter          = "filter";
        inline constexpr const char* Comments        = "comments";
        inline constexpr const char* Authors         = "authors";
        inline constexpr const char* Author          = "author";
        inline constexpr const char* CommentList     = "commentList";
        inline constexpr const char* Comment         = "comment";
        inline constexpr const char* TextBody        = "text";
        inline constexpr const char* Col             = "col";
        inline constexpr const char* Cols            = "cols";
    }    // namespace XmlTag

    namespace XmlAttr
    {
        inline constexpr const char* Ref               = "r";
        inline constexpr const char* Type              = "t";
        inline constexpr const char* Style             = "s";
        inline constexpr const char* ColStyle          = "style";
        inline constexpr const char* Sqref             = "sqref";
        inline constexpr const char* Val               = "val";
        inline constexpr const char* Name              = "name";
        inline constexpr const char* Id                = "id";
        inline constexpr const char* SheetId           = "sheetId";
        inline constexpr const char* State             = "state";
        inline constexpr const char* Rgb               = "rgb";
        inline constexpr const char* Spans             = "spans";
        inline constexpr const char* Hidden            = "hidden";
        inline constexpr const char* CustomHeight      = "customHeight";
        inline constexpr const char* Ht                = "ht";
        inline constexpr const char* Count             = "count";
        inline constexpr const char* DisplayName       = "displayName";
        inline constexpr const char* HeaderRowCount    = "headerRowCount";
        inline constexpr const char* TotalsRowCount    = "totalsRowCount";
        inline constexpr const char* TotalsRowShown    = "totalsRowShown";
        inline constexpr const char* TotalsRowFunction = "totalsRowFunction";
        inline constexpr const char* TotalsRowLabel    = "totalsRowLabel";
        inline constexpr const char* ShowFirstColumn   = "showFirstColumn";
        inline constexpr const char* ShowLastColumn    = "showLastColumn";
        inline constexpr const char* ShowRowStripes    = "showRowStripes";
        inline constexpr const char* ShowColumnStripes = "showColumnStripes";
        inline constexpr const char* XfId              = "xfId";
        inline constexpr const char* FontId            = "fontId";
        inline constexpr const char* FillId            = "fillId";
        inline constexpr const char* BorderId          = "borderId";
        inline constexpr const char* NumFmtId          = "numFmtId";
        inline constexpr const char* AuthorId          = "authorId";
        inline constexpr const char* ShapeId           = "shapeId";
        inline constexpr const char* Min               = "min";
        inline constexpr const char* Max               = "max";
        inline constexpr const char* Width             = "width";
        inline constexpr const char* CustomWidth       = "customWidth";
        inline constexpr const char* OutlineLevel      = "outlineLevel";
        inline constexpr const char* Collapsed         = "collapsed";
        inline constexpr const char* DyDescent         = "x14ac:dyDescent";
        inline constexpr const char* WorkbookViewId    = "workbookViewId";
        inline constexpr const char* TabSelected       = "tabSelected";
        inline constexpr const char* ShowGridLines     = "showGridLines";
        inline constexpr const char* ShowRowColHeaders = "showRowColHeaders";
        inline constexpr const char* ZoomScale         = "zoomScale";
        inline constexpr const char* XSplit            = "xSplit";
        inline constexpr const char* YSplit            = "ySplit";
        inline constexpr const char* TopLeftCell       = "topLeftCell";
        inline constexpr const char* ActivePane        = "activePane";
    }    // namespace XmlAttr

    /**
     * @brief Plain POD properties describing an OOXML Table (ECMA-376 <table> element).
     */
    struct XLTableProperties
    {
        uint32_t    id{0};
        std::string name;
        std::string displayName;
        std::string rangeRef;
        uint32_t    headerRowCount{1};
        uint32_t    totalsRowCount{0};
        bool        totalsRowShown{false};

        friend bool operator==(const XLTableProperties& a, const XLTableProperties& b) noexcept
        {
            return a.id == b.id && a.name == b.name && a.displayName == b.displayName &&
                   a.rangeRef == b.rangeRef && a.headerRowCount == b.headerRowCount &&
                   a.totalsRowCount == b.totalsRowCount && a.totalsRowShown == b.totalsRowShown;
        }
        friend bool operator!=(const XLTableProperties& a, const XLTableProperties& b) noexcept { return !(a == b); }
    };

    /**
     * @brief Plain POD properties describing an OOXML Row (<row> element).
     * @details Compact value-semantic struct.
     */
    struct XLRowProperties
    {
        uint32_t rowIndex{0};         ///< Row index (1-based, 1..1048576)
        float    height{0.0f};        ///< Row height in points
        uint32_t styleIndex{0};       ///< Style index referencing cellXfs
        uint8_t  outlineLevel{0};     ///< Outline grouping level (0..7)
        bool     customHeight{false};
        bool     hidden{false};
        bool     collapsed{false};

        friend bool operator==(const XLRowProperties& a, const XLRowProperties& b) noexcept
        {
            return a.rowIndex == b.rowIndex && a.height == b.height &&
                   a.styleIndex == b.styleIndex && a.outlineLevel == b.outlineLevel &&
                   a.customHeight == b.customHeight && a.hidden == b.hidden && a.collapsed == b.collapsed;
        }
        friend bool operator!=(const XLRowProperties& a, const XLRowProperties& b) noexcept { return !(a == b); }
    };

    /**
     * @brief Plain POD properties describing an OOXML Column (<col> element).
     * @details Compact 16-byte value-semantic struct aligned with Apple architecture guidelines.
     */
    struct XLColumnProperties
    {
        uint16_t minCol{1};           ///< Start column (1-based, 1..16384)
        uint16_t maxCol{1};           ///< End column (1-based, 1..16384)
        float    width{0.0f};         ///< Column width in character units
        uint32_t styleIndex{0};       ///< Style index referencing cellXfs
        uint8_t  outlineLevel{0};     ///< Outline grouping level (0..7)
        bool     customWidth{false};
        bool     hidden{false};
        bool     collapsed{false};

        friend bool operator==(const XLColumnProperties& a, const XLColumnProperties& b) noexcept
        {
            return a.minCol == b.minCol && a.maxCol == b.maxCol && a.width == b.width &&
                   a.styleIndex == b.styleIndex && a.outlineLevel == b.outlineLevel &&
                   a.customWidth == b.customWidth && a.hidden == b.hidden && a.collapsed == b.collapsed;
        }
        friend bool operator!=(const XLColumnProperties& a, const XLColumnProperties& b) noexcept { return !(a == b); }
    };

    /**
     * @brief Value-semantic POD properties describing worksheet window panes and freeze state.
     */
    struct XLPaneProperties
    {
        uint16_t    xSplit{0};                  ///< Frozen column count (0 = no vertical split)
        uint32_t    ySplit{0};                  ///< Frozen row count (0 = no horizontal split)
        std::string topLeftCell;                ///< Top-left visible cell in active pane (e.g. "C5")
        std::string state{"frozen"};            ///< "frozen" or "split"
        std::string activePane;                 ///< "bottomRight", "topRight", or "bottomLeft"

        [[nodiscard]] bool empty() const noexcept { return xSplit == 0 && ySplit == 0; }

        friend bool operator==(const XLPaneProperties& a, const XLPaneProperties& b) noexcept
        {
            return a.xSplit == b.xSplit && a.ySplit == b.ySplit && a.topLeftCell == b.topLeftCell &&
                   a.state == b.state && a.activePane == b.activePane;
        }
        friend bool operator!=(const XLPaneProperties& a, const XLPaneProperties& b) noexcept { return !(a == b); }
    };

    namespace XmlOps
    {
        [[nodiscard]] inline XLRowProperties readRowProps(XMLNode node) noexcept
        {
            XLRowProperties props;
            if (node.empty()) return props;
            props.rowIndex     = node.attribute(XmlAttr::Ref).as_uint();
            props.height       = node.attribute(XmlAttr::Ht).as_float();
            props.customHeight = node.attribute(XmlAttr::CustomHeight).as_bool();
            props.hidden       = node.attribute(XmlAttr::Hidden).as_bool();
            props.outlineLevel = static_cast<uint8_t>(std::clamp(node.attribute(XmlAttr::OutlineLevel).as_uint(0), 0u, 7u));
            props.collapsed    = node.attribute(XmlAttr::Collapsed).as_bool();
            props.styleIndex   = node.attribute(XmlAttr::Style).as_uint();
            return props;
        }

        inline void writeRowProps(XMLNode node, const XLRowProperties& props)
        {
            if (node.empty()) return;
            if (props.rowIndex > 0) setAttr(node, XmlAttr::Ref, props.rowIndex);
            if (props.height > 0.0f || props.customHeight) {
                setAttr(node, XmlAttr::Ht, props.height);
                setAttr(node, XmlAttr::CustomHeight, props.customHeight ? "1" : "0");
            }
            if (props.hidden) {
                setAttr(node, XmlAttr::Hidden, "1");
            } else {
                node.remove_attribute(XmlAttr::Hidden);
            }
            uint8_t lvl = static_cast<uint8_t>(std::clamp(static_cast<unsigned>(props.outlineLevel), 0u, 7u));
            if (lvl > 0) {
                setAttr(node, XmlAttr::OutlineLevel, lvl);
            } else {
                node.remove_attribute(XmlAttr::OutlineLevel);
            }
            if (props.collapsed) {
                setAttr(node, XmlAttr::Collapsed, "1");
            } else {
                node.remove_attribute(XmlAttr::Collapsed);
            }
            if (props.styleIndex > 0) {
                setAttr(node, XmlAttr::Style, props.styleIndex);
            } else {
                node.remove_attribute(XmlAttr::Style);
            }
        }

        [[nodiscard]] inline XLColumnProperties readColProps(XMLNode node) noexcept
        {
            XLColumnProperties props;
            if (node.empty()) return props;
            props.minCol       = static_cast<uint16_t>(node.attribute(XmlAttr::Min).as_uint(1));
            props.maxCol       = static_cast<uint16_t>(node.attribute(XmlAttr::Max).as_uint(1));
            props.width        = node.attribute(XmlAttr::Width).as_float(0.0f);
            props.customWidth  = node.attribute(XmlAttr::CustomWidth).as_bool();
            props.hidden       = node.attribute(XmlAttr::Hidden).as_bool();
            props.outlineLevel = static_cast<uint8_t>(std::clamp(node.attribute(XmlAttr::OutlineLevel).as_uint(0), 0u, 7u));
            props.collapsed    = node.attribute(XmlAttr::Collapsed).as_bool();
            props.styleIndex   = node.attribute(XmlAttr::ColStyle).as_uint(0);
            return props;
        }

        inline void writeColProps(XMLNode node, const XLColumnProperties& props)
        {
            if (node.empty()) return;
            setAttr(node, XmlAttr::Min, props.minCol);
            setAttr(node, XmlAttr::Max, props.maxCol);
            if (props.width > 0.0f || props.customWidth) {
                setAttr(node, XmlAttr::Width, props.width);
                setAttr(node, XmlAttr::CustomWidth, props.customWidth ? "1" : "0");
            }
            if (props.hidden) {
                setAttr(node, XmlAttr::Hidden, "1");
            } else {
                node.remove_attribute(XmlAttr::Hidden);
            }
            uint8_t lvl = static_cast<uint8_t>(std::clamp(static_cast<unsigned>(props.outlineLevel), 0u, 7u));
            if (lvl > 0) {
                setAttr(node, XmlAttr::OutlineLevel, lvl);
            } else {
                node.remove_attribute(XmlAttr::OutlineLevel);
            }
            if (props.collapsed) {
                setAttr(node, XmlAttr::Collapsed, "1");
            } else {
                node.remove_attribute(XmlAttr::Collapsed);
            }
            if (props.styleIndex > 0) {
                setAttr(node, XmlAttr::ColStyle, props.styleIndex);
            } else {
                node.remove_attribute(XmlAttr::ColStyle);
            }
        }

        [[nodiscard]] inline XLPaneProperties readPaneProps(XMLNode sheetViewNode) noexcept
        {
            XLPaneProperties props;
            if (sheetViewNode.empty()) return props;
            auto pane = sheetViewNode.child(XmlTag::Pane);
            if (pane.empty()) return props;
            props.xSplit      = static_cast<uint16_t>(pane.attribute(XmlAttr::XSplit).as_uint(0));
            props.ySplit      = pane.attribute(XmlAttr::YSplit).as_uint(0);
            props.topLeftCell = pane.attribute(XmlAttr::TopLeftCell).value();
            props.state       = pane.attribute(XmlAttr::State).as_string("frozen");
            props.activePane  = pane.attribute(XmlAttr::ActivePane).value();
            return props;
        }

        inline void writePaneProps(XMLNode sheetViewNode, const XLPaneProperties& props)
        {
            if (sheetViewNode.empty()) return;
            // Clean existing pane and selection per ECMA-376 invariant
            sheetViewNode.remove_child(XmlTag::Pane);
            while (auto sel = sheetViewNode.child(XmlTag::Selection)) {
                sheetViewNode.remove_child(sel);
            }
            if (props.empty()) return;

            // 1. ECMA-376 strict sequence: <pane> MUST precede <selection>
            auto pane = sheetViewNode.prepend_child(XmlTag::Pane);
            if (props.xSplit > 0) setAttr(pane, XmlAttr::XSplit, props.xSplit);
            if (props.ySplit > 0) setAttr(pane, XmlAttr::YSplit, props.ySplit);
            if (!props.topLeftCell.empty()) setAttr(pane, XmlAttr::TopLeftCell, props.topLeftCell);
            if (!props.state.empty()) setAttr(pane, XmlAttr::State, props.state);
            if (!props.activePane.empty()) setAttr(pane, XmlAttr::ActivePane, props.activePane);

            // 2. <selection> appended after <pane>
            if (props.xSplit > 0 && props.ySplit > 0) {
                auto sel1 = sheetViewNode.insert_child_after(XmlTag::Selection, pane);
                setAttr(sel1, XmlTag::Pane, "topRight");
                auto sel2 = sheetViewNode.insert_child_after(XmlTag::Selection, sel1);
                setAttr(sel2, XmlTag::Pane, "bottomLeft");
            }
        }

        [[nodiscard]] inline XLTableProperties readTableProps(XMLNode node)
        {
            XLTableProperties props;
            if (node.empty()) return props;
            props.id             = node.attribute(XmlAttr::Id).as_uint();
            props.name           = node.attribute(XmlAttr::Name).value();
            props.displayName    = node.attribute(XmlAttr::DisplayName).value();
            props.rangeRef       = node.attribute(XmlAttr::Ref).value();
            props.headerRowCount = node.attribute(XmlAttr::HeaderRowCount).as_uint(1);
            props.totalsRowCount = node.attribute(XmlAttr::TotalsRowCount).as_uint(0);
            props.totalsRowShown = node.attribute(XmlAttr::TotalsRowShown).as_bool();
            return props;
        }

        inline void writeTableProps(XMLNode node, const XLTableProperties& props)
        {
            if (node.empty()) return;
            if (props.id > 0) setAttr(node, XmlAttr::Id, props.id);
            if (!props.name.empty()) setAttr(node, XmlAttr::Name, props.name);
            if (!props.displayName.empty()) setAttr(node, XmlAttr::DisplayName, props.displayName);
            if (!props.rangeRef.empty()) setAttr(node, XmlAttr::Ref, props.rangeRef);
            setAttr(node, XmlAttr::HeaderRowCount, props.headerRowCount);
            if (props.totalsRowCount > 0) setAttr(node, XmlAttr::TotalsRowCount, props.totalsRowCount);
            if (props.totalsRowShown) setAttr(node, XmlAttr::TotalsRowShown, "1");
        }
    }    // namespace XmlOps

    /**
     * @brief Zero-overhead typed view proxy over a raw XML cell node (<c>).
     * @details Provides direct semantic accessors avoiding raw strings and repeatedly
     *          re-scanning node attribute linked lists from scratch.
     */
    class CellNodeView
    {
    public:
        explicit CellNodeView(XMLNode node) noexcept : m_node(node) {}

        [[nodiscard]] bool    empty() const noexcept { return m_node.empty(); }
        [[nodiscard]] XMLNode node() const noexcept { return m_node; }

        [[nodiscard]] const char* ref() const noexcept { return m_node.attribute(XmlAttr::Ref).value(); }

        [[nodiscard]] const char* type() const noexcept { return m_node.attribute(XmlAttr::Type).value(); }

        [[nodiscard]] uint32_t styleIndex() const noexcept { return m_node.attribute(XmlAttr::Style).as_uint(); }

        [[nodiscard]] const char* value() const noexcept
        {
            auto v = m_node.child(XmlTag::Value);
            return v.text().get();
        }

        [[nodiscard]] const char* formula() const noexcept
        {
            auto f = m_node.child(XmlTag::Formula);
            return f.text().get();
        }

        [[nodiscard]] bool hasFormula() const noexcept { return !m_node.child(XmlTag::Formula).empty(); }

    private:
        XMLNode m_node;
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLXMLSCHEMA_HPP

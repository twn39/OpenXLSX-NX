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
#include "XLXmlParser.hpp"

#include <cstdint>
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
        inline constexpr std::string_view ExtendedProperties   = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        inline constexpr std::string_view CustomProperties     = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    } // namespace XmlNS

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
    } // namespace XmlTag

    namespace XmlAttr
    {
        inline constexpr const char* Ref          = "r";
        inline constexpr const char* Type         = "t";
        inline constexpr const char* Style        = "s";
        inline constexpr const char* Sqref        = "sqref";
        inline constexpr const char* Val          = "val";
        inline constexpr const char* Name         = "name";
        inline constexpr const char* Id           = "id";
        inline constexpr const char* SheetId      = "sheetId";
        inline constexpr const char* State        = "state";
        inline constexpr const char* Rgb          = "rgb";
        inline constexpr const char* Spans        = "spans";
        inline constexpr const char* Hidden       = "hidden";
        inline constexpr const char* CustomHeight = "customHeight";
        inline constexpr const char* Ht           = "ht";
    } // namespace XmlAttr

    /**
     * @brief Zero-overhead typed view proxy over a raw XML cell node (<c>).
     * @details Provides direct semantic accessors avoiding raw strings and repeatedly
     *          re-scanning node attribute linked lists from scratch.
     */
    class CellNodeView
    {
    public:
        explicit CellNodeView(XMLNode node) noexcept : m_node(node) {}

        [[nodiscard]] bool empty() const noexcept { return m_node.empty(); }
        [[nodiscard]] XMLNode node() const noexcept { return m_node; }

        [[nodiscard]] const char* ref() const noexcept
        {
            return m_node.attribute(XmlAttr::Ref).value();
        }

        [[nodiscard]] const char* type() const noexcept
        {
            return m_node.attribute(XmlAttr::Type).value();
        }

        [[nodiscard]] uint32_t styleIndex() const noexcept
        {
            return m_node.attribute(XmlAttr::Style).as_uint();
        }

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

        [[nodiscard]] bool hasFormula() const noexcept
        {
            return !m_node.child(XmlTag::Formula).empty();
        }

    private:
        XMLNode m_node;
    };

} // namespace OpenXLSX

#endif // OPENXLSX_XLXMLSCHEMA_HPP

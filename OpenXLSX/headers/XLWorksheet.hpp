#include <optional>

#ifndef OPENXLSX_XLWORKSHEET_HPP
#define OPENXLSX_XLWORKSHEET_HPP

#include <vector>
#include <string>
#include <string_view>
#include "OpenXLSX-Exports.hpp"
#include "XLSheetBase.hpp"
#include "XLConstants.hpp"

#include "XLCell.hpp"
#include "XLCellRange.hpp"
#include "XLCellReference.hpp"
#include "XLRow.hpp"
#include "XLColumn.hpp"
#include "XLMergeCells.hpp"
#include "XLDataValidation.hpp"
#include "XLTables.hpp"
#include "XLDrawing.hpp"
#include "XLComments.hpp"
#include "XLRelationships.hpp"
#include "XLConditionalFormatting.hpp"
#include "XLPageSetup.hpp"
#include "XLChart.hpp"
#include "XLStreamWriter.hpp"
#include "XLStreamReader.hpp"
#include "XLPivotTable.hpp"
#include "XLImageOptions.hpp"

namespace OpenXLSX
{
        const std::vector<std::string_view> XLWorksheetNodeOrder = {    // worksheet XML root node required child sequence
        "sheetPr",
        "dimension",
        "sheetViews",
        "sheetFormatPr",
        "cols",
        "sheetData",
        "sheetCalcPr",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "autoFilter",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
        "mergeCells",
        "phoneticPr",
        "conditionalFormatting",
        "dataValidations",
        "hyperlinks",
        "printOptions",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "rowBreaks",
        "colBreaks",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "smartTags",
        "drawing",
        "legacyDrawing",
        "legacyDrawingHF",
        "drawingHF",
        "picture",
        "oleObjects",
        "controls",
        "webPublishItems",
        "tableParts",
        "pivotTables",
        "extLst"
    };

    const std::vector<std::string_view> XLSheetViewNodeOrder = {    // worksheet XML <sheetView> child sequence
        "pane",
        "selection",
        "pivotSelection",
        "extLst"};

    /**
     * @brief A class encapsulating an Excel worksheet. Access to XLWorksheet objects should be via the workbook object.
     */
    class OPENXLSX_EXPORT XLWorksheet final : public XLSheetBase<XLWorksheet>
    {
        friend class XLCell;
        friend class XLRow;
        friend class XLWorkbook;
        friend class XLTableCollection;
        friend class XLSheetBase<XLWorksheet>;

    public:
        XLWorksheet() : XLSheetBase(nullptr) {};
        explicit XLWorksheet(XLXmlData* xmlData);
        ~XLWorksheet() = default;
        XLWorksheet(const XLWorksheet& other);
        XLWorksheet(XLWorksheet&& other);
        XLWorksheet& operator=(const XLWorksheet& other);
        XLWorksheet& operator=(XLWorksheet&& other);

        XLCellAssignable cell(const std::string& ref) const;
        XLCellAssignable cell(const XLCellReference& ref) const;
        XLCellAssignable cell(uint32_t rowNumber, uint16_t columnNumber) const;
        XLCellAssignable cell(XLRowIndex row, XLColIndex col) const { return cell(row.val, col.val); }

        /**
         * @brief Starts a high-performance, low-memory stream writer for this worksheet.
         * @warning Initiating a stream writer locks the DOM for this sheet.
         */
        XLStreamWriter streamWriter();

        /**
         * @brief Create a stream reader for memory efficient reading of large documents
         * @return An XLStreamReader object.
         */
        XLStreamReader streamReader() const;


        XLCellAssignable findCell(const std::string& ref) const;
        XLCellAssignable findCell(const XLCellReference& ref) const;
        XLCellAssignable findCell(uint32_t rowNumber, uint16_t columnNumber) const;

        std::optional<XLCell> peekCell(const std::string& ref) const;
        std::optional<XLCell> peekCell(const XLCellReference& ref) const;
        std::optional<XLCell> peekCell(uint32_t rowNumber, uint16_t columnNumber) const;
        std::optional<XLCell> peekCell(XLRowIndex row, XLColIndex col) const { return peekCell(row.val, col.val); }



        XLCellRange range() const;
        XLCellRange range(const XLCellReference& topLeft, const XLCellReference& bottomRight) const;
        XLCellRange range(std::string const& topLeft, std::string const& bottomRight) const;
        XLCellRange range(std::string const& rangeReference) const;

        XLRowRange rows() const;
        XLRowRange rows(uint32_t rowCount) const;
        XLRowRange rows(uint32_t firstRow, uint32_t lastRow) const;

        XLRow row(uint32_t rowNumber) const;

        /**
         * @brief Append a new row with the given values.
         * @param values A vector of XLCellValue items.
         */
        void appendRow(const std::vector<XLCellValue>& values);

        /**
         * @brief Append a new row with the given container of values.
         * @tparam T A container type.
         * @param values The container of values.
         */
        template<typename T,
                 typename = std::enable_if_t<!std::is_same_v<T, std::vector<XLCellValue>> &&
                                                 std::is_base_of_v<std::bidirectional_iterator_tag,
                                                                   typename std::iterator_traits<typename T::iterator>::iterator_category>,
                                             T>>
        void appendRow(const T& values)
        {
            row(rowCount() + 1).values() = values;
        }

        XLColumn column(uint16_t columnNumber) const;
        XLColumn column(std::string const& columnRef) const;

        /**
         * @brief Automatically adjusts the column width based on the content of its cells.
         */
        void autoFitColumn(uint16_t columnNumber);

        void groupRows(uint32_t rowFirst, uint32_t rowLast, uint8_t outlineLevel = 1, bool collapsed = false);
        void groupColumns(uint16_t colFirst, uint16_t colLast, uint8_t outlineLevel = 1, bool collapsed = false);

        void setAutoFilter(const XLCellRange& range);
        void clearAutoFilter();
        bool hasAutoFilter() const;
        std::string autoFilter() const;
        XLAutoFilter autofilterObject() const;

        XLCellReference lastCell() const noexcept;
        uint16_t columnCount() const noexcept;
        uint32_t rowCount() const noexcept;

        bool deleteRow(uint32_t rowNumber);
        void updateSheetName(const std::string& oldName, const std::string& newName);
        void updateDimension();

        XLMergeCells& merges();
        XLDataValidations& dataValidations();

        XLCellRange mergeCells(XLCellRange const& rangeToMerge, bool emptyHiddenCells = false);
        XLCellRange mergeCells(const std::string& rangeReference, bool emptyHiddenCells = false);
        void unmergeCells(XLCellRange const& rangeToMerge);
        void unmergeCells(const std::string& rangeReference);

        XLStyleIndex getColumnFormat(uint16_t column) const;
        XLStyleIndex getColumnFormat(const std::string& column) const;
        bool setColumnFormat(uint16_t column, XLStyleIndex cellFormatIndex);
        bool setColumnFormat(const std::string& column, XLStyleIndex cellFormatIndex);

        XLStyleIndex getRowFormat(uint16_t row) const;
        bool setRowFormat(uint32_t row, XLStyleIndex cellFormatIndex);

        XLConditionalFormats conditionalFormats() const;
        void addConditionalFormatting(const std::string& sqref, const XLCfRule& rule);
        void addConditionalFormatting(const std::string& sqref, const XLCfRule& rule, const XLDxf& dxf);
        void addConditionalFormatting(const XLCellRange& range, const XLCfRule& rule);
        void addConditionalFormatting(const XLCellRange& range, const XLCfRule& rule, const XLDxf& dxf);

        void removeConditionalFormatting(const std::string& sqref);
        void removeConditionalFormatting(const XLCellRange& range);
        void clearAllConditionalFormatting();

        XLPageMargins pageMargins() const;
        XLPrintOptions printOptions() const;
        XLPageSetup pageSetup() const;
        XLHeaderFooter headerFooter() const;


        /**
         * @brief Set the print area for this worksheet.
         * @param sqref The cell range (e.g., "A1:F50").
         */
        void setPrintArea(const std::string& sqref);

        /**
         * @brief Set rows to repeat at the top of each printed page.
         * @param firstRow The first row to repeat (1-based).
         * @param lastRow The last row to repeat (1-based).
         */
        void setPrintTitleRows(uint32_t firstRow, uint32_t lastRow);

        /**
         * @brief Set columns to repeat at the left of each printed page.
         * @param firstCol The first column to repeat (1-based).
         * @param lastCol The last column to repeat (1-based).
         */
        void setPrintTitleCols(uint16_t firstCol, uint16_t lastCol);

        bool protectSheet(bool set = true);
        bool protectObjects(bool set = true);
        bool protectScenarios(bool set = true);

        bool allowInsertColumns(bool set = true);
        bool allowInsertRows(bool set = true);
        bool allowDeleteColumns(bool set = true);
        bool allowDeleteRows(bool set = true);
        bool allowFormatCells(bool set = true);
        bool allowFormatColumns(bool set = true);
        bool allowFormatRows(bool set = true);
        bool allowInsertHyperlinks(bool set = true);
        bool allowSort(bool set = true);
        bool allowAutoFilter(bool set = true);
        bool allowPivotTables(bool set = true);
        bool allowSelectLockedCells(bool set = true);
        bool allowSelectUnlockedCells(bool set = true);

        bool denyInsertColumns() { return allowInsertColumns(false); }
        bool denyInsertRows() { return allowInsertRows(false); }
        bool denyDeleteColumns() { return allowDeleteColumns(false); }
        bool denyDeleteRows() { return allowDeleteRows(false); }
        bool denyFormatCells() { return allowFormatCells(false); }
        bool denyFormatColumns() { return allowFormatColumns(false); }
        bool denyFormatRows() { return allowFormatRows(false); }
        bool denyInsertHyperlinks() { return allowInsertHyperlinks(false); }
        bool denySort() { return allowSort(false); }
        bool denyAutoFilter() { return allowAutoFilter(false); }
        bool denyPivotTables() { return allowPivotTables(false); }
        bool denySelectLockedCells() { return allowSelectLockedCells(false); }
        bool denySelectUnlockedCells() { return allowSelectUnlockedCells(false); }

        bool setPasswordHash(std::string hash);
        bool setPassword(std::string password);
        bool clearPassword();
        bool clearSheetProtection();

        bool sheetProtected() const;
        bool objectsProtected() const;
        bool scenariosProtected() const;

        bool insertColumnsAllowed() const;
        bool insertRowsAllowed() const;
        bool deleteColumnsAllowed() const;
        bool deleteRowsAllowed() const;
        bool formatCellsAllowed() const;
        bool formatColumnsAllowed() const;
        bool formatRowsAllowed() const;
        bool insertHyperlinksAllowed() const;
        bool sortAllowed() const;
        bool autoFilterAllowed() const;
        bool pivotTablesAllowed() const;
        bool selectLockedCellsAllowed() const;
        bool selectUnlockedCellsAllowed() const;

        std::string passwordHash() const;
        bool passwordIsSet() const;
        std::string sheetProtectionSummary() const;

        bool hasRelationships() const;
        bool hasDrawing() const;
        bool hasVmlDrawing() const;
        bool hasComments() const;
        bool hasTables() const;

        XLDrawing& drawing();
        void addImage(const std::string& name, const std::string& data, uint32_t row, uint32_t col, uint32_t width, uint32_t height);

        /**
         * @brief Insert an image at the specified cell with original dimensions (automatic parsing).
         * @param cellReference The top-left cell where the image should be anchored (e.g. "B2").
         * @param imagePath The file path to the local image.
         */
        void insertImage(const std::string& cellReference, const std::string& imagePath);

        /**
         * @brief Insert an image with advanced formatting options (scaling, offset, behavior).
         * @param cellReference The top-left cell where the image should be anchored (e.g. "B2").
         * @param imagePath The file path to the local image.
         * @param options An XLImageOptions struct specifying formatting and behavior.
         */
        void insertImage(const std::string& cellReference, const std::string& imagePath, const XLImageOptions& options);
        void addScaledImage(const std::string& name, const std::string& data, uint32_t row, uint32_t col, double scalingFactor = 1.0);
        std::vector<XLDrawingItem> images();
        
        XLChart addChart(XLChartType type, std::string_view name, uint32_t row, uint32_t col, uint32_t width, uint32_t height);
        XLChart addChart(XLChartType type, const XLChartAnchor& anchor);


        /**
         * @brief Create and add a pivot table to this worksheet.
         * @param options The configuration options for the pivot table.
         * @return The created XLPivotTable object.
         */
        XLPivotTable addPivotTable(const XLPivotTableOptions& options);


        XLVmlDrawing& vmlDrawing();
        XLComments& comments();
        /**
         * @brief Add a comment to a cell seamlessly.
         * @param cellRef The cell reference (e.g. A1).
         * @param text The comment text.
         * @param author The author of the comment.
         */
        void addComment(const std::string& cellRef, const std::string& text, const std::string& author = "");

        XLTableCollection& tables();

        void addHyperlink(std::string_view cellRef, std::string_view url, std::string_view tooltip = "");
        void addInternalHyperlink(std::string_view cellRef, std::string_view location, std::string_view tooltip = "");
        [[nodiscard]] bool hasHyperlink(std::string_view cellRef) const;
        [[nodiscard]] std::string getHyperlink(std::string_view cellRef) const;
        void removeHyperlink(std::string_view cellRef);

        [[nodiscard]] bool hasPanes() const;
        void freezePanes(uint16_t column, uint32_t row);
        
        /**
         * @brief Freeze panes at the specified top-left cell.
         * This is a highly ergonomic method. E.g. freezePanes("B2") freezes Row 1 and Column A.
         * @param topLeftCell The cell reference where the scrolling data should start.
         */
        void freezePanes(const std::string& topLeftCell);
        
        void splitPanes(double xSplit, double ySplit, std::string_view topLeftCell = "", XLPane activePane = XLPane::BottomRight);
        void clearPanes();

        void setZoom(uint16_t scale);
        [[nodiscard]] uint16_t zoom() const;

        void insertRowBreak(uint32_t row);
        void removeRowBreak(uint32_t row);

        [[nodiscard]] static std::string makeInternalLocation(std::string_view sheetName, std::string_view cellRef);

    private:
        uint16_t sheetXmlNumber() const;
        XLRelationships& relationships();

        XLColor getColor_impl() const;
        void setColor_impl(const XLColor& color);
        bool isSelected_impl() const;
        void setSelected_impl(bool selected);
        bool isActive_impl() const;
        bool setActive_impl();

        XMLNode prepareSheetViewForPanes();

    private:
        XLRelationships   m_relationships{};
        XLMergeCells      m_merges{};
        XLDataValidations m_dataValidations{};
        XLDrawing         m_drawing{};
        XLVmlDrawing      m_vmlDrawing{};
        XLComments        m_comments{};
        XLTableCollection m_tables{};
        inline static const std::vector<std::string_view>& m_nodeOrder = XLWorksheetNodeOrder;
    };
}

#endif

#include <optional>
#include <memory>

#ifndef OPENXLSX_XLWORKSHEET_HPP
#    define OPENXLSX_XLWORKSHEET_HPP

#    include "OpenXLSX-Exports.hpp"
#    include "XLConstants.hpp"
#    include "XLSheetBase.hpp"
#    include <string>
#    include <string_view>
#    include <vector>

#    include "XLCell.hpp"
#    include "XLCellRange.hpp"
#    include "XLCellReference.hpp"
#    include "XLChart.hpp"
#    include "XLColumn.hpp"
#    include "XLConditionalFormatting.hpp"
#    include "XLPageSetup.hpp"
#    include "XLRow.hpp"
#    include "XLSparkline.hpp"

namespace OpenXLSX
{
    struct XLWorksheetImpl;
    class XLComments;
    class XLDataValidations;
    class XLDrawing;
    class XLVmlDrawing;
    class XLMergeCells;
    class XLPivotTable;
    class XLPivotTableOptions;
    struct XLSlicerOptions;
    class XLRelationships;
    class XLStreamReader;
    class XLStreamWriter;
    class XLTableCollection;
    class XLThreadedComments;

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
        "extLst"};

    const std::vector<std::string_view> XLSheetViewNodeOrder = {    // worksheet XML <sheetView> child sequence
        "pane",
        "selection",
        "pivotSelection",
        "extLst"};

    /**
     * @brief A structure defining all granular sheet protection options.
     * @details Default values are aligned with the OOXML standard for a protected sheet.
     */
    struct OPENXLSX_EXPORT XLSheetProtectionOptions {
        bool sheet = true;
        bool objects = false;
        bool scenarios = false;

        bool formatCells = false;
        bool formatColumns = false;
        bool formatRows = false;
        bool insertColumns = false;
        bool insertRows = false;
        bool insertHyperlinks = false;
        bool deleteColumns = false;
        bool deleteRows = false;
        bool sort = false;
        bool autoFilter = false;
        bool pivotTables = false;

        bool selectLockedCells = true;
        bool selectUnlockedCells = true;
    };

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
        XLWorksheet();
        explicit XLWorksheet(XLXmlData* xmlData);
        ~XLWorksheet();
        XLWorksheet(const XLWorksheet& other);
        XLWorksheet(XLWorksheet&& other) noexcept;
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
        { row(rowCount() + 1).values() = values; }

        XLColumn column(uint16_t columnNumber) const;
        XLColumn column(std::string const& columnRef) const;

        /**
         * @brief Automatically adjusts the column width based on the content of its cells.
         */
        void autoFitColumn(uint16_t columnNumber);

        void groupRows(uint32_t rowFirst, uint32_t rowLast, uint8_t outlineLevel = 1, bool collapsed = false);
        void groupColumns(uint16_t colFirst, uint16_t colLast, uint8_t outlineLevel = 1, bool collapsed = false);

        void         setAutoFilter(const XLCellRange& range);
        void         clearAutoFilter();
        bool         hasAutoFilter() const;
        std::string  autoFilter() const;
        XLAutoFilter autofilterObject() const;

        /**
         * @brief Set the sort state for the worksheet.
         * @param ref The cell range to apply the sort state to (e.g. "A1:C10").
         * @param colId The 0-based column index relative to the range to sort by.
         * @param descending True for descending order, false for ascending.
         */
        void addSortCondition(const std::string& ref, uint16_t colId, bool descending = false);

        /**
         * @brief Evaluates the current autoFilter conditions and hides rows that do not match the criteria.
         * Note: Only basic text/value matching and some simple custom filters are evaluated. Complex dynamic filters might not be
         * supported.
         */
        void applyAutoFilter();

        XLCellReference lastCell() const noexcept;
        uint16_t        columnCount() const noexcept;
        uint32_t        rowCount() const noexcept;

        bool deleteRow(uint32_t rowNumber);

        /**
         * @brief Insert one or more empty rows before the given row number.
         * @details All rows at rowNumber and below are shifted down. Merged ranges,
         *          formulas, drawing anchors, data validations and the autoFilter
         *          reference are all updated to reflect the new layout.
         * @param rowNumber 1-based row to insert before.
         * @param count     Number of rows to insert (default 1).
         * @return true on success.
         */
        bool insertRow(uint32_t rowNumber, uint32_t count = 1);

        /**
         * @brief Delete one or more rows and shift subsequent rows up.
         * @details All subsystems (merges, formulas, drawings, validations, autoFilter)
         *          are updated automatically.
         * @param rowNumber 1-based first row to delete.
         * @param count     Number of consecutive rows to delete (default 1).
         * @return true on success.
         */
        bool deleteRow(uint32_t rowNumber, uint32_t count);

        /**
         * @brief Insert one or more empty columns before the given column number.
         * @param colNumber 1-based column to insert before.
         * @param count     Number of columns to insert (default 1).
         * @return true on success.
         */
        bool insertColumn(uint16_t colNumber, uint16_t count = 1);

        /**
         * @brief Delete one or more columns and shift subsequent columns left.
         * @param colNumber 1-based first column to delete.
         * @param count     Number of consecutive columns to delete (default 1).
         * @return true on success.
         */
        bool deleteColumn(uint16_t colNumber, uint16_t count = 1);
        void updateSheetName(const std::string& oldName, const std::string& newName);
        void updateDimension();

        XLMergeCells&      merges();
        XLDataValidations& dataValidations();

        XLCellRange mergeCells(XLCellRange const& rangeToMerge, bool emptyHiddenCells = false);
        XLCellRange mergeCells(const std::string& rangeReference, bool emptyHiddenCells = false);
        void        unmergeCells(XLCellRange const& rangeToMerge);
        void        unmergeCells(const std::string& rangeReference);

        XLStyleIndex getColumnFormat(uint16_t column) const;
        XLStyleIndex getColumnFormat(const std::string& column) const;
        bool         setColumnFormat(uint16_t column, XLStyleIndex cellFormatIndex);
        bool         setColumnFormat(const std::string& column, XLStyleIndex cellFormatIndex);

        XLStyleIndex getRowFormat(uint16_t row) const;
        bool         setRowFormat(uint32_t row, XLStyleIndex cellFormatIndex);

        XLConditionalFormats conditionalFormats() const;
        void                 addConditionalFormatting(const std::string& sqref, const XLCfRule& rule);
        void                 addConditionalFormatting(const std::string& sqref, const XLCfRule& rule, const XLDxf& dxf);
        void                 addConditionalFormatting(const XLCellRange& range, const XLCfRule& rule);
        void                 addConditionalFormatting(const XLCellRange& range, const XLCfRule& rule, const XLDxf& dxf);

        void removeConditionalFormatting(const std::string& sqref);
        void removeConditionalFormatting(const XLCellRange& range);
        void clearAllConditionalFormatting();

        XLPageMargins  pageMargins() const;
        XLPrintOptions printOptions() const;
        XLPageSetup    pageSetup() const;
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

        /**
         * @brief Protect the worksheet with granular options.
         * @param options A structure containing specific protection flags.
         * @param password Optional password to protect the sheet.
         * @return true if successful.
         */
        bool protect(const XLSheetProtectionOptions& options, std::string_view password = "");

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
        bool        passwordIsSet() const;
        std::string sheetProtectionSummary() const;

        /**
         * @brief Add a sparkline to the worksheet.
         * @param location The cell where the sparkline is displayed (e.g., "A1").
         * @param dataRange The cell range containing the data (e.g., "B1:Z1").
         * @param type The type of the sparkline (Line, Column, Stacked).
         */
        void addSparkline(const std::string& location, const std::string& dataRange, XLSparklineType type = XLSparklineType::Line);
        
        /**
         * @brief Add a sparkline graph with advanced configuration options.
         * @param location The cell or range where the sparkline will be rendered (e.g. "F1").
         * @param dataRange The source data range for the sparkline (e.g. "A1:E1").
         * @param options A structure containing extensive visual formatting options.
         */
        void addSparkline(const std::string& location, const std::string& dataRange, const XLSparklineOptions& options);

        bool hasRelationships() const;
        bool hasDrawing() const;
        bool hasVmlDrawing() const;
        bool hasComments() const;
        bool hasThreadedComments() const;
        bool hasTables() const;

        XLDrawing& drawing();
        void       addImage(const std::string&    name,
                            const std::string&    data,
                            uint32_t              row,
                            uint32_t              col,
                            uint32_t              width,
                            uint32_t              height,
                            const XLImageOptions& options = XLImageOptions());

        /**
         * @brief Insert an image at the specified cell with original dimensions (automatic parsing).
         * @param cellReference The top-left cell where the image should be anchored (e.g. "B2").
         * @param imagePath The file path to the local image.
         */

        void addShape(std::string_view cellReference, const XLVectorShapeOptions& options);

        /**
         * @brief Add a slicer connected to a table.
         * @param cellReference The top-left cell where the slicer should be anchored (e.g. "D1").
         * @param table The table to filter.
         * @param columnName The exact name of the table column to filter.
         * @param options The visual and naming options for the slicer.
         */
        void addTableSlicer(std::string_view       cellReference,
                            const XLTable&         table,
                            std::string_view       columnName,
                            const XLSlicerOptions& options = XLSlicerOptions());

        /**
         * @brief Add a slicer connected to a Pivot Table.
         * @param cellReference The top-left cell where the slicer should be anchored.
         * @param pivotTable The Pivot Table to filter.
         * @param columnName The exact name of the pivot field (source column name) to filter.
         * @param options The visual and naming options for the slicer.
         */
        void addPivotSlicer(std::string_view       cellReference,
                            const XLPivotTable&    pivotTable,
                            std::string_view       columnName,
                            const XLSlicerOptions& options = XLSlicerOptions());

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
         * @brief Retrieves all pivot tables belonging to this worksheet.
         * @return A vector of XLPivotTable objects.
         */
        std::vector<XLPivotTable> pivotTables();

        /**
         * @brief Create and add a pivot table to this worksheet.
         * @param options The configuration options for the pivot table.
         * @return The created XLPivotTable object.
         */
        XLPivotTable addPivotTable(const XLPivotTableOptions& options);

        /**
         * @brief Delete a pivot table by name from this worksheet.
         * @param name The name of the pivot table to delete.
         * @return True if successfully deleted, false if not found.
         */
        bool deletePivotTable(std::string_view name);

        XLVmlDrawing& vmlDrawing();
        XLComments&   comments();
        XLThreadedComments& threadedComments();
        /**
         * @brief Add a comment to a cell seamlessly.
         * @param cellRef The cell reference (e.g. A1).
         * @param text The comment text.
         * @param author The author of the comment.
         */
        void addNote(std::string_view cellRef, std::string_view text, std::string_view author = "");

        /**
         * @brief Remove a modern threaded comment (and its replies) from a cell.
         * @param cellRef The cell reference (e.g. A1).
         */
        void deleteComment(std::string_view cellRef);

        /**
         * @brief Remove a legacy yellow note from a cell.
         * @param cellRef The cell reference (e.g. A1).
         */
        void deleteNote(std::string_view cellRef);

        /**
         * @brief Add a modern threaded comment to a cell seamlessly.
         * @param cellRef The cell reference (e.g. A1).
         * @param text The comment text.
         * @param author The author of the comment.
         */
        XLThreadedComment addComment(std::string_view cellRef, std::string_view text, std::string_view author = "");

        /**
         * @brief Add a reply to an existing threaded comment.
         * @param parentId The parent threaded comment ID.
         * @param text The comment text.
         * @param author The author of the comment.
         */
        XLThreadedComment addReply(const std::string& parentId, const std::string& text, const std::string& author = "");

        XLTableCollection& tables();

        void                      addHyperlink(std::string_view cellRef, std::string_view url, std::string_view tooltip = "");
        void                      addInternalHyperlink(std::string_view cellRef, std::string_view location, std::string_view tooltip = "");
        [[nodiscard]] bool        hasHyperlink(std::string_view cellRef) const;
        [[nodiscard]] std::string getHyperlink(std::string_view cellRef) const;
        void                      removeHyperlink(std::string_view cellRef);

        [[nodiscard]] bool hasPanes() const;
        void               freezePanes(uint16_t column, uint32_t row);

        /**
         * @brief Freeze panes at the specified top-left cell.
         * This is a highly ergonomic method. E.g. freezePanes("B2") freezes Row 1 and Column A.
         * @param topLeftCell The cell reference where the scrolling data should start.
         */
        void freezePanes(const std::string& topLeftCell);

        void splitPanes(double xSplit, double ySplit, std::string_view topLeftCell = "", XLPane activePane = XLPane::BottomRight);
        void clearPanes();

        void                   setZoom(uint16_t scale);
        [[nodiscard]] uint16_t zoom() const;

        void insertRowBreak(uint32_t row);
        void insertColBreak(uint16_t col);
        void removeRowBreak(uint32_t row);
        void removeColBreak(uint16_t col);

        void                   setSheetViewMode(std::string_view mode);
        [[nodiscard]] std::string sheetViewMode() const;

        void                   setShowGridLines(bool show);
        [[nodiscard]] bool     showGridLines() const;

        void                   setShowRowColHeaders(bool show);
        [[nodiscard]] bool     showRowColHeaders() const;

        void fitToPages(uint32_t fitToWidth, uint32_t fitToHeight);

        [[nodiscard]] static std::string makeInternalLocation(std::string_view sheetName, std::string_view cellRef);

    private:
        uint16_t         sheetXmlNumber() const;
        XLRelationships& relationships();

        XLColor getColor_impl() const;
        void    setColor_impl(const XLColor& color);
        bool    isSelected_impl() const;
        void    setSelected_impl(bool selected);
        bool    isActive_impl() const;
        bool    setActive_impl();

        XMLNode prepareSheetViewForPanes();

        // ── Row/Column structural shift helpers ──────────────────────────────
        // Each helper updates one subsystem for a row or column shift.
        // rowDelta / colDelta > 0 means insert (push out); < 0 means delete (pull in).
        // fromRow / fromCol is the 1-based first affected coordinate.

        /// Rewrite a single cell address string (e.g. "B3") applying row/col offsets.
        /// Absolute-reference components (prefixed with '$') are left unchanged.
        [[nodiscard]] static std::string
            shiftCellRef(std::string_view ref, int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol);

        /// Tokenize a formula and rewrite every CellRef/Range token using shiftCellRef.
        [[nodiscard]] static std::string
            shiftFormulaRefs(std::string_view formula, int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol);

        /// Shift <row r=N> and <c r="XN"> attributes in sheetData.
        void shiftSheetDataRows(int32_t delta, uint32_t fromRow);

        /// Shift column letters in all cell refs and update the <cols> node.
        void shiftSheetDataCols(int32_t delta, uint16_t fromCol);

        /// Update min/max attributes of each <col> element inside <cols>.
        void shiftColsNode(int32_t delta, uint16_t fromCol);

        /// Rewrite all formula cells in the sheet.
        void shiftFormulas(int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol);

        /// Adjust xdr:oneCellAnchor / xdr:twoCellAnchor row+col indices in the drawing.
        void shiftDrawingAnchors(int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol);

        /// Update sqref attributes in every <dataValidation> node.
        void shiftDataValidations(int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol);

        /// Update the <autoFilter ref="…"> attribute if present.
        void shiftAutoFilter(int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol);

    private:
        std::unique_ptr<XLWorksheetImpl> m_impl;
        inline static const std::vector<std::string_view>& m_nodeOrder = XLWorksheetNodeOrder;

        // O(1) Hint Cache for cell node DOM traversal
        mutable uint32_t                                   m_hintRowNumber{0};
        mutable XMLNode                                    m_hintRowNode{};
        mutable uint16_t                                   m_hintColNumber{0};
        mutable XMLNode                                    m_hintCellNode{};
    };
}    // namespace OpenXLSX

#endif

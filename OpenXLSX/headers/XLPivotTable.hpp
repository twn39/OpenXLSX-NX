#ifndef OPENXLSX_XLPIVOTTABLE_HPP
#define OPENXLSX_XLPIVOTTABLE_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

namespace OpenXLSX
{
    enum class XLPivotSubtotal { Sum, Average, Count, Max, Min, Product, CountNums, StdDev, StdDevP, Var, VarP };

    /**
     * @brief "Show Values As" calculation for pivot data fields (OOXML showDataAs / x14:dataField).
     * @details Types marked x14 are written via extLst (Excel 2010+). Classic types use @showDataAs.
     */
    enum class XLPivotShowValuesAs : uint8_t {
        Normal = 0,
        PercentOfGrandTotal,          ///< percentOfTotal
        PercentOfColumnTotal,         ///< percentOfCol
        PercentOfRowTotal,            ///< percentOfRow
        PercentOf,                    ///< percent (needs baseField + baseItem)
        PercentOfParentRowTotal,      ///< x14 percentOfParentRow
        PercentOfParentColumnTotal,   ///< x14 percentOfParentCol
        PercentOfParentTotal,         ///< x14 percentOfParent (needs baseField)
        DifferenceFrom,               ///< difference (baseField + baseItem)
        PercentDifferenceFrom,        ///< percentDiff (baseField + baseItem)
        RunningTotalIn,               ///< runTotal (baseField)
        PercentRunningTotalIn,        ///< x14 percentOfRunningTotal (baseField)
        RankSmallestToLargest,        ///< x14 rankAscending (baseField)
        RankLargestToSmallest,        ///< x14 rankDescending (baseField)
        Index                         ///< x14 index
    };

    struct XLPivotShowValuesAsOptions
    {
        XLPivotShowValuesAs type{XLPivotShowValuesAs::Normal};
        std::string         baseField;    ///< cache field name when required
        std::string         baseItem;     ///< shared-item label when required
    };

    struct XLPivotField
    {
        std::string              name;
        XLPivotSubtotal          subtotal = XLPivotSubtotal::Sum;
        std::string              customName;       ///< data field caption; also axis display name when set
        uint32_t                 numFmtId = 0;
        std::vector<std::string> selectedItems;

        // Axis / layout (row, column, filter) — excelize PivotTableField parity
        bool        compact{false};
        bool        outline{false};
        bool        showAll{false};
        bool        insertBlankRow{false};
        bool        defaultSubtotal{true};

        // Data field only
        XLPivotShowValuesAsOptions showValuesAs{};
    };

    class OPENXLSX_EXPORT XLPivotTableOptions
    {
    public:
        /**
         * @brief Construct fully initialized Pivot Table options.
         * @param name The unique name for the pivot table (e.g., "PivotTable1").
         * @param sourceRange Sheet range, Excel Table name, or defined name.
         * @param targetCell Anchor cell (e.g. "A3") or full location range (e.g. "A3:G20").
         */
        XLPivotTableOptions(std::string name, std::string sourceRange, std::string targetCell)
            : m_name(std::move(name)), m_sourceRange(std::move(sourceRange)), m_targetCell(std::move(targetCell))
        {}

        // --- Fluent Configuration Builders --- //

        XLPivotTableOptions& addRowField(std::string fieldName, std::vector<std::string> selectedItems = {})
        {
            XLPivotField f;
            f.name          = std::move(fieldName);
            f.selectedItems = std::move(selectedItems);
            m_rows.push_back(std::move(f));
            return *this;
        }

        XLPivotTableOptions& addRowField(XLPivotField field)
        {
            m_rows.push_back(std::move(field));
            return *this;
        }

        XLPivotTableOptions& addColumnField(std::string fieldName, std::vector<std::string> selectedItems = {})
        {
            XLPivotField f;
            f.name          = std::move(fieldName);
            f.selectedItems = std::move(selectedItems);
            m_columns.push_back(std::move(f));
            return *this;
        }

        XLPivotTableOptions& addColumnField(XLPivotField field)
        {
            m_columns.push_back(std::move(field));
            return *this;
        }

        XLPivotTableOptions& addDataField(std::string         fieldName,
                                          std::string         customName = "",
                                          XLPivotSubtotal     subtotal   = XLPivotSubtotal::Sum,
                                          uint32_t            numFmtId   = 0,
                                          XLPivotShowValuesAsOptions showValuesAs = {})
        {
            XLPivotField f;
            f.name         = std::move(fieldName);
            f.customName   = std::move(customName);
            f.subtotal     = subtotal;
            f.numFmtId     = numFmtId;
            f.showValuesAs = std::move(showValuesAs);
            m_data.push_back(std::move(f));
            return *this;
        }

        XLPivotTableOptions& addDataField(XLPivotField field)
        {
            m_data.push_back(std::move(field));
            return *this;
        }

        XLPivotTableOptions& addFilterField(std::string fieldName, std::vector<std::string> selectedItems = {})
        {
            XLPivotField f;
            f.name          = std::move(fieldName);
            f.selectedItems = std::move(selectedItems);
            m_filters.push_back(std::move(f));
            return *this;
        }

        XLPivotTableOptions& addFilterField(XLPivotField field)
        {
            m_filters.push_back(std::move(field));
            return *this;
        }

        XLPivotTableOptions& setPivotTableStyle(std::string styleName)
        {
            m_pivotTableStyleName = std::move(styleName);
            return *this;
        }
        XLPivotTableOptions& setDataOnRows(bool value)
        {
            m_dataOnRows = value;
            return *this;
        }
        XLPivotTableOptions& setRowGrandTotals(bool value)
        {
            m_rowGrandTotals = value;
            return *this;
        }
        XLPivotTableOptions& setColGrandTotals(bool value)
        {
            m_colGrandTotals = value;
            return *this;
        }
        XLPivotTableOptions& setShowDrill(bool value)
        {
            m_showDrill = value;
            return *this;
        }
        XLPivotTableOptions& setUseAutoFormatting(bool value)
        {
            m_useAutoFormatting = value;
            return *this;
        }
        XLPivotTableOptions& setPageOverThenDown(bool value)
        {
            m_pageOverThenDown = value;
            return *this;
        }
        XLPivotTableOptions& setMergeItem(bool value)
        {
            m_mergeItem = value;
            return *this;
        }
        XLPivotTableOptions& setCompactData(bool value)
        {
            m_compactData = value;
            return *this;
        }
        XLPivotTableOptions& setShowError(bool value)
        {
            m_showError = value;
            return *this;
        }
        XLPivotTableOptions& setShowRowHeaders(bool value)
        {
            m_showRowHeaders = value;
            return *this;
        }
        XLPivotTableOptions& setShowColHeaders(bool value)
        {
            m_showColHeaders = value;
            return *this;
        }
        XLPivotTableOptions& setShowRowStripes(bool value)
        {
            m_showRowStripes = value;
            return *this;
        }
        XLPivotTableOptions& setShowColStripes(bool value)
        {
            m_showColStripes = value;
            return *this;
        }
        XLPivotTableOptions& setShowLastColumn(bool value)
        {
            m_showLastColumn = value;
            return *this;
        }

        XLPivotTableOptions& setClassicLayout(bool value)
        {
            m_classicLayout = value;
            return *this;
        }
        XLPivotTableOptions& setFieldPrintTitles(bool value)
        {
            m_fieldPrintTitles = value;
            return *this;
        }
        XLPivotTableOptions& setItemPrintTitles(bool value)
        {
            m_itemPrintTitles = value;
            return *this;
        }

        /**
         * @brief Optional explicit location rectangle (e.g. "G4:M30"). When empty, the library
         *        estimates size from targetCell + field cardinality.
         */
        XLPivotTableOptions& setLocationRange(std::string range)
        {
            m_locationRange = std::move(range);
            return *this;
        }

        // --- Getters --- //
        [[nodiscard]] const std::string& name() const { return m_name; }
        [[nodiscard]] const std::string& sourceRange() const { return m_sourceRange; }
        [[nodiscard]] const std::string& targetCell() const { return m_targetCell; }
        [[nodiscard]] const std::string& locationRange() const { return m_locationRange; }
        [[nodiscard]] const std::vector<XLPivotField>& rows() const { return m_rows; }
        [[nodiscard]] const std::vector<XLPivotField>& columns() const { return m_columns; }
        [[nodiscard]] const std::vector<XLPivotField>& data() const { return m_data; }
        [[nodiscard]] const std::vector<XLPivotField>& filters() const { return m_filters; }
        [[nodiscard]] const std::string& pivotTableStyleName() const { return m_pivotTableStyleName; }

        [[nodiscard]] bool dataOnRows() const { return m_dataOnRows; }
        [[nodiscard]] bool rowGrandTotals() const { return m_rowGrandTotals; }
        [[nodiscard]] bool colGrandTotals() const { return m_colGrandTotals; }
        [[nodiscard]] bool showDrill() const { return m_showDrill; }
        [[nodiscard]] bool useAutoFormatting() const { return m_useAutoFormatting; }
        [[nodiscard]] bool pageOverThenDown() const { return m_pageOverThenDown; }
        [[nodiscard]] bool mergeItem() const { return m_mergeItem; }
        [[nodiscard]] bool compactData() const { return m_compactData; }
        [[nodiscard]] bool showError() const { return m_showError; }
        [[nodiscard]] bool showRowHeaders() const { return m_showRowHeaders; }
        [[nodiscard]] bool showColHeaders() const { return m_showColHeaders; }
        [[nodiscard]] bool showRowStripes() const { return m_showRowStripes; }
        [[nodiscard]] bool showColStripes() const { return m_showColStripes; }
        [[nodiscard]] bool showLastColumn() const { return m_showLastColumn; }

        [[nodiscard]] bool classicLayout() const { return m_classicLayout; }
        [[nodiscard]] bool fieldPrintTitles() const { return m_fieldPrintTitles; }
        [[nodiscard]] bool itemPrintTitles() const { return m_itemPrintTitles; }

        // Mutators used by options() extraction
        void setNameInternal(std::string v) { m_name = std::move(v); }
        void setSourceRangeInternal(std::string v) { m_sourceRange = std::move(v); }
        void setTargetCellInternal(std::string v) { m_targetCell = std::move(v); }
        void setRowsInternal(std::vector<XLPivotField> v) { m_rows = std::move(v); }
        void setColumnsInternal(std::vector<XLPivotField> v) { m_columns = std::move(v); }
        void setDataInternal(std::vector<XLPivotField> v) { m_data = std::move(v); }
        void setFiltersInternal(std::vector<XLPivotField> v) { m_filters = std::move(v); }
        void setPivotTableStyleNameInternal(std::string v) { m_pivotTableStyleName = std::move(v); }
        void setLocationRangeInternal(std::string v) { m_locationRange = std::move(v); }
        void setBoolFlags(bool dataOnRows,
                          bool rowGrandTotals,
                          bool colGrandTotals,
                          bool showDrill,
                          bool useAutoFormatting,
                          bool pageOverThenDown,
                          bool mergeItem,
                          bool compactData,
                          bool showError,
                          bool showRowHeaders,
                          bool showColHeaders,
                          bool showRowStripes,
                          bool showColStripes,
                          bool showLastColumn,
                          bool classicLayout,
                          bool fieldPrintTitles,
                          bool itemPrintTitles)
        {
            m_dataOnRows        = dataOnRows;
            m_rowGrandTotals    = rowGrandTotals;
            m_colGrandTotals    = colGrandTotals;
            m_showDrill         = showDrill;
            m_useAutoFormatting = useAutoFormatting;
            m_pageOverThenDown  = pageOverThenDown;
            m_mergeItem         = mergeItem;
            m_compactData       = compactData;
            m_showError         = showError;
            m_showRowHeaders    = showRowHeaders;
            m_showColHeaders    = showColHeaders;
            m_showRowStripes    = showRowStripes;
            m_showColStripes    = showColStripes;
            m_showLastColumn    = showLastColumn;
            m_classicLayout     = classicLayout;
            m_fieldPrintTitles  = fieldPrintTitles;
            m_itemPrintTitles   = itemPrintTitles;
        }

    private:
        std::string m_name;
        std::string m_sourceRange;
        std::string m_targetCell;
        std::string m_locationRange;    ///< optional explicit location ref

        std::vector<XLPivotField> m_rows;
        std::vector<XLPivotField> m_columns;
        std::vector<XLPivotField> m_data;
        std::vector<XLPivotField> m_filters;

        bool m_dataOnRows{false};
        bool m_rowGrandTotals{true};
        bool m_colGrandTotals{true};
        bool m_showDrill{true};
        bool m_useAutoFormatting{false};
        bool m_pageOverThenDown{false};
        bool m_mergeItem{false};
        bool m_compactData{true};
        bool m_showError{false};
        bool m_showRowHeaders{true};
        bool m_showColHeaders{true};
        bool m_showRowStripes{false};
        bool m_showColStripes{false};
        bool m_showLastColumn{false};

        bool m_classicLayout{false};
        bool m_fieldPrintTitles{false};
        bool m_itemPrintTitles{false};

        std::string m_pivotTableStyleName{"PivotStyleLight16"};
    };

    class OPENXLSX_EXPORT XLPivotTable final : public XLXmlFile
    {
    public:
        class XLRelationships relationships();

        XLPivotTable() : XLXmlFile(nullptr) {}
        explicit XLPivotTable(XLXmlData* xmlData);
        ~XLPivotTable()                                        = default;
        XLPivotTable(const XLPivotTable& other)                = default;
        XLPivotTable(XLPivotTable&& other) noexcept            = default;
        XLPivotTable& operator=(const XLPivotTable& other)     = default;
        XLPivotTable& operator=(XLPivotTable&& other) noexcept = default;

        void setName(std::string_view name);

        /**
         * @brief Tell Excel to refresh this pivot table when the file is opened.
         */
        void setRefreshOnLoad(bool refresh = true);

        /**
         * @brief Get the name of the pivot table.
         */
        std::string name() const;

        /**
         * @brief Get the target location anchor (first cell of location/@ref).
         */
        std::string targetCell() const;

        /**
         * @brief Get the data source range or table/defined name.
         */
        std::string sourceRange() const;

        /**
         * @brief Best-effort reconstruction of create-time options from package XML.
         * @details Round-trip is complete for tables written by OpenXLSX / excelize-style parts;
         *          Excel-authored files are parsed as much as the definition allows.
         */
        [[nodiscard]] XLPivotTableOptions options() const;

        /**
         * @brief Get the pivot cache definition object.
         */
        class XLPivotCacheDefinition cacheDefinition() const;

        /**
         * @brief Change the data source range for the pivot table.
         */
        void changeSourceRange(std::string_view newRange);
    };

    class OPENXLSX_EXPORT XLPivotCacheDefinition final : public XLXmlFile
    {
    public:
        class XLRelationships relationships();

        XLPivotCacheDefinition() : XLXmlFile(nullptr) {}
        explicit XLPivotCacheDefinition(XLXmlData* xmlData);
        ~XLPivotCacheDefinition()                                                  = default;
        XLPivotCacheDefinition(const XLPivotCacheDefinition& other)                = default;
        XLPivotCacheDefinition(XLPivotCacheDefinition&& other) noexcept            = default;
        XLPivotCacheDefinition& operator=(const XLPivotCacheDefinition& other)     = default;
        XLPivotCacheDefinition& operator=(XLPivotCacheDefinition&& other) noexcept = default;

        std::string sourceRange() const;
        void        changeSourceRange(std::string_view newRange);

        /**
         * @brief Cache field names in definition order.
         */
        [[nodiscard]] std::vector<std::string> fieldNames() const;
    };

    class OPENXLSX_EXPORT XLPivotCacheRecords final : public XLXmlFile
    {
    public:
        XLPivotCacheRecords() : XLXmlFile(nullptr) {}
        explicit XLPivotCacheRecords(XLXmlData* xmlData);
        ~XLPivotCacheRecords()                                               = default;
        XLPivotCacheRecords(const XLPivotCacheRecords& other)                = default;
        XLPivotCacheRecords(XLPivotCacheRecords&& other) noexcept            = default;
        XLPivotCacheRecords& operator=(const XLPivotCacheRecords& other)     = default;
        XLPivotCacheRecords& operator=(XLPivotCacheRecords&& other) noexcept = default;
    };

    // Helpers shared by create/extract (implemented in XLPivotTable.cpp)
    [[nodiscard]] const char* pivotShowValuesAsToString(XLPivotShowValuesAs t);
    [[nodiscard]] XLPivotShowValuesAs pivotShowValuesAsFromString(std::string_view s);
    [[nodiscard]] bool pivotShowValuesAsNeedsX14(XLPivotShowValuesAs t);
    [[nodiscard]] bool pivotShowValuesAsNeedsBaseField(XLPivotShowValuesAs t);
    [[nodiscard]] bool pivotShowValuesAsNeedsBaseItem(XLPivotShowValuesAs t);
    [[nodiscard]] const char* pivotSubtotalToString(XLPivotSubtotal s);
    [[nodiscard]] XLPivotSubtotal pivotSubtotalFromString(std::string_view s);

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLPIVOTTABLE_HPP

#ifndef OPENXLSX_XLPIVOTTABLE_HPP
#define OPENXLSX_XLPIVOTTABLE_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include <string>
#include <string_view>
#include <vector>

namespace OpenXLSX
{
    enum class XLPivotSubtotal { Sum, Average, Count, Max, Min, Product };

    struct XLPivotField
    {
        std::string     name;
        XLPivotSubtotal subtotal = XLPivotSubtotal::Sum;
        std::string     customName;
        uint32_t        numFmtId = 0; // Number format ID (e.g., 3 for '#,##0', 4 for '#,##0.00', 9 for '0%')
    };

    class OPENXLSX_EXPORT XLPivotTableOptions
    {
    public:
        /**
         * @brief Construct fully initialized Pivot Table options.
         * @param name The unique name for the pivot table (e.g., "PivotTable1").
         * @param sourceRange The data source range (e.g., "Sheet1!A1:D100").
         * @param targetCell The cell to anchor the pivot table (e.g., "A3").
         */
        XLPivotTableOptions(std::string name, std::string sourceRange, std::string targetCell)
            : m_name(std::move(name)), m_sourceRange(std::move(sourceRange)), m_targetCell(std::move(targetCell)) {}

        // --- Fluent Configuration Builders --- //

        XLPivotTableOptions& addRowField(std::string fieldName) {
            m_rows.push_back({std::move(fieldName), XLPivotSubtotal::Sum, "", 0});
            return *this;
        }

        XLPivotTableOptions& addColumnField(std::string fieldName) {
            m_columns.push_back({std::move(fieldName), XLPivotSubtotal::Sum, "", 0});
            return *this;
        }

        XLPivotTableOptions& addDataField(std::string fieldName, std::string customName = "", XLPivotSubtotal subtotal = XLPivotSubtotal::Sum, uint32_t numFmtId = 0) {
            m_data.push_back({std::move(fieldName), subtotal, std::move(customName), numFmtId});
            return *this;
        }

        XLPivotTableOptions& addFilterField(std::string fieldName) {
            m_filters.push_back({std::move(fieldName), XLPivotSubtotal::Sum, "", 0});
            return *this;
        }

        XLPivotTableOptions& setPivotTableStyle(std::string styleName) { m_pivotTableStyleName = std::move(styleName); return *this; }
        XLPivotTableOptions& setDataOnRows(bool value)        { m_dataOnRows = value; return *this; }
        XLPivotTableOptions& setRowGrandTotals(bool value)    { m_rowGrandTotals = value; return *this; }
        XLPivotTableOptions& setColGrandTotals(bool value)    { m_colGrandTotals = value; return *this; }
        XLPivotTableOptions& setShowDrill(bool value)         { m_showDrill = value; return *this; }
        XLPivotTableOptions& setUseAutoFormatting(bool value) { m_useAutoFormatting = value; return *this; }
        XLPivotTableOptions& setPageOverThenDown(bool value)  { m_pageOverThenDown = value; return *this; }
        XLPivotTableOptions& setMergeItem(bool value)         { m_mergeItem = value; return *this; }
        XLPivotTableOptions& setCompactData(bool value)       { m_compactData = value; return *this; }
        XLPivotTableOptions& setShowError(bool value)         { m_showError = value; return *this; }
        XLPivotTableOptions& setShowRowHeaders(bool value)    { m_showRowHeaders = value; return *this; }
        XLPivotTableOptions& setShowColHeaders(bool value)    { m_showColHeaders = value; return *this; }
        XLPivotTableOptions& setShowRowStripes(bool value)    { m_showRowStripes = value; return *this; }
        XLPivotTableOptions& setShowColStripes(bool value)    { m_showColStripes = value; return *this; }
        XLPivotTableOptions& setShowLastColumn(bool value)    { m_showLastColumn = value; return *this; }

        // --- Getters --- //
        [[nodiscard]] const std::string& name() const        { return m_name; }
        [[nodiscard]] const std::string& sourceRange() const { return m_sourceRange; }
        [[nodiscard]] const std::string& targetCell() const  { return m_targetCell; }
        [[nodiscard]] const std::vector<XLPivotField>& rows() const    { return m_rows; }
        [[nodiscard]] const std::vector<XLPivotField>& columns() const { return m_columns; }
        [[nodiscard]] const std::vector<XLPivotField>& data() const    { return m_data; }
        [[nodiscard]] const std::vector<XLPivotField>& filters() const { return m_filters; }
        [[nodiscard]] const std::string& pivotTableStyleName() const { return m_pivotTableStyleName; }

        [[nodiscard]] bool dataOnRows() const        { return m_dataOnRows; }
        [[nodiscard]] bool rowGrandTotals() const    { return m_rowGrandTotals; }
        [[nodiscard]] bool colGrandTotals() const    { return m_colGrandTotals; }
        [[nodiscard]] bool showDrill() const         { return m_showDrill; }
        [[nodiscard]] bool useAutoFormatting() const { return m_useAutoFormatting; }
        [[nodiscard]] bool pageOverThenDown() const  { return m_pageOverThenDown; }
        [[nodiscard]] bool mergeItem() const         { return m_mergeItem; }
        [[nodiscard]] bool compactData() const       { return m_compactData; }
        [[nodiscard]] bool showError() const         { return m_showError; }
        [[nodiscard]] bool showRowHeaders() const    { return m_showRowHeaders; }
        [[nodiscard]] bool showColHeaders() const    { return m_showColHeaders; }
        [[nodiscard]] bool showRowStripes() const    { return m_showRowStripes; }
        [[nodiscard]] bool showColStripes() const    { return m_showColStripes; }
        [[nodiscard]] bool showLastColumn() const    { return m_showLastColumn; }

    private:
        std::string m_name;
        std::string m_sourceRange;
        std::string m_targetCell;

        std::vector<XLPivotField> m_rows;
        std::vector<XLPivotField> m_columns;
        std::vector<XLPivotField> m_data;
        std::vector<XLPivotField> m_filters;

        bool m_dataOnRows        {false};
        bool m_rowGrandTotals    {true};
        bool m_colGrandTotals    {true};
        bool m_showDrill         {true};
        bool m_useAutoFormatting {false};
        bool m_pageOverThenDown  {false};
        bool m_mergeItem         {false};
        bool m_compactData       {true};
        bool m_showError         {false};
        bool m_showRowHeaders    {true};
        bool m_showColHeaders    {true};
        bool m_showRowStripes    {false};
        bool m_showColStripes    {false};
        bool m_showLastColumn    {false};
        
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
         * @brief Get the target location (e.g., "A3") of the pivot table in the worksheet.
         */
        std::string targetCell() const;

        /**
         * @brief Get the data source range (e.g., "Sheet1!A1:D100").
         */
        std::string sourceRange() const;

        /**
         * @brief Get the pivot cache definition object.
         */
        class XLPivotCacheDefinition cacheDefinition() const;

        /**
         * @brief Change the data source range for the pivot table.
         * @param newRange The new source range (e.g., "Sheet1!A1:E200").
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

        /**
         * @brief Get the data source range (e.g., "Sheet1!A1:D100").
         */
        std::string sourceRange() const;

        /**
         * @brief Change the data source range for the pivot cache.
         * @param newRange The new source range (e.g., "Sheet1!A1:E200").
         */
        void changeSourceRange(std::string_view newRange);
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
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLPIVOTTABLE_HPP

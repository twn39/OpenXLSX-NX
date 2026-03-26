#ifndef OPENXLSX_XLTABLES_HPP
#define OPENXLSX_XLTABLES_HPP

// ===== External Includes ===== //
#include <cstdint>    // uint8_t, uint16_t, uint32_t
#include <ostream>    // std::basic_ostream
#include <vector>
#include <string>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLException.hpp"
#include "XLXmlData.hpp"
#include "XLXmlFile.hpp"
#include "XLAutoFilter.hpp"
#include "XLTableColumn.hpp"

namespace OpenXLSX
{
    class XLWorksheet;
    class XLCellRange;
    
    struct XLSlicerOptions {
        std::string name = "";       // Defaults to column name if empty
        std::string caption = "";    // Defaults to column name if empty
        uint32_t width = 144;        // pixels
        uint32_t height = 180;       // pixels
        int32_t offsetX = 0;         // pixels
        int32_t offsetY = 0;         // pixels
    };

    class XLTable;

    /**
     * @brief The XLTableCollection class manages multiple tables within a worksheet.
     */
    class OPENXLSX_EXPORT XLTableCollection
    {
        friend class XLWorksheet;
    public:
        /**
         * @brief Constructor
         */
        XLTableCollection() = default;

        /**
         * @brief Constructor
         * @param node The worksheet XML node.
         * @param worksheet The parent worksheet.
         */
        explicit XLTableCollection(XMLNode node, XLWorksheet* worksheet);

        /**
         * @brief Get the number of tables in the collection.
         */
        size_t count() const;

        /**
         * @brief Get a table by its 0-based index.
         */
        XLTable operator[](size_t index) const;

        /**
         * @brief Get a table by its name.
         */
        XLTable table(std::string_view name) const;

        /**
         * @brief Add a new table to the worksheet.
         * @param name The table name.
         * @param range The cell range for the table (e.g. "A1:C5").
         * @return The newly created XLTable object.
         */
        XLTable add(std::string_view name, std::string_view range);

        /**
         * @brief Add a new table using a typed cell range.
         * @param name The table name.
         * @param range The XLCellRange object.
         * @return The newly created XLTable object.
         */
        XLTable add(std::string_view name, const XLCellRange& range);

        /**
         * @brief Check if the collection is valid.
         */
        bool valid() const;

    private:
        XMLNode m_sheetNode;
        XLWorksheet* m_worksheet{nullptr};
        mutable std::vector<XLTable> m_tables;
        mutable bool m_loaded{false};

        void load() const;
    };

    /**
     * @brief The XLTable class represents a single Excel table (.xml file).
     */
    class OPENXLSX_EXPORT XLTable : public XLXmlFile
    {
        friend class XLTableCollection;
        friend class XLWorksheet;
    public:
        /**
         * @brief Constructor
         */
        XLTable() : XLXmlFile(nullptr) {};

        /**
         * @brief The constructor.
         * @param xmlData the source XML of the table file
         */
        XLTable(XLXmlData* xmlData);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         */
        XLTable(const XLTable& other) = default;

        /**
         * @brief Move constructor.
         */
        XLTable(XLTable&& other) noexcept = default;

        /**
         * @brief The destructor
         */
        ~XLTable() = default;

        /**
         * @brief Assignment operator
         */
        XLTable& operator=(const XLTable&) = default;

        /**
         * @brief Move assignment operator.
         */
        XLTable& operator=(XLTable&& other) noexcept = default;

        /**
         * @brief Get the name of the table
         * @return std::string The table name
         */
        uint32_t id() const;

        /**
         * @brief Get the name of the table
         * @return std::string The table name
         */
        std::string name() const;

        /**
         * @brief Set the name of the table
         * @param name The new table name (must not contain spaces)
         */
        void setName(std::string_view name);

        /**
         * @brief Get the display name of the table
         * @return std::string The display name
         */
        std::string displayName() const;

        /**
         * @brief Set the display name of the table
         * @param name The new display name
         */
        void setDisplayName(std::string_view name);

        /**
         * @brief Get the range reference of the table (e.g., "A1:C10")
         * @return std::string The range reference
         */
        std::string rangeReference() const;

        /**
         * @brief Set the range reference of the table
         * @param ref The new range reference (e.g., "A1:C10")
         */
        void setRangeReference(std::string_view ref);

        /**
         * @brief Get the table style name
         * @return std::string The style name
         */
        std::string styleName() const;

        /**
         * @brief Set the table style name
         * @param styleName The new style name (e.g., "TableStyleMedium2")
         */
        void setStyleName(std::string_view styleName);

        /**
         * @brief Get the table comment.
         * @return The comment string.
         */
        std::string comment() const;

        /**
         * @brief Set the table comment.
         * @param comment The new comment string.
         */
        void setComment(std::string_view comment);

        /**
         * @brief Get the differential formatting ID for the header row.
         * @return The DxfId, or 0 if not set.
         */
        uint32_t headerRowDxfId() const;

        /**
         * @brief Set the differential formatting ID for the header row.
         * @param id The DxfId.
         */
        void setHeaderRowDxfId(uint32_t id);

        /**
         * @brief Get the differential formatting ID for the data rows.
         * @return The DxfId, or 0 if not set.
         */
        uint32_t dataDxfId() const;

        /**
         * @brief Set the differential formatting ID for the data rows.
         * @param id The DxfId.
         */
        void setDataDxfId(uint32_t id);

        /**
         * @brief Check if the table is published.
         * @return True if published.
         */
        bool published() const;

        /**
         * @brief Set whether the table is published.
         * @param published True if published.
         */
        void setPublished(bool published);

        /**
         * @brief Check if row stripes are shown
         * @return bool True if row stripes are shown
         */
        bool showRowStripes() const;

        /**
         * @brief Set whether row stripes are shown
         * @param show True to show row stripes
         */
        void setShowRowStripes(bool show);

        /**
         * @brief Check if the header row is shown
         * @return bool True if the header row is shown
         */
        bool showHeaderRow() const;

        /**
         * @brief Set whether the header row is shown
         * @param show True to show the header row
         */
        void setShowHeaderRow(bool show);

        /**
         * @brief Check if the totals row is shown
         * @return bool True if the totals row is shown
         */
        bool showTotalsRow() const;

        /**
         * @brief Set whether the totals row is shown
         * @param show True to show the totals row
         */
        void setShowTotalsRow(bool show);

        /**
         * @brief Check if column stripes are shown
         * @return bool True if column stripes are shown
         */
        bool showColumnStripes() const;

        /**
         * @brief Set whether column stripes are shown
         * @param show True to show column stripes
         */
        void setShowColumnStripes(bool show);

        /**
         * @brief Check if the first column is highlighted
         * @return bool True if the first column is highlighted
         */
        bool showFirstColumn() const;

        /**
         * @brief Set whether the first column is highlighted
         * @param show True to highlight the first column
         */
        void setShowFirstColumn(bool show);

        /**
         * @brief Check if the last column is highlighted
         * @return bool True if the last column is highlighted
         */
        bool showLastColumn() const;

        /**
         * @brief Set whether the last column is highlighted
         * @param show True to highlight the last column
         */
        void setShowLastColumn(bool show);

        /**
         * @brief Append a new column to the table
         * @param name The column name
         * @return The appended XLTableColumn
         */
        XLTableColumn appendColumn(std::string_view name);

        /**
         * @brief Get a column by its name
         * @param name The name of the column
         * @return The XLTableColumn object
         */
        XLTableColumn column(std::string_view name) const;

        /**
         * @brief Get a column by its 1-based ID
         * @param id The 1-based ID of the column
         * @return The XLTableColumn object
         */
        XLTableColumn column(uint32_t id) const;

        /**
         * @brief Get the auto filter object for the table.
         * @return XLAutoFilter The auto filter object.
         */
        XLAutoFilter autoFilter() const;

        /**
         * @brief Automatically resizes the table to fit the contiguous data in the worksheet.
         * @param worksheet The worksheet containing this table.
         */
        void resizeToFitData(const XLWorksheet& worksheet);

        /**
         * @brief Automatically initializes the columns of the table from the header row of the specified worksheet range.
         * @param worksheet The parent worksheet.
         */
        void createColumnsFromRange(const XLWorksheet& worksheet);

        /**
         * @brief Print the XML contents of this XLTable instance using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;
    };

    // For compatibility
    using XLTables = XLTableCollection;
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLTABLES_HPP

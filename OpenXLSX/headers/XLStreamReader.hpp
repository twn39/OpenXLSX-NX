//
// XLStreamReader — SAX streaming reader for large worksheets (phase C enhancements).
//

#ifndef OPENXLSX_XLSTREAMREADER_HPP
#define OPENXLSX_XLSTREAMREADER_HPP

#include "OpenXLSX-Exports.hpp"
#include "IZipArchive.hpp"
#include "XLCellValue.hpp"
#include "XLStyles_Common.hpp"    // XLStyleIndex
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

namespace OpenXLSX
{

    class XLWorksheet;

    /**
     * @brief Policy for logical row gaps (missing &lt;row&gt; elements in sheetData).
     */
    enum class XLStreamEmptyRowPolicy : uint8_t {
        /**
         * @brief Only return rows that exist in the XML (legacy default).
         * @details Matches prior OpenXLSX behaviour: if sheet jumps from row 1 to row 3,
         *          the second nextRow() yields row 3.
         */
        SkipMissingRows = 0,
        /**
         * @brief Synthesize empty rows for gaps in row numbers (excelize GetRows-like).
         * @details After row 1, nextRow() yields an empty row 2 before row 3.
         */
        EmitEmptyRows = 1
    };

    /**
     * @brief Options for XLStreamReader construction.
     */
    struct OPENXLSX_EXPORT XLStreamReadOptions
    {
        XLStreamEmptyRowPolicy emptyRows{XLStreamEmptyRowPolicy::SkipMissingRows};
        /**
         * @brief When true, formattedValue() / nextRowStrings() apply number formats via XLNumberFormatter.
         * @details Style lookup uses the workbook styles table (may load styles once).
         */
        bool applyNumberFormats{false};
    };

    /**
     * @brief Rich view of a single streamed cell (value + formula + style).
     */
    struct OPENXLSX_EXPORT XLStreamCellView
    {
        XLCellValue                 value{};
        std::optional<std::string>  formula{};
        std::optional<XLStyleIndex> styleIndex{};
        uint16_t                    column{0};    // 1-based; 0 if unknown
    };

    /**
     * @brief Row-level attributes for the last consumed physical/logical row (excelize GetRowOpts).
     */
    struct OPENXLSX_EXPORT XLStreamRowOptsView
    {
        std::optional<double>       height{};
        std::optional<bool>         hidden{};
        std::optional<uint8_t>      outlineLevel{};
        std::optional<XLStyleIndex> styleIndex{};
        bool                        isSyntheticEmpty{false};    // true when emitted by EmitEmptyRows
    };

    class OPENXLSX_EXPORT XLStreamReader
    {
    public:
        XLStreamReader()                                 = default;
        XLStreamReader(const XLStreamReader&)            = delete;
        XLStreamReader& operator=(const XLStreamReader&) = delete;

        XLStreamReader(XLStreamReader&& other) noexcept;
        XLStreamReader& operator=(XLStreamReader&& other) noexcept;

        ~XLStreamReader();

        /**
         * @brief Checks if there are more rows to read (physical and/or synthetic empty).
         */
        bool hasNext();

        /**
         * @brief Parses and returns the next row of cell values (backward-compatible API).
         * @details Empty cells between non-empty columns are filled as XLValueType::Empty.
         *          Does not allocate a DOM tree per row.
         */
        std::vector<XLCellValue> nextRow();

        /**
         * @brief Next row with formula, style index, and column metadata per cell.
         */
        std::vector<XLStreamCellView> nextRowDetailed();

        /**
         * @brief Next row as display strings (optionally applying number formats).
         * @details Empty cells become empty strings. When options.applyNumberFormats is false,
         *          uses raw string conversion of XLCellValue (similar to excelize RawCellValue path).
         */
        std::vector<std::string> nextRowStrings();

        /**
         * @brief Format a single cell view using workbook styles when applyNumberFormats is enabled.
         */
        std::string formattedValue(const XLStreamCellView& cell) const;

        /**
         * @brief 1-based index of the row last returned by nextRow / nextRowDetailed / nextRowStrings.
         */
        uint32_t currentRow() const;

        /**
         * @brief Row options for the last returned row.
         */
        XLStreamRowOptsView currentRowOpts() const;

        /**
         * @brief Last error message from a soft parse failure (empty if none).
         */
        const std::string& lastError() const { return m_lastError; }

        /**
         * @brief Closes the reader and releases resources.
         */
        void close();

        const XLStreamReadOptions& options() const { return m_options; }

    private:
        friend class XLWorksheet;

        explicit XLStreamReader(const XLWorksheet* worksheet, XLStreamReadOptions options = {});

        void fetchMoreData();
        void cleanup();
        void compactBufferIfNeeded();

        /** @return true if a complete physical &lt;row&gt; was loaded into m_pending*. */
        bool loadNextPhysicalRow();

        /** Consume either a synthetic empty row or the pending physical row into m_out*. */
        bool consumeNextLogicalRow();

        void parsePhysicalRowSlice(const char* p, const char* end);
        XLCellValue parseCellValue(const std::string& cellType, std::string& cellValue) const;
        std::string resolveFormatCode(XLStyleIndex styleIndex) const;
        std::string valueToRawString(const XLCellValue& value) const;

        static std::string builtinNumFmtCode(uint32_t numFmtId);

        const XLWorksheet*  m_worksheet{nullptr};
        IZipArchive         m_archive{};    // copy of archive handle for independent stream close
        void*               m_zipStream{nullptr};
        XLStreamReadOptions m_options{};

        std::string m_buffer;
        size_t      m_bufPos{0};
        bool        m_eof{false};

        uint32_t m_currentRow{0};       // last returned logical row
        uint32_t m_lastPhysicalRow{0};  // last physical row number seen in XML
        bool     m_hasPendingPhysical{false};
        bool     m_physicalExhausted{false};

        // Pending physical row (after loadNextPhysicalRow)
        uint32_t                      m_pendingRowNum{0};
        XLStreamRowOptsView           m_pendingRowOpts{};
        std::vector<XLStreamCellView> m_pendingCells{};

        // Last returned logical row metadata
        XLStreamRowOptsView m_currentRowOpts{};

        // Reusable scratch buffers
        std::string m_tagNameBuf;
        std::string m_attrNameBuf;
        std::string m_attrValueBuf;
        std::string m_textContentBuf;
        std::string m_lastError;

        // Reusable outputs to cut per-row allocations when caller reuses patterns
        std::vector<XLCellValue>      m_reuseValues;
        std::vector<XLStreamCellView> m_reuseDetailed;
        std::vector<std::string>      m_reuseStrings;
    };

}    // namespace OpenXLSX
#endif    // OPENXLSX_XLSTREAMREADER_HPP

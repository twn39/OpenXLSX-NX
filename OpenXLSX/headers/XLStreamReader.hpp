//
// XLStreamReader — SAX streaming reader for large worksheets (phase C enhancements).
//

#ifndef OPENXLSX_XLSTREAMREADER_HPP
#define OPENXLSX_XLSTREAMREADER_HPP

#include "OpenXLSX-Exports.hpp"
#include "IZipArchive.hpp"
#include "XLCellValue.hpp"
#include "XLRichText.hpp"
#include "XLStyles_Common.hpp"    // XLStyleIndex
#include <cstdint>
#include <filesystem>
#include <functional>
#include <optional>
#include <string>
#include <unordered_set>
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
     * @brief How shared-string cells (t="s") are resolved during streaming reads.
     */
    enum class XLStreamSharedStringMode : uint8_t {
        /**
         * @brief Use the workbook's in-memory shared string table (default, fastest for small SST).
         */
        Eager = 0,
        /**
         * @brief Build a byte-offset index into sharedStrings.xml and resolve strings on demand.
         * @details Reduces peak memory when the SST is huge; first access may load the SST part once for indexing.
         */
        Indexed = 1
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
        XLStreamSharedStringMode sharedStringMode{XLStreamSharedStringMode::Eager};
        /**
         * @brief Reject a single physical &lt;row&gt; larger than this many bytes (0 = unlimited).
         * @details Guards against pathological wide rows that defeat streaming memory bounds.
         *          Default 64 MiB.
         */
        size_t maxRowBytes{64 * 1024 * 1024};
        /**
         * @brief When Indexed SST exceeds this size, keep payload on a temp file and mmap-style seek.
         * @details 0 = always keep SST fully in memory. Default 16 MiB (excelize UnzipXMLSizeLimit default).
         */
        size_t sstSpillThreshold{16 * 1024 * 1024};
        /** Empty = system temp directory (used when SST spills). */
        std::filesystem::path tempDir{};
        /**
         * @brief 1-based column projection mask. Empty = all columns.
         * @details Used by nextRowProjected / forEachCellInNextRow; dense nextRow() still returns full width.
         */
        std::vector<uint16_t> columnFilter{};
        /**
         * @brief Reject sharedStrings (or other bulk loads via this reader) larger than this (0 = unlimited).
         * @details Default 256 MiB — zip-bomb style guard for Indexed SST getEntry.
         */
        size_t maxSstEntryBytes{256 * 1024 * 1024};
        /**
         * @brief When true, parse inlineStr &lt;r&gt; runs into XLRichText on XLStreamCellView.
         * @details Plain &lt;t&gt; text still fills value as String; richText is set when runs exist.
         */
        bool parseRichText{false};
        /**
         * @brief Max cells buffered when materializing streamColumns() (0 = unlimited). Default 50M cells.
         */
        size_t maxColumnarCells{50ull * 1000 * 1000};
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

        // Formula metadata (OOXML &lt;f&gt; attributes)
        std::optional<std::string> formulaType{};    // array | shared | dataTable
        std::optional<std::string> formulaRef{};
        std::optional<int>         formulaSi{};
        bool                       formulaCa{false};

        // Rich text (when parseRichText and inlineStr has &lt;r&gt; runs)
        std::optional<XLRichText> richText{};
    };

    /**
     * @brief Column-oriented cursor over a worksheet (excelize Cols-like).
     * @details Performs one full row-stream pass into a columnar buffer, then yields columns.
     *          Memory is O(non-empty cells); use maxColumnarCells to cap.
     */
    class OPENXLSX_EXPORT XLStreamColumns
    {
    public:
        XLStreamColumns() = default;
        explicit XLStreamColumns(std::vector<std::vector<XLCellValue>> columns);

        bool next();
        /** 1-based current column index (0 before first next / after end). */
        uint16_t currentColumn() const { return m_index; }
        const std::vector<XLCellValue>& values() const;
        size_t columnCount() const { return m_columns.size(); }

    private:
        std::vector<std::vector<XLCellValue>> m_columns;
        uint16_t                              m_index{0};    // 1-based after successful next()
        bool                                  m_started{false};
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
         * @brief Next row containing only non-empty cells (no sparse padding).
         */
        std::vector<XLStreamCellView> nextRowSparse();

        /**
         * @brief Fill @p out with the next dense row values; returns false at end.
         * @details Reuses @p out capacity to reduce allocations in tight ETL loops.
         */
        bool nextRowInto(std::vector<XLCellValue>& out);

        /**
         * @brief Fill @p out with the next detailed row; returns false at end.
         */
        bool nextRowDetailedInto(std::vector<XLStreamCellView>& out);

        /**
         * @brief Next row with only columns listed in options.columnFilter (or non-empty cells if filter empty).
         * @details Sparse: no empty padding between columns. Ideal for wide-table ETL projection.
         */
        std::vector<XLStreamCellView> nextRowProjected();

        /**
         * @brief Consume the next logical row and invoke @p fn for each projected cell.
         * @return false at end of sheet.
         */
        bool forEachCellInNextRow(const std::function<void(const XLStreamCellView&)>& fn);

        /**
         * @brief Materialize all columns via one full stream pass (excelize Cols equivalent).
         * @details Invalidates further nextRow* use on this reader (consumes the stream).
         */
        XLStreamColumns streamColumns();

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
        void        ensureIndexedSst() const;
        std::string lookupIndexedSst(int32_t index) const;
        bool        columnAllowed(uint16_t col) const;
        void        rebuildColumnFilterSet() const;

        static std::string builtinNumFmtCode(uint32_t numFmtId);

        const XLWorksheet*  m_worksheet{nullptr};
        IZipArchive         m_archive{};    // copy of archive handle for independent stream close
        void*               m_zipStream{nullptr};
        XLStreamReadOptions m_options{};
        mutable std::unordered_set<uint16_t> m_columnFilterSet{};
        mutable bool                         m_columnFilterReady{false};

        // Indexed SST (lazy): in-memory string or file-backed for large tables
        mutable bool                  m_sstIndexReady{false};
        mutable std::string           m_sstXml;             // small SST or empty when file-backed
        mutable std::filesystem::path m_sstTempPath;        // large SST spilled to disk
        mutable bool                  m_sstFileBacked{false};
        mutable std::vector<uint32_t> m_sstOffsets;    // start offset of each <si> element

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

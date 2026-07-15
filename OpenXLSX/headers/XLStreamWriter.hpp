//
// XLStreamWriter — streaming worksheet write API (excelize-compatible phases).
// Phase 0: cols / panes / page breaks (before first row)
// Phase 1: sheetData rows (strictly increasing row numbers)
// Phase 2: mergeCells / tableParts (injected at close)
//

#ifndef OPENXLSX_XLSTREAMWRITER_HPP
#define OPENXLSX_XLSTREAMWRITER_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"
#include "XLDateTime.hpp"
#include "XLSharedStrings.hpp"
#include "XLStyles_Common.hpp"    // XLStyleIndex
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace OpenXLSX
{

    class XLWorksheet;

    /**
     * @brief How streamWriter materializes the worksheet header.
     * @details Auto uses Fresh when the sheet has no rows (cheap skeleton), else Hybrid
     *          (serialize current DOM and append inside sheetData).
     */
    enum class XLStreamOpenMode : uint8_t { Auto = 0, Fresh = 1, Hybrid = 2 };

    /**
     * @brief How &lt;dimension&gt; is finalized on close without reloading the full sheet XML.
     * @details FixedSlot writes a constant-width placeholder (max Excel range) and patches it
     *          in-place via seek/overwrite — O(1) memory. Omit leaves the max-range placeholder.
     */
    enum class XLStreamDimensionMode : uint8_t {
        FixedSlot = 0,    // default: accurate used-range via fixed-width in-place patch
        OmitPatch = 1     // keep A1:XFD1048576 placeholder (no seek patch)
    };

    /**
     * @brief Options for XLStreamWriter construction.
     */
    struct OPENXLSX_EXPORT XLStreamWriteOptions
    {
        XLStreamOpenMode openMode{XLStreamOpenMode::Auto};
        /**
         * @brief Spill in-memory buffer to a temp file after this many bytes.
         * @details 0 = always use a temp file. Default 4 MiB (excelize uses 16 MiB).
         */
        size_t memorySpillThreshold{4 * 1024 * 1024};
        /** Empty = system temp directory. */
        std::filesystem::path tempDir{};
        bool   useSharedStrings{false};
        size_t maxUniqueStrings{100000};
        /**
         * @brief Hard budget for unique shared-string payload bytes while streaming.
         * @details When exceeded, further strings fall back to inlineStr (avoids SST OOM).
         *          0 = unlimited (not recommended for large exports). Default 64 MiB.
         */
        size_t maxSstBytes{64 * 1024 * 1024};
        XLStreamDimensionMode dimensionMode{XLStreamDimensionMode::FixedSlot};
        /**
         * @brief When true, addTable registration failures on close throw instead of being swallowed.
         */
        bool strictTableRegistration{true};
        /**
         * @brief Hybrid mode: if the serialized existing sheet XML exceeds this size, keep the
         *        pre-sheetData header on a temp file and stream it at first row (avoids a second
         *        full string peak). 0 = always keep header in memory. Default 4 MiB.
         */
        size_t hybridHeaderSpillThreshold{4 * 1024 * 1024};
        /**
         * @brief Default number format for XLStreamCell::withDateTime when no style is given.
         * @details Excel-compatible; created once via styles().findOrCreateStyle and cached.
         */
        std::string defaultDateTimeFormat{"yyyy-mm-dd hh:mm:ss"};
    };

    /**
     * @brief Structured freeze/split panes for stream writers (excelize SetPanes equivalent).
     * @details Call setPanes() before the first row. Prefer this over setPanesXml().
     */
    struct OPENXLSX_EXPORT XLStreamPanesOptions
    {
        /** Freeze panes: number of columns frozen on the left (0 = none). excelize FreezePanes x. */
        uint16_t freezeCols{0};
        /** Freeze panes: number of rows frozen on the top (0 = none). */
        uint32_t freezeRows{0};
        /**
         * @brief Optional top-left cell of the scrollable region (1-based address, e.g. "B2").
         * @details When empty and freeze is set, derived as (freezeCols+1, freezeRows+1).
         */
        std::string topLeftCell{};
        /** When true, use split state instead of frozen (xSplit/ySplit as twips-like units from Excel). */
        bool split{false};
        double xSplit{0.0};
        double ySplit{0.0};
        /** activePane attribute: bottomRight | topRight | bottomLeft | topLeft */
        std::string activePane{"bottomRight"};
    };

    /**
     * @brief Optional row-level attributes for streaming writes (excelize RowOpts equivalent).
     */
    struct OPENXLSX_EXPORT XLStreamRowOpts
    {
        std::optional<double>       height;          // points; writes ht + customHeight="1"
        std::optional<bool>         hidden;
        std::optional<uint8_t>      outlineLevel;    // 0–7
        std::optional<XLStyleIndex> styleIndex;      // row s= + customFormat="1"
    };

    /**
     * @brief Minimal table definition for streaming AddTable (excelize-compatible constraints).
     * @details Only one table per StreamWriter. Call after rows are written, before close().
     */
    struct OPENXLSX_EXPORT XLStreamTableOptions
    {
        std::string name{"Table1"};
        std::string styleName{"TableStyleMedium2"};
        std::string range;    // e.g. "A1:D10"
        bool        showRowStripes{true};
        bool        showColumnStripes{false};
        bool        showFirstColumn{false};
        bool        showLastColumn{false};
        /** Optional header labels; if empty, columns are named Column1..N from the range width. */
        std::vector<std::string> headers;
    };

    /**
     * @brief Stream hyperlink (injected as &lt;hyperlinks&gt; on close; external uses sheet relationships).
     */
    struct OPENXLSX_EXPORT XLStreamHyperlink
    {
        std::string cellRef;     // e.g. "A1"
        std::string target;      // URL or internal location (Sheet1!A1)
        bool        external{true};
        std::string tooltip;
        std::string display;     // optional display attribute for internal links
    };

    /**
     * @brief Formula type attributes for shared / array formulas (OOXML &lt;f t=…&gt;).
     */
    enum class XLStreamFormulaType : uint8_t {
        Normal = 0,
        Array,
        Shared,
        DataTable
    };

    /**
     * @brief A structure that represents a cell to be appended via the streaming API.
     */
    class OPENXLSX_EXPORT XLStreamCell
    {
    public:
        explicit XLStreamCell(XLCellValue val) : value(std::move(val)), styleIndex(std::nullopt), formula(std::nullopt) {}

        XLStreamCell(XLCellValue val, XLStyleIndex style) : value(std::move(val)), styleIndex(style), formula(std::nullopt) {}

        XLStreamCell(XLCellValue val, XLStyleIndex style, std::string formulaText)
            : value(std::move(val)), styleIndex(style), formula(std::move(formulaText))
        {}

        static XLStreamCell withFormula(std::string formulaText, XLCellValue cached = XLCellValue {})
        {
            XLStreamCell cell(std::move(cached));
            cell.formula = std::move(formulaText);
            return cell;
        }

        /**
         * @brief Array formula (t="array", optional ref= for multi-cell CSE).
         */
        static XLStreamCell withArrayFormula(std::string formulaText, std::string ref = {}, XLCellValue cached = XLCellValue {})
        {
            XLStreamCell cell = withFormula(std::move(formulaText), std::move(cached));
            cell.formulaType  = XLStreamFormulaType::Array;
            if (!ref.empty()) cell.formulaRef = std::move(ref);
            return cell;
        }

        /**
         * @brief Shared formula master (t="shared" si=… ref=…) or dependent (empty formula body, si= only).
         */
        static XLStreamCell withSharedFormula(int si, std::string formulaText = {}, std::string ref = {}, XLCellValue cached = XLCellValue {})
        {
            XLStreamCell cell = withFormula(std::move(formulaText), std::move(cached));
            cell.formulaType  = XLStreamFormulaType::Shared;
            cell.formulaSi    = si;
            if (!ref.empty()) cell.formulaRef = std::move(ref);
            return cell;
        }

        /**
         * @brief Date/time cell: stores Excel serial and requests an automatic date number format.
         * @param dt Excel date/time value.
         * @param style Optional explicit style; when omitted the writer caches a default date style.
         */
        static XLStreamCell withDateTime(const XLDateTime& dt, std::optional<XLStyleIndex> style = std::nullopt)
        {
            XLStreamCell cell(XLCellValue(dt.serial()));
            cell.styleIndex  = style;
            cell.treatAsDate = true;
            return cell;
        }

        XLCellValue                 value;
        std::optional<XLStyleIndex> styleIndex;
        std::optional<std::string>  formula;
        /** When true and styleIndex is empty, writer applies the cached default date/time style. */
        bool treatAsDate{false};

        XLStreamFormulaType         formulaType{XLStreamFormulaType::Normal};
        std::optional<std::string>  formulaRef;     // array/shared ref attribute
        std::optional<int>          formulaSi;      // shared formula index
        bool                        formulaCa{false}; // calculate always
    };

    class OPENXLSX_EXPORT XLStreamWriter
    {
    public:
        XLStreamWriter()                                 = default;
        XLStreamWriter(const XLStreamWriter&)            = delete;
        XLStreamWriter& operator=(const XLStreamWriter&) = delete;

        XLStreamWriter(XLStreamWriter&& other) noexcept;
        XLStreamWriter& operator=(XLStreamWriter&& other) noexcept;

        ~XLStreamWriter();

        bool isStreamActive() const;

        // ── Phase 0 (must be called before any row) ──────────────────────────

        /**
         * @brief Set column width for [minCol, maxCol] (1-based, inclusive). excelize SetColWidth.
         */
        void setColWidth(uint16_t minCol, uint16_t maxCol, double width);

        /**
         * @brief Show/hide columns. excelize SetColVisible.
         */
        void setColVisible(uint16_t minCol, uint16_t maxCol, bool visible);

        /**
         * @brief Column outline level 1–7. excelize SetColOutlineLevel.
         */
        void setColOutlineLevel(uint16_t col, uint8_t level);

        /**
         * @brief Apply a cell style index to an entire column range. excelize SetColStyle.
         */
        void setColStyle(uint16_t minCol, uint16_t maxCol, XLStyleIndex style);

        /**
         * @brief Freeze/split panes XML fragment (sheetViews). Must be before first row.
         * @param panesXml Inner content or full &lt;sheetViews&gt;…&lt;/sheetViews&gt; element.
         * @deprecated Prefer setPanes(XLStreamPanesOptions) for structured freeze/split.
         */
        void setPanesXml(std::string panesXml);

        /**
         * @brief Structured freeze or split panes (must be before first row). excelize SetPanes.
         */
        void setPanes(const XLStreamPanesOptions& panes);

        /**
         * @brief Freeze panes so that @p topLeftCell is the first unfrozen cell (e.g. "B2").
         */
        void freezePanes(const std::string& topLeftCell);

        /**
         * @brief Freeze the first @p cols columns and @p rows rows.
         */
        void freezePanes(uint16_t cols, uint32_t rows);

        /**
         * @brief Record a horizontal/vertical page break at a cell (stored until close).
         */
        void insertPageBreak(std::string_view cellRef);

        // ── Phase 1: rows ────────────────────────────────────────────────────

        void appendRow(const std::vector<XLCellValue>& values);
        void appendRow(const std::vector<XLCellValue>& values, const XLStreamRowOpts& opts);
        void appendRow(std::initializer_list<XLCellValue> values);
        void appendRow(std::initializer_list<XLCellValue> values, const XLStreamRowOpts& opts);

        void appendRow(const std::vector<XLStreamCell>& cells);
        void appendRow(const std::vector<XLStreamCell>& cells, const XLStreamRowOpts& opts);
        void appendRow(std::initializer_list<XLStreamCell> cells);
        void appendRow(std::initializer_list<XLStreamCell> cells, const XLStreamRowOpts& opts);

        void setRow(uint32_t row, uint16_t startCol, const std::vector<XLStreamCell>& cells, const XLStreamRowOpts& opts = {});
        void setRow(uint32_t row, uint16_t startCol, const std::vector<XLCellValue>& values, const XLStreamRowOpts& opts = {});
        void setRow(uint32_t row, uint16_t startCol, std::initializer_list<XLStreamCell> cells, const XLStreamRowOpts& opts = {});
        void setRow(uint32_t row, uint16_t startCol, std::initializer_list<XLCellValue> values, const XLStreamRowOpts& opts = {});

        void setRow(const std::string& cellRef, const std::vector<XLStreamCell>& cells, const XLStreamRowOpts& opts = {});
        void setRow(const std::string& cellRef, const std::vector<XLCellValue>& values, const XLStreamRowOpts& opts = {});
        void setRow(const std::string& cellRef, std::initializer_list<XLStreamCell> cells, const XLStreamRowOpts& opts = {});
        void setRow(const std::string& cellRef, std::initializer_list<XLCellValue> values, const XLStreamRowOpts& opts = {});

        // ── Phase 2 (after rows, before close) ───────────────────────────────

        /**
         * @brief Merge a cell range (excelize MergeCell). Injected at close after sheetData.
         */
        void mergeCell(std::string_view topLeft, std::string_view bottomRight);

        /**
         * @brief Register one Excel table for this stream (excelize AddTable constraints).
         * @details Call after rows, before close(). Creates table part + relationship on close.
         */
        void addTable(const XLStreamTableOptions& table);

        /**
         * @brief Sheet autoFilter range (e.g. "A1:D100"). Injected after sheetData on close.
         */
        void setAutoFilter(std::string_view range);

        /**
         * @brief External hyperlink (URL). Relationship + &lt;hyperlinks&gt; injected on close.
         */
        void addHyperlink(std::string_view cellRef, std::string_view url, std::string_view tooltip = {});

        /**
         * @brief Internal hyperlink (location like Sheet2!A1 or #'Sheet 2'!A1).
         */
        void addInternalHyperlink(std::string_view cellRef, std::string_view location, std::string_view tooltip = {});

        /**
         * @brief Flushes buffers, writes footer, patches dimension, deactivates stream.
         */
        void close();

        void flush() { close(); }

        std::string getTempFilePath() const;

        uint32_t lastRow() const { return m_lastWrittenRow; }
        uint16_t maxColumn() const { return m_maxColumn; }

        /** Bytes committed to the sink (memory and/or temp file), excluding the active row buffer. */
        uint64_t bytesWritten() const { return m_sinkBytesWritten; }

        /** True when dimension was installed as a fixed-width slot for O(1) patch. */
        bool hasDimensionSlot() const { return m_dimensionSlotSize > 0; }

        const XLStreamWriteOptions& options() const { return m_options; }

    private:
        friend class XLWorksheet;

        explicit XLStreamWriter(XLWorksheet* worksheet, XLStreamWriteOptions options);

        // Legacy constructor path
        explicit XLStreamWriter(XLWorksheet* worksheet, bool useSharedStrings, size_t maxUniqueStrings);

        template<typename T>
        void setRowImpl(uint32_t row, uint16_t startCol, const std::vector<T>& items, const XLStreamRowOpts* opts);

        void writeRowOpen(uint32_t row, const XLStreamRowOpts* opts);
        void writeCellValueBody(const XLCellValue& value);
        void flushWriteBuffer();
        void flushSheetDataClose();
        void patchDimension();
        void installDimensionSlot();
        void markWriterOpen(bool open) noexcept;
        void checkStreamWritable() const;
        void checkPhase0() const;
        void ensureSheetDataStarted();
        void sinkWrite(std::string_view data);
        void sinkSpillToFile();
        void materializeToWorksheet();
        std::string buildColsXml() const;
        std::string buildMergeCellsXml() const;
        std::string buildPageBreaksXml() const;
        std::string buildAutoFilterXml() const;
        void        writeFormulaOpenTag(const XLStreamCell* cellMeta, const std::string& formulaText);
        void        normalizeColRange(uint16_t& minCol, uint16_t& maxCol) const;
        void        injectBeforeSheetData(std::string fragment);
        void        streamPreHeaderToSink();
        std::string buildPanesXml(const XLStreamPanesOptions& panes) const;
        XLStyleIndex ensureDefaultDateStyle();

        struct ColDirective
        {
            uint16_t                    minCol{1};
            uint16_t                    maxCol{1};
            std::optional<double>       width;
            std::optional<bool>         hidden;
            std::optional<uint8_t>      outlineLevel;
            std::optional<XLStyleIndex> styleIndex;
        };

        static constexpr size_t kRowFlushThreshold = 256 * 1024;
        static constexpr double kMaxRowHeight      = 409.0;
        static constexpr double kMaxColWidth       = 255.0;
        /** Fixed-width dimension placeholder: max Excel used range (A1:XFD1048576). */
        static constexpr std::string_view kDimensionSlot = R"(<dimension ref="A1:XFD1048576"/>)";

        XLStreamWriteOptions m_options{};

        // Sink: memory-first, optional spill
        std::string           m_memBuffer;
        std::filesystem::path m_tempPath;
        std::ofstream         m_fileStream;
        bool                  m_usingFile{false};
        uint64_t              m_sinkBytesWritten{0};

        // Fixed-slot dimension patch (O(1) close; no full-file reload)
        uint64_t m_dimensionByteOffset{0};
        size_t   m_dimensionSlotSize{0};

        // Header assembly (not flushed until first row / close)
        std::string           m_preContent;     // in-memory header (Hybrid may include existing rows)
        std::filesystem::path m_preHeaderPath;  // large Hybrid header kept on disk
        bool                  m_preHeaderFileBacked{false};
        std::string           m_bottomHalf;     // from </sheetData> inclusive
        bool                  m_headerWritten{false};
        bool                  m_sheetDataStarted{false};    // Phase 1 entered
        bool                  m_preHasOpenSheetData{false}; // preContent already contains <sheetData>

        std::vector<ColDirective> m_colDirectives;
        std::string               m_panesXml;
        std::vector<std::string>  m_mergeRefs;    // "A1:B2"
        std::vector<std::string>  m_rowBreaks;    // row numbers as strings
        std::vector<std::string>  m_colBreaks;
        std::optional<XLStreamTableOptions> m_table;
        std::optional<std::string>          m_autoFilterRange;
        std::vector<XLStreamHyperlink>      m_hyperlinks;

        uint32_t m_currentRow{1};
        uint32_t m_lastWrittenRow{0};
        uint16_t m_maxColumn{0};
        uint16_t m_baseMaxCol{0};
        uint32_t m_baseLastRow{0};
        bool     m_active{false};

        std::string m_writeBuffer;

        XLWorksheet* m_worksheet{nullptr};
        bool         m_wroteFormula{false};
        size_t       m_sstBytesUsed{0};
        XLStyleIndex m_cachedDateStyle{XLInvalidStyleIndex};
        bool         m_dateStyleReady{false};

        mutable FlatHashMap<std::string_view, int32_t> m_localCache{};
        static constexpr size_t                        kLocalCacheLimit = 5000;
    };

}    // namespace OpenXLSX
#endif    // OPENXLSX_XLSTREAMWRITER_HPP

//
// Created for OpenXLSX streaming write API (phase A/B: lifecycle, dimension, formula, row opts, setRow).
//

#ifndef OPENXLSX_XLSTREAMWRITER_HPP
#define OPENXLSX_XLSTREAMWRITER_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"
#include "XLStyles_Common.hpp"    // XLStyleIndex
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <optional>
#include <string>
#include <vector>

#include "XLSharedStrings.hpp"

namespace OpenXLSX
{

    class XLWorksheet;

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
     * @brief A structure that represents a cell to be appended via the streaming API.
     * @details By explicitly constructing an XLStreamCell, users can attach specific styles
     * and/or formulas to a streamed row without buffering the entire DOM in memory.
     */
    class OPENXLSX_EXPORT XLStreamCell
    {
    public:
        /**
         * @brief Explicitly constructs a styled streaming cell from an XLCellValue.
         * @param val The cell value (e.g. integer, string, double).
         */
        explicit XLStreamCell(XLCellValue val) : value(std::move(val)), styleIndex(std::nullopt), formula(std::nullopt) {}

        /**
         * @brief Constructs a styled streaming cell with an explicitly assigned style index.
         * @param val The cell value.
         * @param style The XLStyleIndex (e.g. retrieved from doc.styles().findOrCreateStyle(...))
         */
        XLStreamCell(XLCellValue val, XLStyleIndex style) : value(std::move(val)), styleIndex(style), formula(std::nullopt) {}

        /**
         * @brief Constructs a streaming cell with style and formula.
         * @param val Cached value (may be Empty if only the formula is written).
         * @param style Cell style index.
         * @param formulaText Excel formula body without a leading '=' (as stored in OOXML &lt;f&gt;).
         */
        XLStreamCell(XLCellValue val, XLStyleIndex style, std::string formulaText)
            : value(std::move(val)), styleIndex(style), formula(std::move(formulaText))
        {}

        /**
         * @brief Factory for formula cells with an optional cached value.
         */
        static XLStreamCell withFormula(std::string formulaText, XLCellValue cached = XLCellValue {})
        {
            XLStreamCell cell(std::move(cached));
            cell.formula = std::move(formulaText);
            return cell;
        }

        XLCellValue                 value;
        std::optional<XLStyleIndex> styleIndex;
        std::optional<std::string>  formula;
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

        /**
         * @brief Appends a row of unstyled values at the next sequential row.
         */
        void appendRow(const std::vector<XLCellValue>& values);
        void appendRow(const std::vector<XLCellValue>& values, const XLStreamRowOpts& opts);

        /**
         * @brief Appends a row of styled/formula cells at the next sequential row.
         */
        void appendRow(const std::vector<XLStreamCell>& cells);
        void appendRow(const std::vector<XLStreamCell>& cells, const XLStreamRowOpts& opts);

        /**
         * @brief Writes a row at an explicit 1-based row index and starting column.
         * @details Row numbers must be strictly increasing across the stream session
         *          (same constraint as excelize StreamWriter.SetRow).
         */
        void setRow(uint32_t row, uint16_t startCol, const std::vector<XLStreamCell>& cells, const XLStreamRowOpts& opts = {});
        void setRow(uint32_t row, uint16_t startCol, const std::vector<XLCellValue>& values, const XLStreamRowOpts& opts = {});

        /**
         * @brief Convenience overload: start cell reference such as "C10".
         */
        void setRow(const std::string& cellRef, const std::vector<XLStreamCell>& cells, const XLStreamRowOpts& opts = {});
        void setRow(const std::string& cellRef, const std::vector<XLCellValue>& values, const XLStreamRowOpts& opts = {});

        /**
         * @brief Flushes buffers, writes the sheet footer, patches &lt;dimension&gt;, and deactivates the stream.
         * @details Idempotent. Prefer calling this (or flush()) before XLDocument::save().
         *          save() will throw if a stream writer is still open.
         */
        void close();

        /**
         * @brief Alias for close() (excelize-style naming).
         */
        void flush() { close(); }

        std::string getTempFilePath() const;

        /** @brief 1-based last successfully written row number, or 0 if none. */
        uint32_t lastRow() const { return m_lastWrittenRow; }

        /** @brief Maximum 1-based column index written in this stream session (0 if none). */
        uint16_t maxColumn() const { return m_maxColumn; }

    private:
        friend class XLWorksheet;

        explicit XLStreamWriter(XLWorksheet* worksheet, bool useSharedStrings = false, size_t maxUniqueStrings = 100000);

        template<typename T>
        void setRowImpl(uint32_t row, uint16_t startCol, const std::vector<T>& items, const XLStreamRowOpts* opts);

        void writeRowOpen(uint32_t row, const XLStreamRowOpts* opts);
        void writeCellValueBody(const XLCellValue& value);
        void flushWriteBuffer();
        void flushSheetDataClose();
        void patchDimensionInTempFile();
        void markWriterOpen(bool open) noexcept;
        void checkStreamWritable() const;

        // Flush the in-memory write buffer to disk when it exceeds this threshold.
        // 256 KB gives a good balance between memory use and syscall frequency.
        static constexpr size_t kFlushThreshold = 256 * 1024;
        static constexpr double kMaxRowHeight   = 409.0;

        std::filesystem::path m_tempPath;
        std::ofstream         m_stream;
        uint32_t              m_currentRow{1};       // next row for appendRow
        uint32_t              m_lastWrittenRow{0};   // last setRow/appendRow row
        uint16_t              m_maxColumn{0};
        uint16_t              m_baseMaxCol{0};       // columnCount() snapshot at stream open
        uint32_t              m_baseLastRow{0};      // DOM rowCount() at stream open
        bool                  m_active{false};
        std::string           m_bottomHalf;

        // Write buffer — avoids one syscall per cell by coalescing multiple
        // small writes into a single fstream::write() call.
        std::string m_writeBuffer;

        XLWorksheet* m_worksheet{nullptr};
        bool         m_useSharedStrings{false};
        size_t       m_maxUniqueStrings{100000};
        bool         m_wroteFormula{false};

        // Bounded zero-copy local cache
        mutable FlatHashMap<std::string_view, int32_t> m_localCache{};
        static constexpr size_t                       kLocalCacheLimit = 5000;
    };

}    // namespace OpenXLSX
#endif    // OPENXLSX_XLSTREAMWRITER_HPP

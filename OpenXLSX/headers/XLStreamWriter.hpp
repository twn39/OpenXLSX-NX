#ifndef OPENXLSX_XLSTREAMWRITER_HPP
#define OPENXLSX_XLSTREAMWRITER_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"
#include "XLStyles_Common.hpp"    // XLStyleIndex
#include <cstddef>
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
     * @brief A structure that represents a cell to be appended via the streaming API.
     * @details By explicitly constructing an XLStreamCell, users can attach specific styles
     * to a streamed row without buffering the entire DOM in memory.
     */
    class OPENXLSX_EXPORT XLStreamCell
    {
    public:
        /**
         * @brief Explicitly constructs a styled streaming cell from an XLCellValue.
         * @param val The cell value (e.g. integer, string, double).
         */
        explicit XLStreamCell(XLCellValue val) : value(std::move(val)), styleIndex(std::nullopt) {}

        /**
         * @brief Constructs a styled streaming cell with an explicitly assigned style index.
         * @param val The cell value.
         * @param style The XLStyleIndex (e.g. retrieved from doc.workbook().styles().create(...))
         */
        XLStreamCell(XLCellValue val, XLStyleIndex style) : value(std::move(val)), styleIndex(style) {}

        XLCellValue                 value;
        std::optional<XLStyleIndex> styleIndex;
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
         * @brief Appends a row of unstyled values to the stream.
         * @param values A vector of XLCellValue items.
         */
        void appendRow(const std::vector<XLCellValue>& values);

        /**
         * @brief Appends a row of styled values to the stream.
         * @param cells A vector of XLStreamCell items.
         */
        void appendRow(const std::vector<XLStreamCell>& cells);

        std::string getTempFilePath() const;
        void        close();

    private:
        friend class XLWorksheet;

        explicit XLStreamWriter(XLWorksheet* worksheet, bool useSharedStrings = false, size_t maxUniqueStrings = 100000);

        template<typename T>
        void appendRowImpl(const std::vector<T>& items);

        void flushWriteBuffer();
        void flushSheetDataClose();

        // Flush the in-memory write buffer to disk when it exceeds this threshold.
        // 256 KB gives a good balance between memory use and syscall frequency.
        static constexpr size_t kFlushThreshold = 256 * 1024;

        std::filesystem::path m_tempPath;
        std::ofstream         m_stream;
        uint32_t              m_currentRow{1};
        bool                  m_active{false};
        std::string           m_bottomHalf;

        // Write buffer — avoids one syscall per cell by coalescing multiple
        // small writes into a single fstream::write() call.
        std::string m_writeBuffer;

        XLWorksheet* m_worksheet{nullptr};
        bool         m_useSharedStrings{false};
        size_t       m_maxUniqueStrings{100000};

        // Bounded zero-copy local cache
        mutable FlatHashMap<std::string_view, int32_t> m_localCache{};
        static constexpr size_t                       kLocalCacheLimit = 5000;
    };

}    // namespace OpenXLSX
#endif    // OPENXLSX_XLSTREAMWRITER_HPP

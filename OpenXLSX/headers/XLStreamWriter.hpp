#ifndef OPENXLSX_XLSTREAMWRITER_HPP
#define OPENXLSX_XLSTREAMWRITER_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"
#include <cstddef>
#include <filesystem>
#include <fstream>
#include <string>
#include <vector>

namespace OpenXLSX {

    class XLWorksheet;

    class OPENXLSX_EXPORT XLStreamWriter {
    public:
        XLStreamWriter() = default;
        XLStreamWriter(const XLStreamWriter&) = delete;
        XLStreamWriter& operator=(const XLStreamWriter&) = delete;

        XLStreamWriter(XLStreamWriter&& other) noexcept;
        XLStreamWriter& operator=(XLStreamWriter&& other) noexcept;

        ~XLStreamWriter();

        bool        isStreamActive() const;
        void        appendRow(const std::vector<XLCellValue>& values);
        std::string getTempFilePath() const;
        void        close();

    private:
        friend class XLWorksheet;

        explicit XLStreamWriter(XLWorksheet* worksheet);

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
    };

}
#endif // OPENXLSX_XLSTREAMWRITER_HPP

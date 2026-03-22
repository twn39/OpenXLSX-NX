#ifndef OPENXLSX_XLSTREAMWRITER_HPP
#define OPENXLSX_XLSTREAMWRITER_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"
#include <string>
#include <vector>
#include <fstream>
#include <filesystem>
#include <memory>

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

        bool isStreamActive() const;
        void appendRow(const std::vector<XLCellValue>& values);
        std::string getTempFilePath() const;
        void close();

    private:
        friend class XLWorksheet;
        
        explicit XLStreamWriter(XLWorksheet* worksheet);
        
        void flushSheetDataClose();

        XLWorksheet* m_worksheet{nullptr};
        std::filesystem::path m_tempPath;
        std::unique_ptr<std::ofstream> m_stream;
        uint32_t m_currentRow{1};
        bool m_active{false};
        std::string m_bottomHalf;
    };

}
#endif // OPENXLSX_XLSTREAMWRITER_HPP

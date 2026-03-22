#ifndef OPENXLSX_XLSTREAMREADER_HPP
#define OPENXLSX_XLSTREAMREADER_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"
#include <string>
#include <vector>

namespace OpenXLSX {

    class XLWorksheet;

    class OPENXLSX_EXPORT XLStreamReader {
    public:
        XLStreamReader() = default;
        XLStreamReader(const XLStreamReader&) = delete;
        XLStreamReader& operator=(const XLStreamReader&) = delete;

        XLStreamReader(XLStreamReader&& other) noexcept;
        XLStreamReader& operator=(XLStreamReader&& other) noexcept;

        ~XLStreamReader();

        /**
         * @brief Checks if there are more rows to read.
         * @return true if there are more rows, false otherwise.
         */
        bool hasNext();

        /**
         * @brief Parses and returns the next row of data.
         * @return A vector of XLCellValue representing the row. Empty cells are filled as XLValueType::Empty.
         */
        std::vector<XLCellValue> nextRow();

    private:
        friend class XLWorksheet;
        
        explicit XLStreamReader(const XLWorksheet* worksheet);
        
        void fetchMoreData();
        std::string extractNextRowXml();
        void cleanup();

        const XLWorksheet* m_worksheet{nullptr};
        void* m_zipStream{nullptr};
        
        std::string m_buffer;
        bool m_eof{false};
        uint32_t m_currentRow{0};
    };

}
#endif // OPENXLSX_XLSTREAMREADER_HPP
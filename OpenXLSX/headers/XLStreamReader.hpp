#ifndef OPENXLSX_XLSTREAMREADER_HPP
#define OPENXLSX_XLSTREAMREADER_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"
#include <cstdint>
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
         * @brief Parses and returns the next row of data using a SAX-style state machine.
         * @details Does not allocate a DOM tree per row; instead, it scans the raw XML
         *          bytes directly. This is the key optimization for large file streaming.
         * @return A vector of XLCellValue representing the row. Empty cells are filled as XLValueType::Empty.
         */
        std::vector<XLCellValue> nextRow();

        /**
         * @brief Returns the 1-based index of the row last read by nextRow().
         * @return The current row index.
         */
        uint32_t currentRow() const;

    private:
        friend class XLWorksheet;
        
        explicit XLStreamReader(const XLWorksheet* worksheet);
        
        void fetchMoreData();

        /**
         * @brief SAX state machine states for parsing worksheet XML.
         */
        enum class SaxState : uint8_t {
            Scanning,       // Scanning for an opening '<'
            InTagName,      // Reading the tag name
            InAttrName,     // Reading an attribute name
            AfterAttrName,  // After attr name, waiting for '='
            InAttrValue,    // Reading an attribute value (inside quotes)
            InTextContent,  // Inside a tag's text content (e.g., <v>42</v>)
            InClosingTag,   // Reading a closing tag (after '</')
        };

        void cleanup();

        const XLWorksheet* m_worksheet{nullptr};
        void*              m_zipStream{nullptr};
        
        std::string m_buffer;
        bool        m_eof{false};
        uint32_t    m_currentRow{0};

        // Reusable scratch buffers to avoid per-row heap allocation
        std::string m_tagNameBuf;
        std::string m_attrNameBuf;
        std::string m_attrValueBuf;
        std::string m_textContentBuf;
    };

}
#endif // OPENXLSX_XLSTREAMREADER_HPP
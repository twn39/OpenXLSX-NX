// ===== External Includes ===== //
#include <stdexcept>

// ===== OpenXLSX Includes ===== //
#include "XLStreamWriter.hpp"
#include "XLWorksheet.hpp"
#include "XLCellReference.hpp"
#include "XLException.hpp"

namespace OpenXLSX {

    XLStreamWriter::XLStreamWriter(XLWorksheet* worksheet) 
        : m_worksheet(worksheet), 
          m_tempPath(std::filesystem::temp_directory_path() / ("openxlsx_stream_" + std::to_string(std::rand()) + ".xml")),
          m_stream(std::make_unique<std::ofstream>(m_tempPath, std::ios::binary)),
          m_currentRow(1),
          m_active(true)
    {
        if (!m_stream->is_open()) {
            throw XLInternalError("Failed to open temporary stream file");
        }
    }

    XLStreamWriter::~XLStreamWriter() {
        if (m_stream && m_stream->is_open()) {
            m_stream->close();
        }
        if (std::filesystem::exists(m_tempPath)) {
            std::filesystem::remove(m_tempPath);
        }
    }

    XLStreamWriter::XLStreamWriter(XLStreamWriter&& other) noexcept
        : m_worksheet(other.m_worksheet),
          m_tempPath(std::move(other.m_tempPath)),
          m_stream(std::move(other.m_stream)),
          m_currentRow(other.m_currentRow),
          m_active(other.m_active)
    {
        other.m_worksheet = nullptr;
        other.m_active = false;
    }

    XLStreamWriter& XLStreamWriter::operator=(XLStreamWriter&& other) noexcept {
        if (this != &other) {
            m_worksheet = other.m_worksheet;
            m_tempPath = std::move(other.m_tempPath);
            m_stream = std::move(other.m_stream);
            m_currentRow = other.m_currentRow;
            m_active = other.m_active;
            other.m_worksheet = nullptr;
            other.m_active = false;
        }
        return *this;
    }

    bool XLStreamWriter::isStreamActive() const { return m_active && m_stream && m_stream->is_open(); }

    std::string XLStreamWriter::getTempFilePath() const { return m_tempPath.string(); }

    void XLStreamWriter::appendRow(const std::vector<XLCellValue>& values) {
        if (!isStreamActive()) throw XLInternalError("Stream writer is not active");

        *m_stream << "<row r=\"" << m_currentRow << "\">";
        
        uint16_t colIdx = 1;
        for (const auto& val : values) {
            if (val.type() != XLValueType::Empty) {
                std::string cellRef = XLCellReference::columnAsString(colIdx) + std::to_string(m_currentRow);
                
                *m_stream << "<c r=\"" << cellRef << "\"";
                
                // For optimal streaming, bypass the shared string table completely
                // and write strings as "inlineStr". This keeps memory flat.
                switch (val.type()) {
                    case XLValueType::String:
                        *m_stream << " t=\"inlineStr\"><is><t xml:space=\"preserve\">" << val.get<std::string>() << "</t></is></c>";
                        break;
                    case XLValueType::Boolean:
                        *m_stream << " t=\"b\"><v>" << (val.get<bool>() ? "1" : "0") << "</v></c>";
                        break;
                    case XLValueType::Float:
                    case XLValueType::Integer:
                        *m_stream << " t=\"n\"><v>" << XLCellValue(val).getString() << "</v></c>";
                        break;
                    default:
                        *m_stream << "><v>" << XLCellValue(val).getString() << "</v></c>";
                        break;
                }
            }
            colIdx++;
        }
        
        *m_stream << "</row>";
        m_currentRow++;
    }

    void XLStreamWriter::flushSheetDataClose() {
        if (m_stream && m_stream->is_open()) {
            *m_stream << "</sheetData></worksheet>";
            m_stream->close();
        }
        m_active = false;
    }
}

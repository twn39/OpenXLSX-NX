// ===== External Includes ===== //
#include <stdexcept>
#include <string_view>
#include <random>
#include <chrono>

// ===== OpenXLSX Includes ===== //
#include "XLStreamWriter.hpp"
#include "XLWorksheet.hpp"
#include "XLCellReference.hpp"
#include "XLException.hpp"

namespace {
    std::string escapeXml(std::string_view str) {
        std::string escaped;
        escaped.reserve(str.size() + 10); // Reserve a little extra to avoid small reallocations
        for (char c : str) {
            switch (c) {
                case '<': escaped += "&lt;"; break;
                case '>': escaped += "&gt;"; break;
                case '&': escaped += "&amp;"; break;
                case '"': escaped += "&quot;"; break;
                case '\'': escaped += "&apos;"; break;
                default: escaped += c; break;
            }
        }
        return escaped;
    }
}

namespace OpenXLSX {

    XLStreamWriter::XLStreamWriter(XLWorksheet* /*worksheet*/) 
        : m_tempPath(std::filesystem::temp_directory_path() / fmt::format("openxlsx_stream_{}_{}.xml", std::chrono::system_clock::now().time_since_epoch().count(), std::rand())),
          m_stream(m_tempPath, std::ios::binary),
          m_currentRow(1),
          m_active(true)
    {
        if (!m_stream.is_open()) {
            throw XLInternalError("Failed to open temporary stream file: " + m_tempPath.string());
        }
    }

    XLStreamWriter::~XLStreamWriter() {
        if (m_active) {
            flushSheetDataClose();
        }
    }

    XLStreamWriter::XLStreamWriter(XLStreamWriter&& other) noexcept
        : m_tempPath(std::move(other.m_tempPath)),
          m_stream(std::move(other.m_stream)),
          m_currentRow(other.m_currentRow),
          m_active(other.m_active),
          m_bottomHalf(std::move(other.m_bottomHalf))
    {
        other.m_active = false;
    }

    XLStreamWriter& XLStreamWriter::operator=(XLStreamWriter&& other) noexcept {
        if (this != &other) {
            if (m_active) {
                flushSheetDataClose();
            }
            m_tempPath = std::move(other.m_tempPath);
            m_stream = std::move(other.m_stream);
            m_currentRow = other.m_currentRow;
            m_active = other.m_active;
            m_bottomHalf = std::move(other.m_bottomHalf);
            other.m_active = false;
        }
        return *this;
    }

    bool XLStreamWriter::isStreamActive() const { return m_active && m_stream.is_open(); }

    std::string XLStreamWriter::getTempFilePath() const { return m_tempPath.string(); }

    void XLStreamWriter::appendRow(const std::vector<XLCellValue>& values) {
        if (!isStreamActive()) throw XLInternalError("Stream writer is not active");

        m_stream << "<row r=\"" << m_currentRow << "\">";
        
        uint16_t colIdx = 1;
        for (const auto& val : values) {
            if (val.type() != XLValueType::Empty) {
                std::string cellRef = fmt::format("{}{}", XLCellReference::columnAsString(colIdx), m_currentRow);
                
                m_stream << "<c r=\"" << cellRef << "\"";
                
                // For optimal streaming, bypass the shared string table completely
                // and write strings as "inlineStr". This keeps memory flat.
                switch (val.type()) {
                    case XLValueType::String:
                        m_stream << " t=\"inlineStr\"><is><t xml:space=\"preserve\">" << escapeXml(val.get<std::string>()) << "</t></is></c>";
                        break;
                    case XLValueType::Boolean:
                        m_stream << " t=\"b\"><v>" << (val.get<bool>() ? "1" : "0") << "</v></c>";
                        break;
                    case XLValueType::Float:
                    case XLValueType::Integer:
                        m_stream << " t=\"n\"><v>" << XLCellValue(val).getString() << "</v></c>";
                        break;
                    default:
                        m_stream << "><v>" << escapeXml(XLCellValue(val).getString()) << "</v></c>";
                        break;
                }
            }
            colIdx++;
        }
        
        m_stream << "</row>";
        m_currentRow++;
    }

    void XLStreamWriter::close() {
        if (m_active) {
            flushSheetDataClose();
        }
    }

    void XLStreamWriter::flushSheetDataClose() {
        if (m_stream.is_open()) {
            m_stream << m_bottomHalf;
            m_stream.close();
        }
        m_active = false;
    }
}

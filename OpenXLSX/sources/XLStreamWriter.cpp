// ===== External Includes ===== //
#include <chrono>
#include <random>

// ===== OpenXLSX Includes ===== //
#include "XLException.hpp"
#include "XLStreamWriter.hpp"
#include "XLWorksheet.hpp"

namespace {

    // Fast XML escape: appends escaped characters to `out` to avoid intermediate allocations.
    void appendEscaped(std::string& out, std::string_view sv)
    {
        for (char c : sv) {
            switch (c) {
                case '<':  out += "&lt;";   break;
                case '>':  out += "&gt;";   break;
                case '&':  out += "&amp;";  break;
                case '"':  out += "&quot;"; break;
                case '\'': out += "&apos;"; break;
                default:   out += c;        break;
            }
        }
    }

    // Convert a 1-based column index to an Excel column-letter string.
    // Returns the result as a small std::string (at most 3 chars for valid Excel columns).
    std::string columnToString(uint16_t col)
    {
        std::string result;
        result.reserve(3);
        while (col > 0) {
            result += static_cast<char>('A' + static_cast<char>((col - 1U) % 26U));
            col = static_cast<uint16_t>((col - 1U) / 26U);
        }
        // Letters were built in reverse order
        std::reverse(result.begin(), result.end());
        return result;
    }

}    // anonymous namespace

namespace OpenXLSX {

    XLStreamWriter::XLStreamWriter(XLWorksheet* /*worksheet*/)
        : m_tempPath(std::filesystem::temp_directory_path() /
                     (std::string("openxlsx_stream_") +
                      std::to_string(std::chrono::system_clock::now().time_since_epoch().count()) +
                      "_" +
                      []() -> std::string {
                          std::mt19937 rng(std::random_device{}());
                          return std::to_string(rng());
                      }() +
                      ".xml")),
          m_stream(m_tempPath, std::ios::binary),
          m_active(true)
    {
        if (!m_stream.is_open())
            throw XLInternalError("Failed to open temporary stream file: " + m_tempPath.string());

        // Reserve write buffer up-front — avoids reallocations during normal use
        m_writeBuffer.reserve(kFlushThreshold + 64UL * 1024UL);
    }

    XLStreamWriter::~XLStreamWriter()
    {
        if (m_active) flushSheetDataClose();
    }

    XLStreamWriter::XLStreamWriter(XLStreamWriter&& other) noexcept
        : m_tempPath(std::move(other.m_tempPath)),
          m_stream(std::move(other.m_stream)),
          m_currentRow(other.m_currentRow),
          m_active(other.m_active),
          m_bottomHalf(std::move(other.m_bottomHalf)),
          m_writeBuffer(std::move(other.m_writeBuffer))
    {
        other.m_active = false;
    }

    XLStreamWriter& XLStreamWriter::operator=(XLStreamWriter&& other) noexcept
    {
        if (this != &other) {
            if (m_active) flushSheetDataClose();
            m_tempPath    = std::move(other.m_tempPath);
            m_stream      = std::move(other.m_stream);
            m_currentRow  = other.m_currentRow;
            m_active      = other.m_active;
            m_bottomHalf  = std::move(other.m_bottomHalf);
            m_writeBuffer = std::move(other.m_writeBuffer);
            other.m_active = false;
        }
        return *this;
    }

    bool        XLStreamWriter::isStreamActive() const { return m_active && m_stream.is_open(); }
    std::string XLStreamWriter::getTempFilePath() const { return m_tempPath.string(); }

    // ─────────────────────────────────────────────────────────────────────────
    //  appendRow — batches all cell XML into m_writeBuffer, flushing to disk
    //  only when the buffer reaches kFlushThreshold.  This reduces the number
    //  of fstream write() syscalls dramatically vs. one write per cell.
    // ─────────────────────────────────────────────────────────────────────────
    void XLStreamWriter::appendRow(const std::vector<XLCellValue>& values)
    {
        if (!isStreamActive()) throw XLInternalError("Stream writer is not active");

        m_writeBuffer += "<row r=\"";
        m_writeBuffer += std::to_string(m_currentRow);
        m_writeBuffer += "\">";

        uint16_t colIdx = 1;
        for (const auto& val : values) {
            if (val.type() != XLValueType::Empty) {
                const std::string colStr = columnToString(colIdx);

                m_writeBuffer += "<c r=\"";
                m_writeBuffer += colStr;
                m_writeBuffer += std::to_string(m_currentRow);
                m_writeBuffer += '"';

                switch (val.type()) {
                    case XLValueType::String:
                        m_writeBuffer += R"( t="inlineStr"><is><t xml:space="preserve">)";
                        appendEscaped(m_writeBuffer, val.get<std::string>());
                        m_writeBuffer += "</t></is></c>";
                        break;

                    case XLValueType::Boolean:
                        m_writeBuffer += R"( t="b"><v>)";
                        m_writeBuffer += (val.get<bool>() ? '1' : '0');
                        m_writeBuffer += "</v></c>";
                        break;

                    case XLValueType::Integer:
                        m_writeBuffer += R"( t="n"><v>)";
                        m_writeBuffer += std::to_string(val.get<int64_t>());
                        m_writeBuffer += "</v></c>";
                        break;

                    case XLValueType::Float:
                        m_writeBuffer += R"( t="n"><v>)";
                        m_writeBuffer += XLCellValue(val).getString();
                        m_writeBuffer += "</v></c>";
                        break;

                    default:
                        m_writeBuffer += "><v>";
                        appendEscaped(m_writeBuffer, XLCellValue(val).getString());
                        m_writeBuffer += "</v></c>";
                        break;
                }
            }
            ++colIdx;
        }

        m_writeBuffer += "</row>";
        ++m_currentRow;

        if (m_writeBuffer.size() >= kFlushThreshold) flushWriteBuffer();
    }

    void XLStreamWriter::close()
    {
        if (m_active) flushSheetDataClose();
    }

    void XLStreamWriter::flushWriteBuffer()
    {
        if (!m_writeBuffer.empty() && m_stream.is_open()) {
            m_stream.write(m_writeBuffer.data(), static_cast<std::streamsize>(m_writeBuffer.size()));
            m_writeBuffer.clear();
        }
    }

    void XLStreamWriter::flushSheetDataClose()
    {
        if (m_stream.is_open()) {
            flushWriteBuffer();
            m_stream << m_bottomHalf;
            m_stream.close();
        }
        m_active = false;
    }

}    // namespace OpenXLSX

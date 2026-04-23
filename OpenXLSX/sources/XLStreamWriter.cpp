// ===== External Includes ===== //
#include <charconv>
#include <chrono>
#include <fmt/format.h>
#include <random>

// ===== OpenXLSX Includes ===== //
#include "XLException.hpp"
#include "XLStreamWriter.hpp"
#include "XLWorksheet.hpp"

namespace OpenXLSX
{

    XLStreamWriter::XLStreamWriter(XLWorksheet* worksheet)
        : m_tempPath(std::filesystem::temp_directory_path() /
                     (std::string("openxlsx_stream_") + std::to_string(std::chrono::system_clock::now().time_since_epoch().count()) + "_" +
                          []() -> std::string {
                         std::mt19937 rng(std::random_device{}());
                         return std::to_string(rng());
                     }() + ".xml")),
          m_stream(m_tempPath, std::ios::binary),
          m_currentRow(worksheet ? worksheet->rowCount() + 1 : 1),
          m_active(true)
    {
        if (!m_stream.is_open()) throw XLInternalError("Failed to open temporary stream file: " + m_tempPath.string());

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
    { other.m_active = false; }

    XLStreamWriter& XLStreamWriter::operator=(XLStreamWriter&& other) noexcept
    {
        if (this != &other) {
            if (m_active) flushSheetDataClose();
            m_tempPath     = std::move(other.m_tempPath);
            m_stream       = std::move(other.m_stream);
            m_currentRow   = other.m_currentRow;
            m_active       = other.m_active;
            m_bottomHalf   = std::move(other.m_bottomHalf);
            m_writeBuffer  = std::move(other.m_writeBuffer);
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
    template<typename T>
    void XLStreamWriter::appendRowImpl(const std::vector<T>& items)
    {
        if (!isStreamActive()) throw XLInternalError("Stream writer is not active");

        char rowBuf[16];
        auto [rowPtr, _] = std::to_chars(rowBuf, rowBuf + sizeof(rowBuf), m_currentRow);

        m_writeBuffer += "<row r=\"";
        m_writeBuffer.append(rowBuf, rowPtr);
        m_writeBuffer += "\">";

        uint16_t colIdx = 1;
        char cellRefBuf[16];

        for (const auto& item : items) {
            const XLCellValue*          valPtr   = nullptr;
            std::optional<XLStyleIndex> styleIdx = std::nullopt;

            if constexpr (std::is_same_v<T, XLCellValue>) { valPtr = &item; }
            else {
                valPtr   = &item.value;
                styleIdx = item.styleIndex;
            }

            if (valPtr->type() != XLValueType::Empty) {
                makeCellAddress(m_currentRow, colIdx, cellRefBuf);

                m_writeBuffer += "<c r=\"";
                m_writeBuffer += cellRefBuf;
                m_writeBuffer += '"';

                if (styleIdx.has_value() && styleIdx.value() != XLDefaultCellFormat && styleIdx.value() != XLInvalidStyleIndex) {
                    char styleBuf[12];
                    auto [stylePtr, __] = std::to_chars(styleBuf, styleBuf + sizeof(styleBuf), styleIdx.value());
                    m_writeBuffer += " s=\"";
                    m_writeBuffer.append(styleBuf, stylePtr);
                    m_writeBuffer += '"';
                }

                switch (valPtr->type()) {
                    case XLValueType::String:
                        m_writeBuffer += R"( t="inlineStr"><is><t xml:space="preserve">)";
                        appendEscaped(m_writeBuffer, valPtr->get<std::string>());
                        m_writeBuffer += "</t></is></c>";
                        break;
                        
                    case XLValueType::RichText: {
                        m_writeBuffer += R"( t="inlineStr"><is>)";
                        const auto& rt = valPtr->get<XLRichText>();
                        for (const auto& run : rt.runs()) {
                            m_writeBuffer += "<r>";
                            if (run.fontName() || run.fontSize() || run.fontColor() || run.bold() || run.italic() || run.underlineStyle().has_value() || run.strikethrough() || run.vertAlign().has_value()) {
                                m_writeBuffer += "<rPr>";
                                if (run.fontName()) {
                                    m_writeBuffer += R"(<rFont val=")";
                                    appendEscaped(m_writeBuffer, *run.fontName());
                                    m_writeBuffer += R"("/>)";
                                }
                                if (run.fontSize()) {
                                    char szBuf[12];
                                    auto [szPtr, _szEc] = std::to_chars(szBuf, szBuf + sizeof(szBuf), *run.fontSize());
                                    m_writeBuffer += R"(<sz val=")";
                                    m_writeBuffer.append(szBuf, szPtr);
                                    m_writeBuffer += R"("/>)";
                                }
                                if (run.fontColor()) {
                                    m_writeBuffer += R"(<color rgb=")";
                                    m_writeBuffer += run.fontColor()->hex();
                                    m_writeBuffer += R"("/>)";
                                }
                                if (run.bold() && *run.bold()) m_writeBuffer += "<b/>";
                                if (run.italic() && *run.italic()) m_writeBuffer += "<i/>";
                                if (run.underlineStyle().has_value() && run.underlineStyle().value() != XLUnderlineNone && run.underlineStyle().value() != XLUnderlineInvalid) {
                                    m_writeBuffer += "<u";
                                    if (run.underlineStyle().value() == XLUnderlineDouble) m_writeBuffer += R"( val="double")";
                                    else if (run.underlineStyle().value() == XLUnderlineSingleAccounting) m_writeBuffer += R"( val="singleAccounting")";
                                    else if (run.underlineStyle().value() == XLUnderlineDoubleAccounting) m_writeBuffer += R"( val="doubleAccounting")";
                                    else if (run.underlineStyle().value() == XLUnderlineSingle) m_writeBuffer += R"( val="single")";
                                    m_writeBuffer += "/>";
                                }
                                if (run.strikethrough() && *run.strikethrough()) m_writeBuffer += "<strike/>";
                                if (run.vertAlign()) {
                                    if (*run.vertAlign() == XLSuperscript) m_writeBuffer += R"(<vertAlign val="superscript"/>)";
                                    else if (*run.vertAlign() == XLSubscript) m_writeBuffer += R"(<vertAlign val="subscript"/>)";
                                }
                                m_writeBuffer += "</rPr>";
                            }
                            m_writeBuffer += "<t";
                            if (!run.text().empty() && (run.text().front() == ' ' || run.text().back() == ' ')) {
                                m_writeBuffer += R"( xml:space="preserve")";
                            }
                            m_writeBuffer += ">";
                            appendEscaped(m_writeBuffer, run.text());
                            m_writeBuffer += "</t></r>";
                        }
                        m_writeBuffer += "</is></c>";
                        break;
                    }

                    case XLValueType::Boolean:
                        m_writeBuffer += R"( t="b"><v>)";
                        m_writeBuffer += (valPtr->get<bool>() ? '1' : '0');
                        m_writeBuffer += "</v></c>";
                        break;

                    case XLValueType::Integer: {
                        char numBuf[24];
                        auto [numPtr, ___] = std::to_chars(numBuf, numBuf + sizeof(numBuf), valPtr->get<int64_t>());
                        m_writeBuffer += R"( t="n"><v>)";
                        m_writeBuffer.append(numBuf, numPtr);
                        m_writeBuffer += "</v></c>";
                        break;
                    }

                    case XLValueType::Float:
                        m_writeBuffer += R"( t="n"><v>)";
                        fmt::format_to(std::back_inserter(m_writeBuffer), "{}", valPtr->get<double>());
                        m_writeBuffer += "</v></c>";
                        break;

                    default:
                        m_writeBuffer += "><v>";
                        appendEscaped(m_writeBuffer, XLCellValue(*valPtr).getString());
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

    void XLStreamWriter::appendRow(const std::vector<XLCellValue>& values) { appendRowImpl(values); }

    void XLStreamWriter::appendRow(const std::vector<XLStreamCell>& cells) { appendRowImpl(cells); }

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

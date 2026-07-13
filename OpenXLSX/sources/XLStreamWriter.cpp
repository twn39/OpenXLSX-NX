// ===== External Includes ===== //
#include <charconv>
#include <chrono>
#include <fmt/format.h>
#include <fstream>
#include <iterator>
#include <random>
#include <type_traits>
#include <utility>

// ===== OpenXLSX Includes ===== //
#include "XLCellReference.hpp"
#include "XLConstants.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLStreamWriter.hpp"
#include "XLUtilities.hpp"
#include "XLWorksheet.hpp"
#include "XLXmlData.hpp"

namespace OpenXLSX
{
    namespace
    {
        constexpr double kMaxRowHeight = 409.0;

        void validateRowOpts(const XLStreamRowOpts* opts)
        {
            if (!opts) return;
            if (opts->height.has_value() && (*opts->height < 0.0 || *opts->height > kMaxRowHeight)) {
                throw XLInputError("XLStreamWriter: row height must be in range [0, 409]");
            }
            if (opts->outlineLevel.has_value() && *opts->outlineLevel > 7) {
                throw XLInputError("XLStreamWriter: outlineLevel must be in range [0, 7]");
            }
        }
    }    // namespace

    XLStreamWriter::XLStreamWriter(XLWorksheet* worksheet, bool useSharedStrings, size_t maxUniqueStrings)
        : m_tempPath(std::filesystem::temp_directory_path() /
                     (std::string("openxlsx_stream_") + std::to_string(std::chrono::system_clock::now().time_since_epoch().count()) + "_" +
                      []() -> std::string {
                          std::mt19937 rng(std::random_device {}());
                          return std::to_string(rng());
                      }() +
                      ".xml")),
          m_stream(m_tempPath, std::ios::binary),
          m_currentRow(worksheet ? worksheet->rowCount() + 1 : 1),
          m_lastWrittenRow(0),
          m_maxColumn(0),
          m_baseMaxCol(worksheet ? worksheet->columnCount() : 0),
          m_baseLastRow(worksheet ? worksheet->rowCount() : 0),
          m_active(true),
          m_worksheet(worksheet),
          m_useSharedStrings(useSharedStrings),
          m_maxUniqueStrings(maxUniqueStrings)
    {
        if (!m_stream.is_open()) throw XLInternalError("Failed to open temporary stream file: " + m_tempPath.string());

        // Reserve write buffer up-front — avoids reallocations during normal use
        m_writeBuffer.reserve(kFlushThreshold + 64UL * 1024UL);
        markWriterOpen(true);
    }

    XLStreamWriter::~XLStreamWriter()
    {
        try {
            if (m_active) flushSheetDataClose();
        }
        catch (...) {
            // Destructor must not throw; best-effort close already attempted.
            m_active = false;
            markWriterOpen(false);
        }
    }

    XLStreamWriter::XLStreamWriter(XLStreamWriter&& other) noexcept
        : m_tempPath(std::move(other.m_tempPath)),
          m_stream(std::move(other.m_stream)),
          m_currentRow(other.m_currentRow),
          m_lastWrittenRow(other.m_lastWrittenRow),
          m_maxColumn(other.m_maxColumn),
          m_baseMaxCol(other.m_baseMaxCol),
          m_baseLastRow(other.m_baseLastRow),
          m_active(other.m_active),
          m_bottomHalf(std::move(other.m_bottomHalf)),
          m_writeBuffer(std::move(other.m_writeBuffer)),
          m_worksheet(other.m_worksheet),
          m_useSharedStrings(other.m_useSharedStrings),
          m_maxUniqueStrings(other.m_maxUniqueStrings),
          m_wroteFormula(other.m_wroteFormula),
          m_localCache(std::move(other.m_localCache))
    {
        // Ownership of the open-flag moves with the writer object; source is no longer active.
        other.m_active    = false;
        other.m_worksheet = nullptr;
    }

    XLStreamWriter& XLStreamWriter::operator=(XLStreamWriter&& other) noexcept
    {
        if (this != &other) {
            try {
                if (m_active) flushSheetDataClose();
            }
            catch (...) {
                m_active = false;
                markWriterOpen(false);
            }
            m_tempPath         = std::move(other.m_tempPath);
            m_stream           = std::move(other.m_stream);
            m_currentRow       = other.m_currentRow;
            m_lastWrittenRow   = other.m_lastWrittenRow;
            m_maxColumn        = other.m_maxColumn;
            m_baseMaxCol       = other.m_baseMaxCol;
            m_baseLastRow      = other.m_baseLastRow;
            m_active           = other.m_active;
            m_bottomHalf       = std::move(other.m_bottomHalf);
            m_writeBuffer      = std::move(other.m_writeBuffer);
            m_worksheet        = other.m_worksheet;
            m_useSharedStrings = other.m_useSharedStrings;
            m_maxUniqueStrings = other.m_maxUniqueStrings;
            m_wroteFormula     = other.m_wroteFormula;
            m_localCache       = std::move(other.m_localCache);
            other.m_active     = false;
            other.m_worksheet  = nullptr;
        }
        return *this;
    }

    bool        XLStreamWriter::isStreamActive() const { return m_active && m_stream.is_open(); }
    std::string XLStreamWriter::getTempFilePath() const { return m_tempPath.string(); }

    void XLStreamWriter::markWriterOpen(bool open) noexcept
    {
        if (m_worksheet) m_worksheet->setStreamWriterOpen(open);
    }

    void XLStreamWriter::checkStreamWritable() const
    {
        if (!isStreamActive()) throw XLInternalError("Stream writer is not active");
    }

    void XLStreamWriter::writeRowOpen(uint32_t row, const XLStreamRowOpts* opts)
    {
        validateRowOpts(opts);

        char rowBuf[16];
        auto [rowPtr, _] = std::to_chars(rowBuf, rowBuf + sizeof(rowBuf), row);

        m_writeBuffer += "<row r=\"";
        m_writeBuffer.append(rowBuf, rowPtr);
        m_writeBuffer += '"';

        if (opts) {
            if (opts->styleIndex.has_value() && opts->styleIndex.value() != XLDefaultCellFormat &&
                opts->styleIndex.value() != XLInvalidStyleIndex) {
                char styleBuf[12];
                auto [stylePtr, __] = std::to_chars(styleBuf, styleBuf + sizeof(styleBuf), opts->styleIndex.value());
                m_writeBuffer += " s=\"";
                m_writeBuffer.append(styleBuf, stylePtr);
                m_writeBuffer += "\" customFormat=\"1\"";
            }
            if (opts->height.has_value() && *opts->height > 0.0) {
                m_writeBuffer += " ht=\"";
                fmt::format_to(std::back_inserter(m_writeBuffer), "{}", *opts->height);
                m_writeBuffer += "\" customHeight=\"1\"";
            }
            if (opts->outlineLevel.has_value() && *opts->outlineLevel > 0) {
                char lvlBuf[8];
                auto [lvlPtr, ___] = std::to_chars(lvlBuf, lvlBuf + sizeof(lvlBuf), *opts->outlineLevel);
                m_writeBuffer += " outlineLevel=\"";
                m_writeBuffer.append(lvlBuf, lvlPtr);
                m_writeBuffer += '"';
            }
            if (opts->hidden.has_value() && *opts->hidden) { m_writeBuffer += " hidden=\"1\""; }
        }

        m_writeBuffer += '>';
    }

    void XLStreamWriter::writeCellValueBody(const XLCellValue& value)
    {
        switch (value.type()) {
            case XLValueType::String: {
                std::string strVal           = value.get<std::string>();
                bool        writtenAsShared  = false;
                int32_t     sstIdx           = -1;

                if (m_useSharedStrings && m_worksheet) {
                    auto it = m_localCache.find(strVal);
                    if (it != m_localCache.end()) {
                        sstIdx          = it->second;
                        writtenAsShared = true;
                    }
                    else {
                        const XLSharedStringTable& sstView = m_worksheet->sharedStringTable();
                        if (sstView.stringExists(strVal)) {
                            sstIdx          = sstView.getStringIndex(strVal);
                            writtenAsShared = true;
                        }
                        else if (static_cast<size_t>(sstView.stringCount()) < m_maxUniqueStrings) {
                            const auto& sst = m_worksheet->sharedStrings();
                            sstIdx          = sst.getOrCreateStringIndex(strVal);
                            if (m_localCache.size() >= kLocalCacheLimit) { m_localCache.clear(); }
                            m_localCache.emplace(sstView.getStringView(sstIdx), sstIdx);
                            writtenAsShared = true;
                        }
                    }
                }

                if (writtenAsShared) {
                    char idxBuf[12];
                    auto [idxPtr, _] = std::to_chars(idxBuf, idxBuf + sizeof(idxBuf), sstIdx);
                    m_writeBuffer += R"( t="s"><v>)";
                    m_writeBuffer.append(idxBuf, idxPtr);
                    m_writeBuffer += "</v></c>";
                }
                else {
                    m_writeBuffer += R"( t="inlineStr"><is><t xml:space="preserve">)";
                    appendEscaped(m_writeBuffer, strVal);
                    m_writeBuffer += "</t></is></c>";
                }
                break;
            }
            case XLValueType::RichText: {
                m_writeBuffer += R"( t="inlineStr"><is>)";
                const auto& rt = value.get<XLRichText>();
                for (const auto& run : rt.runs()) {
                    m_writeBuffer += "<r>";
                    if (run.fontName() || run.fontSize() || run.fontColor() || run.bold() || run.italic() ||
                        run.underlineStyle().has_value() || run.strikethrough() || run.vertAlign().has_value()) {
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
                        if (run.underlineStyle().has_value() && run.underlineStyle().value() != XLUnderlineNone &&
                            run.underlineStyle().value() != XLUnderlineInvalid) {
                            m_writeBuffer += "<u";
                            if (run.underlineStyle().value() == XLUnderlineDouble) m_writeBuffer += R"( val="double")";
                            else if (run.underlineStyle().value() == XLUnderlineSingleAccounting)
                                m_writeBuffer += R"( val="singleAccounting")";
                            else if (run.underlineStyle().value() == XLUnderlineDoubleAccounting)
                                m_writeBuffer += R"( val="doubleAccounting")";
                            else if (run.underlineStyle().value() == XLUnderlineSingle)
                                m_writeBuffer += R"( val="single")";
                            m_writeBuffer += "/>";
                        }
                        if (run.strikethrough() && *run.strikethrough()) m_writeBuffer += "<strike/>";
                        if (run.vertAlign()) {
                            if (*run.vertAlign() == XLSuperscript) m_writeBuffer += R"(<vertAlign val="superscript"/>)";
                            else if (*run.vertAlign() == XLSubscript)
                                m_writeBuffer += R"(<vertAlign val="subscript"/>)";
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
                m_writeBuffer += (value.get<bool>() ? '1' : '0');
                m_writeBuffer += "</v></c>";
                break;
            case XLValueType::Integer: {
                char numBuf[24];
                auto [numPtr, ___] = std::to_chars(numBuf, numBuf + sizeof(numBuf), value.get<int64_t>());
                m_writeBuffer += R"( t="n"><v>)";
                m_writeBuffer.append(numBuf, numPtr);
                m_writeBuffer += "</v></c>";
                break;
            }
            case XLValueType::Float:
                m_writeBuffer += R"( t="n"><v>)";
                fmt::format_to(std::back_inserter(m_writeBuffer), "{}", value.get<double>());
                m_writeBuffer += "</v></c>";
                break;
            case XLValueType::Error: {
                m_writeBuffer += R"( t="e"><v>)";
                XLCellValue errCopy = value;
                appendEscaped(m_writeBuffer, errCopy.getString());
                m_writeBuffer += "</v></c>";
                break;
            }
            default:
                m_writeBuffer += "><v>";
                appendEscaped(m_writeBuffer, XLCellValue(value).getString());
                m_writeBuffer += "</v></c>";
                break;
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  setRowImpl — shared path for appendRow / setRow
    // ─────────────────────────────────────────────────────────────────────────
    template<typename T>
    void XLStreamWriter::setRowImpl(uint32_t row, uint16_t startCol, const std::vector<T>& items, const XLStreamRowOpts* opts)
    {
        checkStreamWritable();

        if (row < 1 || row > MAX_ROWS) throw XLInputError("XLStreamWriter: row number out of range");
        if (startCol < 1 || startCol > MAX_COLS) throw XLInputError("XLStreamWriter: column number out of range");
        if (row <= m_lastWrittenRow) {
            throw XLInputError("XLStreamWriter: row numbers must be strictly increasing (excelize SetRow order)");
        }

        writeRowOpen(row, opts);

        uint16_t colIdx = startCol;
        char     cellRefBuf[16];

        for (const auto& item : items) {
            const XLCellValue*          valPtr   = nullptr;
            std::optional<XLStyleIndex> styleIdx = std::nullopt;
            const std::string*          formula  = nullptr;

            if constexpr (std::is_same_v<T, XLCellValue>) { valPtr = &item; }
            else {
                valPtr   = &item.value;
                styleIdx = item.styleIndex;
                if (item.formula.has_value() && !item.formula->empty()) formula = &item.formula.value();
            }

            const bool hasFormula = formula != nullptr;
            const bool hasValue   = valPtr->type() != XLValueType::Empty;

            if (hasFormula || hasValue) {
                makeCellAddress(row, colIdx, cellRefBuf);

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

                if (hasFormula) {
                    m_wroteFormula = true;
                    // When only a formula is present (no cached value), close after <f>.
                    if (!hasValue) {
                        m_writeBuffer += "><f>";
                        appendEscaped(m_writeBuffer, *formula);
                        m_writeBuffer += "</f></c>";
                    }
                    else {
                        // Emit type + value body, but inject <f> immediately after attributes open.
                        // writeCellValueBody emits ' t="…"><v>…' or similar starting with space or '>';
                        // For formula cells with cached value, write t/v via a custom path:
                        switch (valPtr->type()) {
                            case XLValueType::Integer: {
                                char numBuf[24];
                                auto [numPtr, ___] = std::to_chars(numBuf, numBuf + sizeof(numBuf), valPtr->get<int64_t>());
                                m_writeBuffer += R"( t="n"><f>)";
                                appendEscaped(m_writeBuffer, *formula);
                                m_writeBuffer += "</f><v>";
                                m_writeBuffer.append(numBuf, numPtr);
                                m_writeBuffer += "</v></c>";
                                break;
                            }
                            case XLValueType::Float: {
                                m_writeBuffer += R"( t="n"><f>)";
                                appendEscaped(m_writeBuffer, *formula);
                                m_writeBuffer += "</f><v>";
                                fmt::format_to(std::back_inserter(m_writeBuffer), "{}", valPtr->get<double>());
                                m_writeBuffer += "</v></c>";
                                break;
                            }
                            case XLValueType::Boolean: {
                                m_writeBuffer += R"( t="b"><f>)";
                                appendEscaped(m_writeBuffer, *formula);
                                m_writeBuffer += "</f><v>";
                                m_writeBuffer += (valPtr->get<bool>() ? '1' : '0');
                                m_writeBuffer += "</v></c>";
                                break;
                            }
                            case XLValueType::String: {
                                m_writeBuffer += R"( t="str"><f>)";
                                appendEscaped(m_writeBuffer, *formula);
                                m_writeBuffer += "</f><v>";
                                appendEscaped(m_writeBuffer, valPtr->get<std::string>());
                                m_writeBuffer += "</v></c>";
                                break;
                            }
                            default: {
                                m_writeBuffer += "><f>";
                                appendEscaped(m_writeBuffer, *formula);
                                m_writeBuffer += "</f><v>";
                                appendEscaped(m_writeBuffer, XLCellValue(*valPtr).getString());
                                m_writeBuffer += "</v></c>";
                                break;
                            }
                        }
                    }
                }
                else {
                    writeCellValueBody(*valPtr);
                }

                if (colIdx > m_maxColumn) m_maxColumn = colIdx;
            }
            ++colIdx;
        }

        m_writeBuffer += "</row>";
        m_lastWrittenRow = row;
        m_currentRow     = row + 1;

        if (m_writeBuffer.size() >= kFlushThreshold) flushWriteBuffer();
    }

    void XLStreamWriter::appendRow(const std::vector<XLCellValue>& values) { setRowImpl(m_currentRow, 1, values, nullptr); }

    void XLStreamWriter::appendRow(const std::vector<XLCellValue>& values, const XLStreamRowOpts& opts)
    {
        setRowImpl(m_currentRow, 1, values, &opts);
    }

    void XLStreamWriter::appendRow(std::initializer_list<XLCellValue> values)
    {
        appendRow(std::vector<XLCellValue>(values));
    }

    void XLStreamWriter::appendRow(std::initializer_list<XLCellValue> values, const XLStreamRowOpts& opts)
    {
        appendRow(std::vector<XLCellValue>(values), opts);
    }

    void XLStreamWriter::appendRow(const std::vector<XLStreamCell>& cells) { setRowImpl(m_currentRow, 1, cells, nullptr); }

    void XLStreamWriter::appendRow(const std::vector<XLStreamCell>& cells, const XLStreamRowOpts& opts)
    {
        setRowImpl(m_currentRow, 1, cells, &opts);
    }

    void XLStreamWriter::appendRow(std::initializer_list<XLStreamCell> cells)
    {
        appendRow(std::vector<XLStreamCell>(cells));
    }

    void XLStreamWriter::appendRow(std::initializer_list<XLStreamCell> cells, const XLStreamRowOpts& opts)
    {
        appendRow(std::vector<XLStreamCell>(cells), opts);
    }

    void XLStreamWriter::setRow(uint32_t row, uint16_t startCol, const std::vector<XLStreamCell>& cells, const XLStreamRowOpts& opts)
    {
        setRowImpl(row, startCol, cells, &opts);
    }

    void XLStreamWriter::setRow(uint32_t row, uint16_t startCol, const std::vector<XLCellValue>& values, const XLStreamRowOpts& opts)
    {
        setRowImpl(row, startCol, values, &opts);
    }

    void XLStreamWriter::setRow(uint32_t row, uint16_t startCol, std::initializer_list<XLStreamCell> cells, const XLStreamRowOpts& opts)
    {
        setRow(row, startCol, std::vector<XLStreamCell>(cells), opts);
    }

    void XLStreamWriter::setRow(uint32_t row, uint16_t startCol, std::initializer_list<XLCellValue> values, const XLStreamRowOpts& opts)
    {
        setRow(row, startCol, std::vector<XLCellValue>(values), opts);
    }

    void XLStreamWriter::setRow(const std::string& cellRef, const std::vector<XLStreamCell>& cells, const XLStreamRowOpts& opts)
    {
        XLCellReference ref(cellRef);
        setRowImpl(ref.row(), ref.column(), cells, &opts);
    }

    void XLStreamWriter::setRow(const std::string& cellRef, const std::vector<XLCellValue>& values, const XLStreamRowOpts& opts)
    {
        XLCellReference ref(cellRef);
        setRowImpl(ref.row(), ref.column(), values, &opts);
    }

    void XLStreamWriter::setRow(const std::string& cellRef, std::initializer_list<XLStreamCell> cells, const XLStreamRowOpts& opts)
    {
        setRow(cellRef, std::vector<XLStreamCell>(cells), opts);
    }

    void XLStreamWriter::setRow(const std::string& cellRef, std::initializer_list<XLCellValue> values, const XLStreamRowOpts& opts)
    {
        setRow(cellRef, std::vector<XLCellValue>(values), opts);
    }

    void XLStreamWriter::close()
    {
        if (m_active) flushSheetDataClose();
    }

    void XLStreamWriter::flushWriteBuffer()
    {
        if (!m_writeBuffer.empty() && m_stream.is_open()) {
            m_stream.write(m_writeBuffer.data(), static_cast<std::streamsize>(m_writeBuffer.size()));
            if (!m_stream.good()) {
                m_writeBuffer.clear();
                throw XLInternalError("XLStreamWriter: failed to write buffer to temporary stream file: " + m_tempPath.string());
            }
            m_writeBuffer.clear();
        }
    }

    void XLStreamWriter::patchDimensionInTempFile()
    {
        // Determine used range: include DOM base rows/cols plus stream-written extents.
        const uint32_t endRow = (m_lastWrittenRow > 0) ? m_lastWrittenRow : m_baseLastRow;
        const uint16_t endCol = (m_maxColumn > m_baseMaxCol) ? m_maxColumn : m_baseMaxCol;

        std::string dimRef = "A1";
        if (endRow > 0 && endCol > 0) {
            dimRef = "A1:" + XLCellReference::columnAsString(endCol) + std::to_string(endRow);
        }

        const std::string dimXml = "<dimension ref=\"" + dimRef + "\"/>";

        std::ifstream in(m_tempPath, std::ios::binary);
        if (!in) return;
        std::string content((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
        in.close();

        const std::string dimTag = "<dimension";
        size_t            pos    = content.find(dimTag);
        if (pos != std::string::npos) {
            size_t end = content.find("/>", pos);
            if (end == std::string::npos) {
                end = content.find('>', pos);
                if (end != std::string::npos) {
                    // Possibly <dimension ...></dimension>
                    size_t close = content.find("</dimension>", end);
                    if (close != std::string::npos) end = close + 12;
                    else
                        ++end;
                }
            }
            else {
                end += 2;
            }
            if (end != std::string::npos && end > pos) { content.replace(pos, end - pos, dimXml); }
        }
        else {
            // Insert before <sheetData if possible (OOXML node order).
            size_t sheetDataPos = content.find("<sheetData");
            if (sheetDataPos != std::string::npos) { content.insert(sheetDataPos, dimXml); }
        }

        std::ofstream out(m_tempPath, std::ios::binary | std::ios::trunc);
        if (!out) throw XLInternalError("XLStreamWriter: failed to rewrite temp file for dimension patch");
        out.write(content.data(), static_cast<std::streamsize>(content.size()));
        if (!out.good()) throw XLInternalError("XLStreamWriter: failed while writing dimension-patched temp file");
        out.close();

        if (m_worksheet) m_worksheet->setStreamExtents(endRow, endCol);
    }

    void XLStreamWriter::flushSheetDataClose()
    {
        if (m_stream.is_open()) {
            flushWriteBuffer();
            m_stream << m_bottomHalf;
            if (!m_stream.good()) {
                m_active = false;
                markWriterOpen(false);
                throw XLInternalError("XLStreamWriter: failed to write sheet footer to temporary stream file");
            }
            m_stream.close();
        }

        // Patch dimension while temp file is complete but still marked streamed.
        try {
            patchDimensionInTempFile();
        }
        catch (...) {
            m_active = false;
            markWriterOpen(false);
            throw;
        }

        if (m_wroteFormula && m_worksheet) {
            try {
                m_worksheet->parentDoc().setFormulaNeedsRecalculation(true);
            }
            catch (...) {
                // Non-fatal: package remains valid; Excel will still recalculate on open in many cases.
            }
        }

        m_active = false;
        markWriterOpen(false);
    }

    // Explicit template instantiations for the two public element types.
    template void XLStreamWriter::setRowImpl<XLCellValue>(uint32_t, uint16_t, const std::vector<XLCellValue>&, const XLStreamRowOpts*);
    template void XLStreamWriter::setRowImpl<XLStreamCell>(uint32_t, uint16_t, const std::vector<XLStreamCell>&, const XLStreamRowOpts*);

}    // namespace OpenXLSX

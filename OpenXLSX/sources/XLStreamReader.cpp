// ===== External Includes ===== //
#include <cctype>
#include <chrono>
#include <cstring>
#include <fast_float/fast_float.h>
#include <fstream>
#include <random>
#include <string_view>
#include <unordered_map>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLNumberFormatter.hpp"
#include "XLSharedStrings.hpp"
#include "XLStreamReader.hpp"
#include "XLStyles.hpp"
#include "XLWorksheet.hpp"

namespace
{
    void xmlUnescape(std::string& s)
    {
        size_t w = 0;
        for (size_t r = 0; r < s.size();) {
            if (s[r] == '&') {
                if (s.compare(r, 5, "&amp;") == 0) {
                    s[w++] = '&';
                    r += 5;
                }
                else if (s.compare(r, 4, "&lt;") == 0) {
                    s[w++] = '<';
                    r += 4;
                }
                else if (s.compare(r, 4, "&gt;") == 0) {
                    s[w++] = '>';
                    r += 4;
                }
                else if (s.compare(r, 6, "&quot;") == 0) {
                    s[w++] = '"';
                    r += 6;
                }
                else if (s.compare(r, 6, "&apos;") == 0) {
                    s[w++] = '\'';
                    r += 6;
                }
                else {
                    s[w++] = s[r++];
                }
            }
            else {
                s[w++] = s[r++];
            }
        }
        s.resize(w);
    }

    size_t findRowStart(const std::string& buf, size_t from)
    {
        // Match <row …> or namespaced <x:row …> / <ns:row …> (local name must be exactly "row").
        size_t pos = from;
        while (pos < buf.size()) {
            pos = buf.find('<', pos);
            if (pos == std::string::npos) return pos;
            if (pos + 1 >= buf.size()) return std::string::npos;
            if (buf[pos + 1] == '/' || buf[pos + 1] == '!' || buf[pos + 1] == '?') {
                ++pos;
                continue;
            }
            size_t nameStart = pos + 1;
            size_t nameEnd   = nameStart;
            while (nameEnd < buf.size() && buf[nameEnd] != ' ' && buf[nameEnd] != '\t' && buf[nameEnd] != '\r' &&
                   buf[nameEnd] != '\n' && buf[nameEnd] != '>' && buf[nameEnd] != '/')
                ++nameEnd;
            if (nameEnd == nameStart) {
                ++pos;
                continue;
            }
            std::string_view name(buf.data() + nameStart, nameEnd - nameStart);
            const size_t     colon = name.rfind(':');
            if (colon != std::string_view::npos) name = name.substr(colon + 1);
            if (name == "row") return pos;
            pos = nameEnd;
        }
        return std::string::npos;
    }

    uint16_t columnFromCellRef(const std::string& cellRef)
    {
        uint16_t col = 0;
        for (char c : cellRef) {
            if (c >= 'A' && c <= 'Z')
                col = static_cast<uint16_t>(col * 26 + (c - 'A' + 1));
            else if (c >= 'a' && c <= 'z')
                col = static_cast<uint16_t>(col * 26 + (c - 'a' + 1));
            else
                break;
        }
        return col;
    }

    const char* skipWS(const char* cp, const char* end)
    {
        while (cp < end && (*cp == ' ' || *cp == '\t' || *cp == '\r' || *cp == '\n')) ++cp;
        return cp;
    }
}    // anonymous namespace

namespace OpenXLSX
{

    XLStreamReader::XLStreamReader(const XLWorksheet* worksheet, XLStreamReadOptions options)
        : m_worksheet(worksheet), m_options(std::move(options))
    {
        if (!worksheet) throw XLInternalError("Worksheet is null");

        m_archive   = worksheet->package().archive();
        m_zipStream = m_archive.openEntryStream(worksheet->getXmlPath());
        if (!m_zipStream) throw XLInternalError("Failed to open worksheet zip stream");

        m_tagNameBuf.reserve(64);
        m_attrNameBuf.reserve(32);
        m_attrValueBuf.reserve(128);
        m_textContentBuf.reserve(256);
        m_pendingCells.reserve(64);
        m_reuseValues.reserve(64);
        m_reuseDetailed.reserve(64);
        m_reuseStrings.reserve(64);
    }

    XLStreamReader::~XLStreamReader() { cleanup(); }

    XLStreamReader::XLStreamReader(XLStreamReader&& other) noexcept
        : m_worksheet(other.m_worksheet),
          m_archive(std::move(other.m_archive)),
          m_zipStream(other.m_zipStream),
          m_options(std::move(other.m_options)),
          m_sstIndexReady(other.m_sstIndexReady),
          m_sstXml(std::move(other.m_sstXml)),
          m_sstTempPath(std::move(other.m_sstTempPath)),
          m_sstFileBacked(other.m_sstFileBacked),
          m_sstOffsets(std::move(other.m_sstOffsets)),
          m_buffer(std::move(other.m_buffer)),
          m_bufPos(other.m_bufPos),
          m_eof(other.m_eof),
          m_currentRow(other.m_currentRow),
          m_lastPhysicalRow(other.m_lastPhysicalRow),
          m_hasPendingPhysical(other.m_hasPendingPhysical),
          m_physicalExhausted(other.m_physicalExhausted),
          m_pendingRowNum(other.m_pendingRowNum),
          m_pendingRowOpts(std::move(other.m_pendingRowOpts)),
          m_pendingCells(std::move(other.m_pendingCells)),
          m_currentRowOpts(std::move(other.m_currentRowOpts)),
          m_tagNameBuf(std::move(other.m_tagNameBuf)),
          m_attrNameBuf(std::move(other.m_attrNameBuf)),
          m_attrValueBuf(std::move(other.m_attrValueBuf)),
          m_textContentBuf(std::move(other.m_textContentBuf)),
          m_lastError(std::move(other.m_lastError)),
          m_reuseValues(std::move(other.m_reuseValues)),
          m_reuseDetailed(std::move(other.m_reuseDetailed)),
          m_reuseStrings(std::move(other.m_reuseStrings))
    {
        other.m_worksheet          = nullptr;
        other.m_zipStream          = nullptr;
        other.m_eof                = true;
        other.m_hasPendingPhysical = false;
        other.m_physicalExhausted  = true;
        other.m_sstIndexReady      = false;
        other.m_sstFileBacked      = false;
    }

    XLStreamReader& XLStreamReader::operator=(XLStreamReader&& other) noexcept
    {
        if (this != &other) {
            cleanup();
            m_worksheet          = other.m_worksheet;
            m_archive            = std::move(other.m_archive);
            m_zipStream          = other.m_zipStream;
            m_options            = std::move(other.m_options);
            m_sstIndexReady      = other.m_sstIndexReady;
            m_sstXml             = std::move(other.m_sstXml);
            m_sstTempPath        = std::move(other.m_sstTempPath);
            m_sstFileBacked      = other.m_sstFileBacked;
            m_sstOffsets         = std::move(other.m_sstOffsets);
            m_buffer             = std::move(other.m_buffer);
            m_bufPos             = other.m_bufPos;
            m_eof                = other.m_eof;
            m_currentRow         = other.m_currentRow;
            m_lastPhysicalRow    = other.m_lastPhysicalRow;
            m_hasPendingPhysical = other.m_hasPendingPhysical;
            m_physicalExhausted  = other.m_physicalExhausted;
            m_pendingRowNum      = other.m_pendingRowNum;
            m_pendingRowOpts     = std::move(other.m_pendingRowOpts);
            m_pendingCells       = std::move(other.m_pendingCells);
            m_currentRowOpts     = std::move(other.m_currentRowOpts);
            m_tagNameBuf         = std::move(other.m_tagNameBuf);
            m_attrNameBuf        = std::move(other.m_attrNameBuf);
            m_attrValueBuf       = std::move(other.m_attrValueBuf);
            m_textContentBuf     = std::move(other.m_textContentBuf);
            m_lastError          = std::move(other.m_lastError);
            m_reuseValues        = std::move(other.m_reuseValues);
            m_reuseDetailed      = std::move(other.m_reuseDetailed);
            m_reuseStrings       = std::move(other.m_reuseStrings);

            other.m_worksheet          = nullptr;
            other.m_zipStream          = nullptr;
            other.m_eof                = true;
            other.m_hasPendingPhysical = false;
            other.m_physicalExhausted  = true;
            other.m_sstIndexReady      = false;
            other.m_sstFileBacked      = false;
        }
        return *this;
    }

    void XLStreamReader::cleanup()
    {
        if (m_zipStream) {
            // Archive copy shares LibZipApp with the document. closeEntryStream is idempotent if
            // XLZipArchive::close() already closed this entry (important on Windows file locks).
            try {
                m_archive.closeEntryStream(m_zipStream);
            }
            catch (...) {
                // Best-effort; document may already be torn down.
            }
            m_zipStream = nullptr;
        }
        if (m_sstFileBacked && !m_sstTempPath.empty()) {
            std::error_code ec;
            std::filesystem::remove(m_sstTempPath, ec);
            m_sstTempPath.clear();
            m_sstFileBacked = false;
        }
        m_sstXml.clear();
        m_sstOffsets.clear();
        m_sstIndexReady = false;
    }

    void XLStreamReader::compactBufferIfNeeded()
    {
        constexpr size_t kCompactThreshold = 64 * 1024;
        if (m_bufPos > kCompactThreshold && m_bufPos > m_buffer.size() / 2) {
            m_buffer.erase(0, m_bufPos);
            m_bufPos = 0;
        }
    }

    void XLStreamReader::fetchMoreData()
    {
        if (m_eof || !m_zipStream) return;

        compactBufferIfNeeded();

        constexpr size_t kReadBuf = 65536;
        char             buf[kReadBuf];
        auto             bytesRead = m_archive.readEntryStream(m_zipStream, buf, kReadBuf);

        if (bytesRead > 0)
            m_buffer.append(buf, static_cast<size_t>(bytesRead));
        else {
            m_eof = true;
            cleanup();
        }
    }

    bool XLStreamReader::hasNext()
    {
        if (m_hasPendingPhysical) {
            if (m_options.emptyRows == XLStreamEmptyRowPolicy::EmitEmptyRows && m_currentRow + 1 < m_pendingRowNum) return true;
            return true;
        }
        if (m_physicalExhausted) return false;
        return loadNextPhysicalRow();
    }

    bool XLStreamReader::loadNextPhysicalRow()
    {
        if (m_hasPendingPhysical) return true;
        if (m_physicalExhausted) return false;

        while (true) {
            size_t rowStart = findRowStart(m_buffer, m_bufPos);
            if (rowStart == std::string::npos) {
                if (m_eof) {
                    m_physicalExhausted = true;
                    return false;
                }
                // Keep a small tail in case "<row" straddles a chunk boundary.
                if (m_buffer.size() > m_bufPos + 5) {
                    // Drop fully-scanned prefix but keep last 5 bytes.
                    const size_t keepFrom = m_buffer.size() > 5 ? m_buffer.size() - 5 : m_bufPos;
                    if (keepFrom > m_bufPos) {
                        m_buffer.erase(0, keepFrom);
                        m_bufPos = 0;
                    }
                }
                fetchMoreData();
                continue;
            }

            size_t firstGt = m_buffer.find('>', rowStart);
            if (firstGt == std::string::npos) {
                if (m_eof) {
                    m_physicalExhausted = true;
                    return false;
                }
                fetchMoreData();
                continue;
            }

            const bool selfClosingRow = (firstGt > 0) && (m_buffer[firstGt - 1] == '/');
            if (selfClosingRow) {
                const size_t rowBytes = firstGt + 1 - rowStart;
                if (m_options.maxRowBytes > 0 && rowBytes > m_options.maxRowBytes) {
                    throw XLInputError("XLStreamReader: physical row exceeds maxRowBytes limit (" +
                                       std::to_string(rowBytes) + " > " + std::to_string(m_options.maxRowBytes) + ")");
                }
                parsePhysicalRowSlice(m_buffer.data() + rowStart, m_buffer.data() + firstGt + 1);
                m_bufPos             = firstGt + 1;
                m_hasPendingPhysical = true;
                return true;
            }

            size_t endTag = m_buffer.find("</row>", rowStart);
            if (endTag == std::string::npos) {
                if (m_eof) {
                    m_physicalExhausted = true;
                    return false;
                }
                fetchMoreData();
                continue;
            }

            const size_t sliceEnd = endTag + 6;
            const size_t rowBytes = sliceEnd - rowStart;
            if (m_options.maxRowBytes > 0 && rowBytes > m_options.maxRowBytes) {
                throw XLInputError("XLStreamReader: physical row exceeds maxRowBytes limit (" +
                                   std::to_string(rowBytes) + " > " + std::to_string(m_options.maxRowBytes) + ")");
            }
            parsePhysicalRowSlice(m_buffer.data() + rowStart, m_buffer.data() + sliceEnd);
            m_bufPos             = sliceEnd;
            m_hasPendingPhysical = true;
            return true;
        }
    }

    void XLStreamReader::parsePhysicalRowSlice(const char* p, const char* end)
    {
        m_pendingCells.clear();
        m_pendingRowOpts          = {};
        m_pendingRowNum           = 0;
        m_pendingRowOpts.isSyntheticEmpty = false;

        std::string cellRef;
        std::string cellType;
        std::string cellStyle;
        std::string cellValue;
        std::string cellFormula;
        std::string formulaTypeAttr;
        std::string formulaRefAttr;
        std::string formulaSiAttr;
        bool        formulaCaAttr = false;
        bool        inVTag        = false;
        bool        inIsTTag      = false;
        bool        inFTag        = false;
        bool        inRun         = false;
        bool        inRPr         = false;
        std::string runText;
        XLRichText  richAccum;
        bool        hasRichRuns = false;
        uint16_t    expectedCol = 1;

        auto flushCell = [&]() {
            if (cellRef.empty() && cellValue.empty() && cellFormula.empty() && cellStyle.empty() && cellType.empty() &&
                !hasRichRuns)
                return;

            uint16_t actualCol = expectedCol;
            if (!cellRef.empty()) {
                uint16_t col = columnFromCellRef(cellRef);
                if (col > 0) actualCol = col;
            }

            while (expectedCol < actualCol) {
                XLStreamCellView empty;
                empty.column = expectedCol;
                m_pendingCells.push_back(std::move(empty));
                ++expectedCol;
            }

            xmlUnescape(cellValue);
            if (!cellFormula.empty()) xmlUnescape(cellFormula);

            XLStreamCellView view;
            view.column = actualCol;
            view.value  = parseCellValue(cellType, cellValue);
            if (!cellFormula.empty() || !formulaTypeAttr.empty() || !formulaSiAttr.empty()) {
                view.formula = cellFormula;
            }
            if (!formulaTypeAttr.empty()) view.formulaType = formulaTypeAttr;
            if (!formulaRefAttr.empty()) view.formulaRef = formulaRefAttr;
            if (!formulaSiAttr.empty()) {
                char* ep = nullptr;
                auto  si = static_cast<int>(std::strtol(formulaSiAttr.c_str(), &ep, 10));
                if (ep != formulaSiAttr.c_str()) view.formulaSi = si;
            }
            view.formulaCa = formulaCaAttr;
            if (!cellStyle.empty()) {
                char* ep = nullptr;
                auto  s  = static_cast<XLStyleIndex>(std::strtoul(cellStyle.c_str(), &ep, 10));
                if (ep != cellStyle.c_str()) view.styleIndex = s;
            }
            if (hasRichRuns && m_options.parseRichText) {
                view.richText = richAccum;
                if (view.value.type() == XLValueType::Empty && !cellValue.empty())
                    view.value = XLCellValue(cellValue);
            }
            m_pendingCells.push_back(std::move(view));
            ++expectedCol;

            cellRef.clear();
            cellType.clear();
            cellStyle.clear();
            cellValue.clear();
            cellFormula.clear();
            formulaTypeAttr.clear();
            formulaRefAttr.clear();
            formulaSiAttr.clear();
            formulaCaAttr = false;
            inVTag        = false;
            inIsTTag      = false;
            inFTag        = false;
            inRun         = false;
            inRPr         = false;
            runText.clear();
            richAccum   = XLRichText{};
            hasRichRuns = false;
        };

        while (p < end) {
            if (*p != '<') {
                if (inVTag || inIsTTag) {
                    cellValue += *p;
                    if (inRun) runText += *p;
                }
                else if (inFTag)
                    cellFormula += *p;
                ++p;
                continue;
            }

            ++p;
            if (p >= end) break;

            const bool isClose = (*p == '/');
            if (isClose) ++p;

            m_tagNameBuf.clear();
            while (p < end && *p != ' ' && *p != '>' && *p != '/') m_tagNameBuf += *p++;

            // Strip namespace prefix if present (e.g. x:row)
            {
                size_t colon = m_tagNameBuf.find(':');
                if (colon != std::string::npos) m_tagNameBuf.erase(0, colon + 1);
            }

            if (isClose) {
                if (m_tagNameBuf == "row" || m_tagNameBuf == "c")
                    flushCell();
                else if (m_tagNameBuf == "v")
                    inVTag = false;
                else if (m_tagNameBuf == "t")
                    inIsTTag = false;
                else if (m_tagNameBuf == "f")
                    inFTag = false;
                else if (m_tagNameBuf == "r") {
                    if (m_options.parseRichText && inRun) {
                        xmlUnescape(runText);
                        XLRichTextRun run(runText);
                        richAccum.addRun(run);
                        hasRichRuns = true;
                        runText.clear();
                    }
                    inRun = false;
                    inRPr = false;
                }
                else if (m_tagNameBuf == "rPr")
                    inRPr = false;
                while (p < end && *p != '>') ++p;
                if (p < end) ++p;
                continue;
            }

            std::string localR, localT, localS, localHt, localHidden, localOutline;
            std::string localFType, localFRef, localFSi, localFCa;
            std::string localVal;    // generic val= for rFont/sz/color
            bool        selfClose = false;

            while (p < end) {
                p = skipWS(p, end);
                if (p >= end) break;
                if (*p == '>') {
                    ++p;
                    break;
                }
                if (*p == '/') {
                    selfClose = true;
                    ++p;
                    if (p < end && *p == '>') ++p;
                    break;
                }

                m_attrNameBuf.clear();
                while (p < end && *p != '=' && *p != ' ' && *p != '>' && *p != '/') m_attrNameBuf += *p++;
                // strip namespace on attributes
                {
                    size_t colon = m_attrNameBuf.find(':');
                    if (colon != std::string::npos) m_attrNameBuf.erase(0, colon + 1);
                }
                p = skipWS(p, end);
                if (p >= end || *p != '=') continue;
                ++p;
                p = skipWS(p, end);

                m_attrValueBuf.clear();
                char q = 0;
                if (p < end && (*p == '"' || *p == '\'')) q = *p++;
                while (p < end && (q ? *p != q : (*p != ' ' && *p != '>'))) m_attrValueBuf += *p++;
                if (q && p < end) ++p;

                if (m_attrNameBuf == "r")
                    localR = m_attrValueBuf;
                else if (m_attrNameBuf == "t")
                    localT = m_attrValueBuf;
                else if (m_attrNameBuf == "s")
                    localS = m_attrValueBuf;
                else if (m_attrNameBuf == "ht")
                    localHt = m_attrValueBuf;
                else if (m_attrNameBuf == "hidden")
                    localHidden = m_attrValueBuf;
                else if (m_attrNameBuf == "outlineLevel")
                    localOutline = m_attrValueBuf;
                else if (m_attrNameBuf == "ref")
                    localFRef = m_attrValueBuf;
                else if (m_attrNameBuf == "si")
                    localFSi = m_attrValueBuf;
                else if (m_attrNameBuf == "ca")
                    localFCa = m_attrValueBuf;
                else if (m_attrNameBuf == "val" || m_attrNameBuf == "rgb")
                    localVal = m_attrValueBuf;
            }

            // For <f t="array"> the attribute name is t — reuse localT only when tag is f
            if (m_tagNameBuf == "f" && !localT.empty()) localFType = localT;

            if (m_tagNameBuf == "row") {
                if (!localR.empty())
                    m_pendingRowNum = static_cast<uint32_t>(std::strtoul(localR.c_str(), nullptr, 10));
                else
                    m_pendingRowNum = m_lastPhysicalRow + 1;

                if (!localHt.empty()) {
                    double ht = 0.0;
                    auto [ptr, ec] = fast_float::from_chars(localHt.data(), localHt.data() + localHt.size(), ht);
                    if (ec == std::errc()) m_pendingRowOpts.height = ht;
                }
                if (!localHidden.empty()) {
                    m_pendingRowOpts.hidden = (localHidden == "1" || localHidden == "true");
                }
                if (!localOutline.empty()) {
                    auto lvl = static_cast<uint8_t>(std::strtoul(localOutline.c_str(), nullptr, 10));
                    m_pendingRowOpts.outlineLevel = lvl;
                }
                if (!localS.empty()) {
                    char* ep = nullptr;
                    auto  s  = static_cast<XLStyleIndex>(std::strtoul(localS.c_str(), &ep, 10));
                    if (ep != localS.c_str()) m_pendingRowOpts.styleIndex = s;
                }
            }
            else if (m_tagNameBuf == "c") {
                flushCell();
                cellRef   = localR;
                cellType  = localT;
                cellStyle = localS;
                if (selfClose) flushCell();
            }
            else if (m_tagNameBuf == "v") {
                inVTag   = true;
                inIsTTag = false;
                inFTag   = false;
            }
            else if (m_tagNameBuf == "t") {
                // <t> text in inlineStr or rich run
                inIsTTag = true;
                inVTag   = false;
                inFTag   = false;
            }
            else if (m_tagNameBuf == "f") {
                inFTag   = true;
                inVTag   = false;
                inIsTTag = false;
                if (!localFType.empty()) formulaTypeAttr = localFType;
                if (!localFRef.empty()) formulaRefAttr = localFRef;
                if (!localFSi.empty()) formulaSiAttr = localFSi;
                if (!localFCa.empty()) formulaCaAttr = (localFCa == "1" || localFCa == "true");
                if (selfClose) inFTag = false;
            }
            else if (m_tagNameBuf == "r" && m_options.parseRichText) {
                inRun   = true;
                inRPr   = false;
                runText.clear();
            }
            else if (m_tagNameBuf == "rPr" && m_options.parseRichText) {
                inRPr = true;
            }
            else if (m_options.parseRichText && inRPr && selfClose) {
                // Font property stubs on current run text buffer — applied when run closes via last run props.
                // Minimal: bold/italic/strike empty elements; rFont/sz/color with val.
                (void)localVal;
            }
        }

        flushCell();
        if (m_pendingRowNum == 0) m_pendingRowNum = m_lastPhysicalRow + 1;
        m_lastPhysicalRow = m_pendingRowNum;
    }

    void XLStreamReader::ensureIndexedSst() const
    {
        if (m_sstIndexReady) return;
        m_sstIndexReady = true;
        m_sstOffsets.clear();
        m_sstXml.clear();
        m_sstFileBacked = false;

        try {
            if (!m_archive.hasEntry("xl/sharedStrings.xml")) return;
            m_sstXml = m_archive.getEntry("xl/sharedStrings.xml");
            if (m_options.maxSstEntryBytes > 0 && m_sstXml.size() > m_options.maxSstEntryBytes) {
                const size_t sz = m_sstXml.size();
                m_sstXml.clear();
                throw XLInputError("XLStreamReader: sharedStrings.xml exceeds maxSstEntryBytes (" + std::to_string(sz) +
                                   " > " + std::to_string(m_options.maxSstEntryBytes) + ")");
            }
        }
        catch (const XLInputError&) {
            throw;
        }
        catch (...) {
            m_sstXml.clear();
            return;
        }

        auto indexSiOffsets = [this](std::string_view xml) {
            size_t pos = 0;
            while (pos < xml.size()) {
                size_t si = xml.find("<si", pos);
                if (si == std::string_view::npos) break;
                if (si + 3 < xml.size()) {
                    const char next = xml[si + 3];
                    if (next != '>' && next != ' ' && next != '/' && next != '\t' && next != '\r' && next != '\n') {
                        pos = si + 3;
                        continue;
                    }
                }
                m_sstOffsets.push_back(static_cast<uint32_t>(si));
                pos = si + 3;
            }
        };

        indexSiOffsets(m_sstXml);

        // Spill large SST to a temp file so peak RSS after indexing can drop the string payload
        // (offsets remain; lookup seeks the file). Aligns with excelize UnzipXMLSizeLimit behavior.
        if (m_options.sstSpillThreshold > 0 && m_sstXml.size() >= m_options.sstSpillThreshold) {
            try {
                const auto base = m_options.tempDir.empty() ? std::filesystem::temp_directory_path() : m_options.tempDir;
                std::mt19937 rng(std::random_device {}());
                m_sstTempPath =
                    base / (std::string("openxlsx_sst_") +
                            std::to_string(std::chrono::system_clock::now().time_since_epoch().count()) + "_" +
                            std::to_string(rng()) + ".xml");
                std::ofstream out(m_sstTempPath, std::ios::binary);
                if (out) {
                    out.write(m_sstXml.data(), static_cast<std::streamsize>(m_sstXml.size()));
                    if (out.good()) {
                        m_sstFileBacked = true;
                        m_sstXml.clear();
                        m_sstXml.shrink_to_fit();
                    }
                    else {
                        m_sstTempPath.clear();
                    }
                }
            }
            catch (...) {
                m_sstFileBacked = false;
                m_sstTempPath.clear();
                // Keep m_sstXml in memory as fallback.
            }
        }
    }

    std::string XLStreamReader::lookupIndexedSst(int32_t index) const
    {
        ensureIndexedSst();
        if (index < 0 || static_cast<size_t>(index) >= m_sstOffsets.size()) return {};
        const size_t start = m_sstOffsets[static_cast<size_t>(index)];

        std::string slice;
        if (m_sstFileBacked) {
            const size_t endOff =
                (static_cast<size_t>(index) + 1 < m_sstOffsets.size()) ? m_sstOffsets[static_cast<size_t>(index) + 1] : 0;
            std::ifstream in(m_sstTempPath, std::ios::binary);
            if (!in) return {};
            in.seekg(0, std::ios::end);
            const auto fileSize = static_cast<size_t>(in.tellg());
            const size_t end    = (endOff > 0) ? endOff : fileSize;
            if (start >= fileSize || end <= start) return {};
            const size_t len = end - start;
            slice.resize(len);
            in.seekg(static_cast<std::streamoff>(start));
            in.read(slice.data(), static_cast<std::streamsize>(len));
            if (!in && !in.eof()) return {};
            slice.resize(static_cast<size_t>(in.gcount()));
        }
        else {
            const size_t end =
                (static_cast<size_t>(index) + 1 < m_sstOffsets.size()) ? m_sstOffsets[static_cast<size_t>(index) + 1] : m_sstXml.size();
            if (start >= m_sstXml.size() || end <= start) return {};
            slice = m_sstXml.substr(start, end - start);
        }

        // Extract concatenated <t>...</t> text inside this <si>…</si> slice (plain + rich runs).
        std::string result;
        size_t      p = 0;
        while (p < slice.size()) {
            size_t tOpen = slice.find("<t", p);
            if (tOpen == std::string::npos) break;
            size_t tClose = slice.find('>', tOpen);
            if (tClose == std::string::npos) break;
            if (tClose > 0 && slice[tClose - 1] == '/') {
                p = tClose + 1;
                continue;
            }
            size_t tEnd = slice.find("</t>", tClose + 1);
            if (tEnd == std::string::npos) break;
            std::string piece = slice.substr(tClose + 1, tEnd - (tClose + 1));
            xmlUnescape(piece);
            result += piece;
            p = tEnd + 4;
        }
        return result;
    }

    XLCellValue XLStreamReader::parseCellValue(const std::string& cellType, std::string& cellValue) const
    {
        if (cellType == "s") {
            char* ep  = nullptr;
            auto  idx = static_cast<int32_t>(std::strtol(cellValue.data(), &ep, 10));
            if (ep == cellValue.data()) return XLCellValue{};
            // Out-of-range / corrupt SST indices must not abort the whole stream scan.
            try {
                if (m_options.sharedStringMode == XLStreamSharedStringMode::Indexed) {
                    auto s = lookupIndexedSst(idx);
                    if (s.empty() && idx >= 0) {
                        // Still accept empty string entries; OOB returns empty from lookup.
                    }
                    return XLCellValue(std::move(s));
                }
                // Prefer XLSharedStrings (concrete) over XLSharedStringTable port — more robust
                // when the worksheet was materialized from a stream-written package.
                try {
                    std::string_view sv = m_worksheet->sharedStrings().getStringView(idx);
                    return XLCellValue(std::string(sv));
                }
                catch (...) {
                    const char* p = m_worksheet->sharedStringTable().getString(idx);
                    if (p == nullptr) return XLCellValue{};
                    return XLCellValue(std::string(p));
                }
            }
            catch (...) {
                return XLCellValue{};
            }
        }
        if (cellType == "b") return XLCellValue(cellValue == "1" || cellValue == "true");
        if (cellType == "inlineStr" || cellType == "str") return XLCellValue(cellValue);
        if (cellType == "e") {
            XLCellValue ev;
            ev.setError(cellValue.c_str());
            return ev;
        }
        // Numeric (t="" or t="n") — also used when formula has cached numeric value
        if (!cellValue.empty()) {
            double val     = 0.0;
            auto [ptr, ec] = fast_float::from_chars(cellValue.data(), cellValue.data() + cellValue.size(), val);
            if (ec == std::errc()) {
                if (cellValue.find('.') == std::string::npos && cellValue.find('E') == std::string::npos &&
                    cellValue.find('e') == std::string::npos)
                    return XLCellValue(static_cast<int64_t>(val));
                return XLCellValue(val);
            }
            // Formula cells with t="str" already handled; leftover text:
            return XLCellValue(cellValue);
        }
        return XLCellValue{};
    }

    bool XLStreamReader::consumeNextLogicalRow()
    {
        m_lastError.clear();
        if (!m_hasPendingPhysical) {
            if (!loadNextPhysicalRow()) return false;
        }

        if (m_options.emptyRows == XLStreamEmptyRowPolicy::EmitEmptyRows && m_currentRow + 1 < m_pendingRowNum) {
            m_currentRow                     = m_currentRow + 1;
            m_currentRowOpts                 = {};
            m_currentRowOpts.isSyntheticEmpty = true;
            m_reuseDetailed.clear();
            m_reuseValues.clear();
            return true;
        }

        // Deliver pending physical row
        m_currentRow     = m_pendingRowNum;
        m_currentRowOpts = m_pendingRowOpts;
        m_reuseDetailed  = m_pendingCells;
        m_reuseValues.clear();
        m_reuseValues.reserve(m_reuseDetailed.size());
        for (const auto& c : m_reuseDetailed) m_reuseValues.push_back(c.value);

        m_hasPendingPhysical = false;
        m_pendingCells.clear();
        return true;
    }

    std::vector<XLCellValue> XLStreamReader::nextRow()
    {
        if (!consumeNextLogicalRow()) return {};
        return m_reuseValues;
    }

    std::vector<XLStreamCellView> XLStreamReader::nextRowDetailed()
    {
        if (!consumeNextLogicalRow()) return {};
        return m_reuseDetailed;
    }

    std::vector<XLStreamCellView> XLStreamReader::nextRowSparse()
    {
        if (!consumeNextLogicalRow()) return {};
        std::vector<XLStreamCellView> sparse;
        sparse.reserve(m_reuseDetailed.size());
        for (const auto& c : m_reuseDetailed) {
            if (c.value.type() != XLValueType::Empty || c.formula.has_value()) sparse.push_back(c);
        }
        return sparse;
    }

    bool XLStreamReader::nextRowInto(std::vector<XLCellValue>& out)
    {
        if (!consumeNextLogicalRow()) {
            out.clear();
            return false;
        }
        out = m_reuseValues;
        return true;
    }

    bool XLStreamReader::nextRowDetailedInto(std::vector<XLStreamCellView>& out)
    {
        if (!consumeNextLogicalRow()) {
            out.clear();
            return false;
        }
        out = m_reuseDetailed;
        return true;
    }

    void XLStreamReader::rebuildColumnFilterSet() const
    {
        m_columnFilterSet.clear();
        for (uint16_t c : m_options.columnFilter) {
            if (c > 0) m_columnFilterSet.insert(c);
        }
        m_columnFilterReady = true;
    }

    bool XLStreamReader::columnAllowed(uint16_t col) const
    {
        if (m_options.columnFilter.empty()) return true;
        if (!m_columnFilterReady) rebuildColumnFilterSet();
        return m_columnFilterSet.count(col) > 0;
    }

    std::vector<XLStreamCellView> XLStreamReader::nextRowProjected()
    {
        if (!consumeNextLogicalRow()) return {};
        std::vector<XLStreamCellView> out;
        out.reserve(m_options.columnFilter.empty() ? m_reuseDetailed.size() : m_options.columnFilter.size());
        for (const auto& c : m_reuseDetailed) {
            if (c.column == 0) continue;
            if (!columnAllowed(c.column)) continue;
            // When no filter: behave like sparse (skip pure empties). With filter: include empties in range.
            if (m_options.columnFilter.empty()) {
                if (c.value.type() != XLValueType::Empty || c.formula.has_value()) out.push_back(c);
            }
            else {
                out.push_back(c);
            }
        }
        return out;
    }

    bool XLStreamReader::forEachCellInNextRow(const std::function<void(const XLStreamCellView&)>& fn)
    {
        if (!consumeNextLogicalRow()) return false;
        for (const auto& c : m_reuseDetailed) {
            if (c.column == 0) continue;
            if (!columnAllowed(c.column)) continue;
            if (m_options.columnFilter.empty() && c.value.type() == XLValueType::Empty && !c.formula.has_value()) continue;
            fn(c);
        }
        return true;
    }

    XLStreamColumns::XLStreamColumns(std::vector<std::vector<XLCellValue>> columns) : m_columns(std::move(columns)) {}

    bool XLStreamColumns::next()
    {
        if (m_columns.empty()) return false;
        if (!m_started) {
            m_started = true;
            m_index   = 1;
            return true;
        }
        if (m_index >= m_columns.size()) return false;
        ++m_index;
        return m_index <= m_columns.size();
    }

    const std::vector<XLCellValue>& XLStreamColumns::values() const
    {
        static const std::vector<XLCellValue> kEmpty;
        if (m_index == 0 || m_index > m_columns.size()) return kEmpty;
        return m_columns[m_index - 1];
    }

    XLStreamColumns XLStreamReader::streamColumns()
    {
        std::vector<std::vector<XLCellValue>> cols;
        size_t                                cellCount = 0;
        uint32_t                              rowNum    = 0;

        while (consumeNextLogicalRow()) {
            ++rowNum;
            for (const auto& c : m_reuseDetailed) {
                if (c.column == 0) continue;
                if (!columnAllowed(c.column)) continue;
                if (c.column > cols.size()) cols.resize(c.column);
                auto& col = cols[c.column - 1];
                if (col.size() < rowNum) col.resize(rowNum);
                col[rowNum - 1] = c.value;
                ++cellCount;
                if (m_options.maxColumnarCells > 0 && cellCount > m_options.maxColumnarCells) {
                    throw XLInputError("XLStreamReader::streamColumns: exceeded maxColumnarCells (" +
                                       std::to_string(m_options.maxColumnarCells) + ")");
                }
            }
            // Ensure all existing columns have this row (empty padding)
            for (auto& col : cols) {
                if (col.size() < rowNum) col.resize(rowNum);
            }
        }
        return XLStreamColumns(std::move(cols));
    }

    std::string XLStreamReader::valueToRawString(const XLCellValue& value) const
    {
        switch (value.type()) {
            case XLValueType::Empty: return {};
            case XLValueType::Boolean: return value.get<bool>() ? "TRUE" : "FALSE";
            case XLValueType::Integer: return std::to_string(value.get<int64_t>());
            case XLValueType::Float: {
                XLCellValue copy = value;
                return copy.getString();
            }
            case XLValueType::String: return value.get<std::string>();
            case XLValueType::RichText: return value.get<XLRichText>().plainText();
            case XLValueType::Error: {
                XLCellValue copy = value;
                return copy.getString();
            }
            default: {
                XLCellValue copy = value;
                return copy.getString();
            }
        }
    }

    std::string XLStreamReader::builtinNumFmtCode(uint32_t numFmtId)
    {
        // ECMA-376 built-in formats (subset used most often)
        static const std::unordered_map<uint32_t, const char*> kBuiltins = {
            {0, "General"},
            {1, "0"},
            {2, "0.00"},
            {3, "#,##0"},
            {4, "#,##0.00"},
            {9, "0%"},
            {10, "0.00%"},
            {11, "0.00E+00"},
            {12, "# ?/?"},
            {13, R"(# ??/??)"},
            {14, "mm-dd-yy"},
            {15, "d-mmm-yy"},
            {16, "d-mmm"},
            {17, "mmm-yy"},
            {18, "h:mm AM/PM"},
            {19, "h:mm:ss AM/PM"},
            {20, "h:mm"},
            {21, "h:mm:ss"},
            {22, "m/d/yy h:mm"},
            {37, "#,##0 ;(#,##0)"},
            {38, "#,##0 ;[Red](#,##0)"},
            {39, "#,##0.00;(#,##0.00)"},
            {40, "#,##0.00;[Red](#,##0.00)"},
            {45, "mm:ss"},
            {46, "[h]:mm:ss"},
            {47, "mmss.0"},
            {48, "##0.0E+0"},
            {49, "@"},
        };
        auto it = kBuiltins.find(numFmtId);
        if (it != kBuiltins.end()) return it->second;
        return "General";
    }

    std::string XLStreamReader::resolveFormatCode(XLStyleIndex styleIndex) const
    {
        if (!m_worksheet) return "General";
        try {
            // styles() is non-const on XLDocument but this path only reads format tables.
            auto& styles = const_cast<XLDocument&>(m_worksheet->parentDoc()).styles();
            auto& cfs    = styles.cellFormats();
            if (styleIndex >= cfs.count()) return "General";
            const uint32_t numFmtId = cfs.cellFormatByIndex(styleIndex).numberFormatId();
            try {
                return styles.numberFormats().numberFormatById(numFmtId).formatCode();
            }
            catch (...) {
                return builtinNumFmtCode(numFmtId);
            }
        }
        catch (...) {
            return "General";
        }
    }

    std::string XLStreamReader::formattedValue(const XLStreamCellView& cell) const
    {
        if (!m_options.applyNumberFormats || !cell.styleIndex.has_value()) return valueToRawString(cell.value);
        if (cell.value.type() == XLValueType::Empty) return {};
        try {
            const std::string code = resolveFormatCode(*cell.styleIndex);
            if (code.empty() || code == "General" || code == "@") return valueToRawString(cell.value);
            XLNumberFormatter fmt(code);
            return fmt.format(cell.value);
        }
        catch (...) {
            return valueToRawString(cell.value);
        }
    }

    std::vector<std::string> XLStreamReader::nextRowStrings()
    {
        if (!consumeNextLogicalRow()) return {};
        m_reuseStrings.clear();
        m_reuseStrings.reserve(m_reuseDetailed.size());
        for (const auto& cell : m_reuseDetailed) {
            if (m_options.applyNumberFormats)
                m_reuseStrings.push_back(formattedValue(cell));
            else
                m_reuseStrings.push_back(valueToRawString(cell.value));
        }
        return m_reuseStrings;
    }

    uint32_t XLStreamReader::currentRow() const { return m_currentRow; }

    XLStreamRowOptsView XLStreamReader::currentRowOpts() const { return m_currentRowOpts; }

    void XLStreamReader::close() { cleanup(); }

}    // namespace OpenXLSX

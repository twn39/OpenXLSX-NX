// ===== External Includes ===== //
#include <cstring>
#include <fast_float/fast_float.h>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLSharedStrings.hpp"
#include "XLStreamReader.hpp"
#include "XLWorksheet.hpp"

namespace
{

    // Decode standard XML character entities in-place.
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

}    // anonymous namespace

namespace OpenXLSX
{

    XLStreamReader::XLStreamReader(const XLWorksheet* worksheet) : m_worksheet(worksheet)
    {
        if (!worksheet) throw XLInternalError("Worksheet is null");

        m_zipStream = worksheet->parentDoc().archive().openEntryStream(worksheet->getXmlPath());
        if (!m_zipStream) throw XLInternalError("Failed to open worksheet zip stream");

        // Pre-reserve scratch buffers to avoid reallocations per row
        m_tagNameBuf.reserve(64);
        m_attrNameBuf.reserve(32);
        m_attrValueBuf.reserve(128);
        m_textContentBuf.reserve(256);
    }

    XLStreamReader::~XLStreamReader() { cleanup(); }

    XLStreamReader::XLStreamReader(XLStreamReader&& other) noexcept
        : m_worksheet(other.m_worksheet),
          m_zipStream(other.m_zipStream),
          m_buffer(std::move(other.m_buffer)),
          m_eof(other.m_eof),
          m_currentRow(other.m_currentRow),
          m_tagNameBuf(std::move(other.m_tagNameBuf)),
          m_attrNameBuf(std::move(other.m_attrNameBuf)),
          m_attrValueBuf(std::move(other.m_attrValueBuf)),
          m_textContentBuf(std::move(other.m_textContentBuf))
    {
        other.m_worksheet = nullptr;
        other.m_zipStream = nullptr;
        other.m_eof       = true;
    }

    XLStreamReader& XLStreamReader::operator=(XLStreamReader&& other) noexcept
    {
        if (this != &other) {
            cleanup();
            m_worksheet      = other.m_worksheet;
            m_zipStream      = other.m_zipStream;
            m_buffer         = std::move(other.m_buffer);
            m_eof            = other.m_eof;
            m_currentRow     = other.m_currentRow;
            m_tagNameBuf     = std::move(other.m_tagNameBuf);
            m_attrNameBuf    = std::move(other.m_attrNameBuf);
            m_attrValueBuf   = std::move(other.m_attrValueBuf);
            m_textContentBuf = std::move(other.m_textContentBuf);

            other.m_worksheet = nullptr;
            other.m_zipStream = nullptr;
            other.m_eof       = true;
        }
        return *this;
    }

    void XLStreamReader::cleanup()
    {
        if (m_zipStream && m_worksheet) {
            m_worksheet->parentDoc().archive().closeEntryStream(m_zipStream);
            m_zipStream = nullptr;
        }
    }

    void XLStreamReader::fetchMoreData()
    {
        if (m_eof || !m_zipStream) return;

        constexpr size_t kReadBuf = 65536;
        char             buf[kReadBuf];    // 64 KB chunks
        auto             bytesRead = m_worksheet->parentDoc().archive().readEntryStream(m_zipStream, buf, kReadBuf);

        if (bytesRead > 0)
            m_buffer.append(buf, static_cast<size_t>(bytesRead));
        else {
            m_eof = true;
            cleanup();
        }
    }

    bool XLStreamReader::hasNext()
    {
        if (!m_buffer.empty() && m_buffer.find("<row") != std::string::npos) return true;
        while (!m_eof) {
            fetchMoreData();
            if (m_buffer.find("<row") != std::string::npos) return true;
        }
        return false;
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  nextRow() — SAX state machine.
    //
    //  Walks raw bytes of m_buffer, tracking which tag we are inside and
    //  collecting only the attributes / text content we need.
    //  No pugi document is created per-row. Reusable member string buffers are
    //  cleared (not destroyed) between calls, so their heap capacity persists.
    // ─────────────────────────────────────────────────────────────────────────
    std::vector<XLCellValue> XLStreamReader::nextRow()
    {
        std::vector<XLCellValue> result;

        // ── Phase 1: ensure m_buffer contains a complete <row>…</row> span ──
        while (true) {
            size_t rowStart = m_buffer.find("<row");
            if (rowStart == std::string::npos) {
                if (m_eof) return result;
                if (m_buffer.size() > 5) m_buffer.erase(0, m_buffer.size() - 5);
                fetchMoreData();
                continue;
            }

            size_t firstGt = m_buffer.find('>', rowStart);
            if (firstGt == std::string::npos) {
                if (m_eof) return result;
                fetchMoreData();
                continue;
            }

            // Self-closing <row … />: empty row
            bool selfClosingRow = (firstGt > 0) && (m_buffer[firstGt - 1] == '/');
            if (selfClosingRow) {
                size_t rPos = m_buffer.find("r=\"", rowStart);
                if (rPos != std::string::npos && rPos < firstGt) {
                    size_t rEnd = m_buffer.find('"', rPos + 3);
                    if (rEnd != std::string::npos && rEnd < firstGt)
                        m_currentRow = static_cast<uint32_t>(std::strtoul(m_buffer.c_str() + rPos + 3, nullptr, 10));
                }
                else {
                    ++m_currentRow;
                }
                m_buffer.erase(0, firstGt + 1);
                return result;
            }

            size_t endTag = m_buffer.find("</row>", rowStart);
            if (endTag != std::string::npos) break;    // full row in buffer

            if (m_eof) return result;
            fetchMoreData();
        }

        // ── Phase 2: note the slice boundaries (offsets in m_buffer) ─────────
        const size_t rowStart = m_buffer.find("<row");
        const size_t endTag   = m_buffer.find("</row>", rowStart);
        const size_t sliceEnd = endTag + 6;    // one-past "</row>"

        // ── Phase 3: SAX scan using raw pointers — no copy of the slice ──────
        const char* p   = m_buffer.data() + rowStart;
        const char* end = m_buffer.data() + sliceEnd;

        auto skipWS = [&](const char* cp) -> const char* {
            while (cp < end && (*cp == ' ' || *cp == '\t' || *cp == '\r' || *cp == '\n')) ++cp;
            return cp;
        };

        // Per-cell state (local — no heap alloc per row)
        std::string cellRef;
        std::string cellType;
        std::string cellValue;
        bool        inVTag      = false;
        bool        inIsTTag    = false;    // inside <is><t> (inline string)
        uint16_t    expectedCol = 1;

        // Commit a parsed cell into result, filling column gaps with empties
        auto flushCell = [&]() {
            if (cellRef.empty()) return;

            uint16_t actualCol = expectedCol;
            {
                uint16_t col = 0;
                for (char c : cellRef) {
                    if (c >= 'A' && c <= 'Z')
                        col = static_cast<uint16_t>(col * 26 + (c - 'A' + 1));
                    else
                        break;
                }
                if (col > 0) actualCol = col;
            }

            while (expectedCol < actualCol) {
                result.emplace_back();
                ++expectedCol;
            }

            xmlUnescape(cellValue);

            if (cellType == "s") {
                char* ep  = nullptr;
                auto  idx = static_cast<int32_t>(std::strtol(cellValue.data(), &ep, 10));
                if (ep != cellValue.data())
                    result.emplace_back(std::string(m_worksheet->parentDoc().sharedStrings().getString(idx)));
                else
                    result.emplace_back();
            }
            else if (cellType == "b") {
                result.emplace_back(cellValue == "1" || cellValue == "true");
            }
            else if (cellType == "inlineStr" || cellType == "str") {
                result.emplace_back(cellValue);
            }
            else if (cellType == "e") {
                XLCellValue ev;
                ev.setError(cellValue.c_str());
                result.emplace_back(ev);
            }
            else {
                // Numeric (t="" or t="n")
                if (!cellValue.empty()) {
                    double val     = 0.0;
                    auto [ptr, ec] = fast_float::from_chars(cellValue.data(), cellValue.data() + cellValue.size(), val);
                    if (ec == std::errc()) {
                        if (cellValue.find('.') == std::string::npos && cellValue.find('E') == std::string::npos &&
                            cellValue.find('e') == std::string::npos)
                            result.emplace_back(static_cast<int64_t>(val));
                        else
                            result.emplace_back(val);
                    }
                    else {
                        result.emplace_back();
                    }
                }
                else {
                    result.emplace_back();
                }
            }

            ++expectedCol;
            cellRef.clear();
            cellType.clear();
            cellValue.clear();
            inVTag   = false;
            inIsTTag = false;
        };

        // ── Tag-by-tag scan ───────────────────────────────────────────────────
        while (p < end) {
            if (*p != '<') {
                if (inVTag || inIsTTag) cellValue += *p;
                ++p;
                continue;
            }

            ++p;
            if (p >= end) break;

            const bool isClose = (*p == '/');
            if (isClose) ++p;

            m_tagNameBuf.clear();
            while (p < end && *p != ' ' && *p != '>' && *p != '/') m_tagNameBuf += *p++;

            if (isClose) {
                if (m_tagNameBuf == "row" || m_tagNameBuf == "c")
                    flushCell();
                else if (m_tagNameBuf == "v")
                    inVTag = false;
                else if (m_tagNameBuf == "t")
                    inIsTTag = false;
                while (p < end && *p != '>') ++p;
                if (p < end) ++p;
                continue;
            }

            // Opening tag — parse only 'r' and 't' attributes
            std::string localR;
            std::string localT;
            bool        selfClose = false;

            while (p < end) {
                p = skipWS(p);
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
                p = skipWS(p);
                if (p >= end || *p != '=') continue;
                ++p;
                p = skipWS(p);

                m_attrValueBuf.clear();
                char q = 0;
                if (p < end && (*p == '"' || *p == '\'')) q = *p++;
                while (p < end && (q ? *p != q : (*p != ' ' && *p != '>'))) m_attrValueBuf += *p++;
                if (q && p < end) ++p;

                if (m_attrNameBuf == "r")
                    localR = m_attrValueBuf;
                else if (m_attrNameBuf == "t")
                    localT = m_attrValueBuf;
            }

            if (m_tagNameBuf == "row") {
                if (!localR.empty())
                    m_currentRow = static_cast<uint32_t>(std::strtoul(localR.c_str(), nullptr, 10));
                else
                    ++m_currentRow;
            }
            else if (m_tagNameBuf == "c") {
                flushCell();
                cellRef  = localR;
                cellType = localT;
                if (selfClose) flushCell();    // <c r="…"/> — empty cell
            }
            else if (m_tagNameBuf == "v") {
                inVTag   = true;
                inIsTTag = false;
            }
            else if (m_tagNameBuf == "t") {
                inIsTTag = true;
                inVTag   = false;
            }
            // All other tags (worksheet, sheetData, f, etc.) are intentionally ignored
        }

        // ── Phase 4: advance m_buffer past the consumed row ───────────────────
        m_buffer.erase(0, sliceEnd);

        return result;
    }

    uint32_t XLStreamReader::currentRow() const { return m_currentRow; }

    void XLStreamReader::close() { cleanup(); }

}    // namespace OpenXLSX

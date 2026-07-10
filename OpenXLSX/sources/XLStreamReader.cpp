// ===== External Includes ===== //
#include <cctype>
#include <cstring>
#include <fast_float/fast_float.h>
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
    }

    XLStreamReader& XLStreamReader::operator=(XLStreamReader&& other) noexcept
    {
        if (this != &other) {
            cleanup();
            m_worksheet          = other.m_worksheet;
            m_archive            = std::move(other.m_archive);
            m_zipStream          = other.m_zipStream;
            m_options            = std::move(other.m_options);
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
        }
        return *this;
    }

    void XLStreamReader::cleanup()
    {
        if (!m_zipStream) return;
        // Prefer archive handle (independent of worksheet lifetime). Skip if package already closed.
        try {
            if (m_archive.isOpen()) m_archive.closeEntryStream(m_zipStream);
        }
        catch (...) {
            // Best-effort; document may already be torn down.
        }
        m_zipStream = nullptr;
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
        bool        inVTag   = false;
        bool        inIsTTag = false;
        bool        inFTag   = false;
        uint16_t    expectedCol = 1;

        auto flushCell = [&]() {
            if (cellRef.empty() && cellValue.empty() && cellFormula.empty() && cellStyle.empty() && cellType.empty()) return;

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
            if (!cellFormula.empty()) view.formula = cellFormula;
            if (!cellStyle.empty()) {
                char* ep = nullptr;
                auto  s  = static_cast<XLStyleIndex>(std::strtoul(cellStyle.c_str(), &ep, 10));
                if (ep != cellStyle.c_str()) view.styleIndex = s;
            }
            m_pendingCells.push_back(std::move(view));
            ++expectedCol;

            cellRef.clear();
            cellType.clear();
            cellStyle.clear();
            cellValue.clear();
            cellFormula.clear();
            inVTag   = false;
            inIsTTag = false;
            inFTag   = false;
        };

        while (p < end) {
            if (*p != '<') {
                if (inVTag || inIsTTag) cellValue += *p;
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
                while (p < end && *p != '>') ++p;
                if (p < end) ++p;
                continue;
            }

            std::string localR, localT, localS, localHt, localHidden, localOutline;
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
            }

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
                // Only treat <t> as inline string text when not inside formula
                inIsTTag = true;
                inVTag   = false;
                inFTag   = false;
            }
            else if (m_tagNameBuf == "f") {
                inFTag   = true;
                inVTag   = false;
                inIsTTag = false;
                if (selfClose) inFTag = false;
            }
        }

        flushCell();
        if (m_pendingRowNum == 0) m_pendingRowNum = m_lastPhysicalRow + 1;
        m_lastPhysicalRow = m_pendingRowNum;
    }

    XLCellValue XLStreamReader::parseCellValue(const std::string& cellType, std::string& cellValue) const
    {
        if (cellType == "s") {
            char* ep  = nullptr;
            auto  idx = static_cast<int32_t>(std::strtol(cellValue.data(), &ep, 10));
            if (ep == cellValue.data()) return XLCellValue{};
            // Out-of-range / corrupt SST indices must not abort the whole stream scan.
            try {
                return XLCellValue(std::string(m_worksheet->sharedStringTable().getString(idx)));
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

// ===== External Includes ===== //
#include <charconv>
#include <chrono>
#include <cstring>
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
#include "XLStyle.hpp"
#include "XLUtilities.hpp"
#include "XLWorksheet.hpp"
#include "XLXmlData.hpp"

namespace OpenXLSX
{
    namespace
    {
        void validateRowOpts(const XLStreamRowOpts* opts)
        {
            if (!opts) return;
            if (opts->height.has_value() && (*opts->height < 0.0 || *opts->height > 409.0)) {
                throw XLInputError("XLStreamWriter: row height must be in range [0, 409]");
            }
            if (opts->outlineLevel.has_value() && *opts->outlineLevel > 7) {
                throw XLInputError("XLStreamWriter: outlineLevel must be in range [0, 7]");
            }
        }

        std::filesystem::path makeTempPath(const std::filesystem::path& dir)
        {
            const auto base = dir.empty() ? std::filesystem::temp_directory_path() : dir;
            std::mt19937 rng(std::random_device {}());
            return base / (std::string("openxlsx_stream_") +
                           std::to_string(std::chrono::system_clock::now().time_since_epoch().count()) + "_" +
                           std::to_string(rng()) + ".xml");
        }

        // Fixed-width dimension slot (max Excel used range). Close patches via seek/overwrite
        // without reloading the entire sheet XML into memory (P0a).
        constexpr std::string_view kDimensionSlot = R"(<dimension ref="A1:XFD1048576"/>)";

        constexpr const char* kFreshWorksheetPrefix =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
            " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
            " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""
            " mc:Ignorable=\"x14ac xr xr2 xr3\""
            " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\""
            " xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\""
            " xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\""
            " xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\">"
            "<dimension ref=\"A1:XFD1048576\"/>"
            "<sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"16\"/>";

        constexpr const char* kFreshBottom = "</sheetData>"
                                             "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" "
                                             "header=\"0.3\" footer=\"0.3\"/>"
                                             "</worksheet>";

        /**
         * @brief Build a fixed-width dimension element matching kDimensionSlot.size().
         * @details Pads with spaces after "/>" so the slot length is constant for in-place patch.
         */
        std::string makeDimensionSlot(std::string_view usedRef)
        {
            std::string tag;
            tag.reserve(kDimensionSlot.size());
            tag += "<dimension ref=\"";
            tag += usedRef;
            tag += "\"/>";
            if (tag.size() > kDimensionSlot.size()) {
                // Should never happen for legal A1:XFD1048576; fall back to max range.
                return std::string(kDimensionSlot);
            }
            tag.append(kDimensionSlot.size() - tag.size(), ' ');
            return tag;
        }

        /**
         * @brief Replace any existing &lt;dimension...&gt; in @p xml with the fixed slot; insert if missing.
         * @return Byte offset of the slot within @p xml, or npos on failure.
         */
        size_t replaceWithDimensionSlot(std::string& xml)
        {
            const std::string slot(kDimensionSlot);
            const size_t      dimPos = xml.find("<dimension");
            if (dimPos == std::string::npos) {
                const size_t sheetDataPos = xml.find("<sheetData");
                if (sheetDataPos != std::string::npos) {
                    xml.insert(sheetDataPos, slot);
                    return sheetDataPos;
                }
                // Append before closing worksheet if present.
                const size_t endWs = xml.find("</worksheet>");
                if (endWs != std::string::npos) {
                    xml.insert(endWs, slot);
                    return endWs;
                }
                xml += slot;
                return xml.size() - slot.size();
            }

            size_t end = xml.find("/>", dimPos);
            if (end != std::string::npos && end < xml.find('>', dimPos)) {
                end += 2;
            }
            else {
                end = xml.find('>', dimPos);
                if (end == std::string::npos) return std::string::npos;
                const size_t close = xml.find("</dimension>", end);
                if (close != std::string::npos) end = close + 12;
                else
                    ++end;
            }
            xml.replace(dimPos, end - dimPos, slot);
            return dimPos;
        }
    }    // namespace

    // ── Construction ─────────────────────────────────────────────────────────

    XLStreamWriter::XLStreamWriter(XLWorksheet* worksheet, XLStreamWriteOptions options)
        : m_options(std::move(options)),
          m_currentRow(worksheet ? worksheet->rowCount() + 1 : 1),
          m_baseMaxCol(worksheet ? worksheet->columnCount() : 0),
          m_baseLastRow(worksheet ? worksheet->rowCount() : 0),
          m_active(true),
          m_worksheet(worksheet)
    {
        m_writeBuffer.reserve(kRowFlushThreshold + 64UL * 1024UL);
        m_memBuffer.reserve(std::min(m_options.memorySpillThreshold, size_t {256 * 1024}));

        if (m_options.memorySpillThreshold == 0) {
            m_tempPath = makeTempPath(m_options.tempDir);
            m_fileStream.open(m_tempPath, std::ios::binary);
            if (!m_fileStream.is_open())
                throw XLInternalError("Failed to open temporary stream file: " + m_tempPath.string());
            m_usingFile = true;
        }

        markWriterOpen(true);
    }

    XLStreamWriter::XLStreamWriter(XLWorksheet* worksheet, bool useSharedStrings, size_t maxUniqueStrings)
        : XLStreamWriter(worksheet,
                         XLStreamWriteOptions {XLStreamOpenMode::Auto,
                                               4 * 1024 * 1024,
                                               {},
                                               useSharedStrings,
                                               maxUniqueStrings,
                                               64 * 1024 * 1024,
                                               XLStreamDimensionMode::FixedSlot,
                                               true})
    {}

    XLStreamWriter::~XLStreamWriter()
    {
        try {
            if (m_active) flushSheetDataClose();
        }
        catch (...) {
            m_active = false;
            markWriterOpen(false);
        }
    }

    XLStreamWriter::XLStreamWriter(XLStreamWriter&& other) noexcept
        : m_options(std::move(other.m_options)),
          m_memBuffer(std::move(other.m_memBuffer)),
          m_tempPath(std::move(other.m_tempPath)),
          m_fileStream(std::move(other.m_fileStream)),
          m_usingFile(other.m_usingFile),
          m_sinkBytesWritten(other.m_sinkBytesWritten),
          m_dimensionByteOffset(other.m_dimensionByteOffset),
          m_dimensionSlotSize(other.m_dimensionSlotSize),
          m_preContent(std::move(other.m_preContent)),
          m_preHeaderPath(std::move(other.m_preHeaderPath)),
          m_preHeaderFileBacked(other.m_preHeaderFileBacked),
          m_bottomHalf(std::move(other.m_bottomHalf)),
          m_headerWritten(other.m_headerWritten),
          m_sheetDataStarted(other.m_sheetDataStarted),
          m_preHasOpenSheetData(other.m_preHasOpenSheetData),
          m_colDirectives(std::move(other.m_colDirectives)),
          m_panesXml(std::move(other.m_panesXml)),
          m_mergeRefs(std::move(other.m_mergeRefs)),
          m_rowBreaks(std::move(other.m_rowBreaks)),
          m_colBreaks(std::move(other.m_colBreaks)),
          m_table(std::move(other.m_table)),
          m_autoFilterRange(std::move(other.m_autoFilterRange)),
          m_hyperlinks(std::move(other.m_hyperlinks)),
          m_currentRow(other.m_currentRow),
          m_lastWrittenRow(other.m_lastWrittenRow),
          m_maxColumn(other.m_maxColumn),
          m_baseMaxCol(other.m_baseMaxCol),
          m_baseLastRow(other.m_baseLastRow),
          m_active(other.m_active),
          m_writeBuffer(std::move(other.m_writeBuffer)),
          m_worksheet(other.m_worksheet),
          m_wroteFormula(other.m_wroteFormula),
          m_sstBytesUsed(other.m_sstBytesUsed),
          m_cachedDateStyle(other.m_cachedDateStyle),
          m_dateStyleReady(other.m_dateStyleReady),
          m_localCache(std::move(other.m_localCache))
    {
        other.m_active              = false;
        other.m_worksheet           = nullptr;
        other.m_usingFile           = false;
        other.m_sinkBytesWritten    = 0;
        other.m_dimensionSlotSize   = 0;
        other.m_sstBytesUsed        = 0;
        other.m_preHeaderFileBacked = false;
        other.m_dateStyleReady      = false;
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
            m_options             = std::move(other.m_options);
            m_memBuffer           = std::move(other.m_memBuffer);
            m_tempPath            = std::move(other.m_tempPath);
            m_fileStream          = std::move(other.m_fileStream);
            m_usingFile           = other.m_usingFile;
            m_sinkBytesWritten    = other.m_sinkBytesWritten;
            m_dimensionByteOffset = other.m_dimensionByteOffset;
            m_dimensionSlotSize   = other.m_dimensionSlotSize;
            m_preContent          = std::move(other.m_preContent);
            m_preHeaderPath       = std::move(other.m_preHeaderPath);
            m_preHeaderFileBacked = other.m_preHeaderFileBacked;
            m_bottomHalf          = std::move(other.m_bottomHalf);
            m_headerWritten       = other.m_headerWritten;
            m_sheetDataStarted    = other.m_sheetDataStarted;
            m_preHasOpenSheetData = other.m_preHasOpenSheetData;
            m_colDirectives       = std::move(other.m_colDirectives);
            m_panesXml            = std::move(other.m_panesXml);
            m_mergeRefs           = std::move(other.m_mergeRefs);
            m_rowBreaks           = std::move(other.m_rowBreaks);
            m_colBreaks           = std::move(other.m_colBreaks);
            m_table               = std::move(other.m_table);
            m_autoFilterRange     = std::move(other.m_autoFilterRange);
            m_hyperlinks          = std::move(other.m_hyperlinks);
            m_currentRow          = other.m_currentRow;
            m_lastWrittenRow      = other.m_lastWrittenRow;
            m_maxColumn           = other.m_maxColumn;
            m_baseMaxCol          = other.m_baseMaxCol;
            m_baseLastRow         = other.m_baseLastRow;
            m_active              = other.m_active;
            m_writeBuffer         = std::move(other.m_writeBuffer);
            m_worksheet           = other.m_worksheet;
            m_wroteFormula        = other.m_wroteFormula;
            m_sstBytesUsed        = other.m_sstBytesUsed;
            m_cachedDateStyle     = other.m_cachedDateStyle;
            m_dateStyleReady      = other.m_dateStyleReady;
            m_localCache          = std::move(other.m_localCache);
            other.m_active              = false;
            other.m_worksheet           = nullptr;
            other.m_usingFile           = false;
            other.m_sinkBytesWritten    = 0;
            other.m_dimensionSlotSize   = 0;
            other.m_sstBytesUsed        = 0;
            other.m_preHeaderFileBacked = false;
            other.m_dateStyleReady      = false;
        }
        return *this;
    }

    bool XLStreamWriter::isStreamActive() const
    {
        return m_active && (m_usingFile ? m_fileStream.is_open() : true);
    }

    std::string XLStreamWriter::getTempFilePath() const { return m_tempPath.string(); }

    void XLStreamWriter::markWriterOpen(bool open) noexcept
    {
        if (m_worksheet) m_worksheet->setStreamWriterOpen(open);
    }

    void XLStreamWriter::checkStreamWritable() const
    {
        if (!isStreamActive()) throw XLInternalError("Stream writer is not active");
    }

    void XLStreamWriter::checkPhase0() const
    {
        checkStreamWritable();
        if (m_sheetDataStarted) {
            throw XLInputError(
                "XLStreamWriter: column/pane settings must be applied before the first row "
                "(excelize SetCol* / SetPanes order)");
        }
    }

    // ── Sink ─────────────────────────────────────────────────────────────────

    void XLStreamWriter::sinkWrite(std::string_view data)
    {
        if (data.empty()) return;
        if (m_usingFile) {
            m_fileStream.write(data.data(), static_cast<std::streamsize>(data.size()));
            if (!m_fileStream.good()) {
                throw XLInternalError("XLStreamWriter: failed to write to temporary stream file: " + m_tempPath.string());
            }
        }
        else {
            m_memBuffer.append(data.data(), data.size());
            if (m_options.memorySpillThreshold > 0 && m_memBuffer.size() >= m_options.memorySpillThreshold) {
                sinkSpillToFile();
            }
        }
        m_sinkBytesWritten += static_cast<uint64_t>(data.size());
    }

    void XLStreamWriter::sinkSpillToFile()
    {
        if (m_usingFile) return;
        m_tempPath = makeTempPath(m_options.tempDir);
        m_fileStream.open(m_tempPath, std::ios::binary);
        if (!m_fileStream.is_open())
            throw XLInternalError("XLStreamWriter: failed to spill stream buffer to: " + m_tempPath.string());
        if (!m_memBuffer.empty()) {
            m_fileStream.write(m_memBuffer.data(), static_cast<std::streamsize>(m_memBuffer.size()));
            if (!m_fileStream.good())
                throw XLInternalError("XLStreamWriter: failed while spilling stream buffer");
            m_memBuffer.clear();
            m_memBuffer.shrink_to_fit();
        }
        m_usingFile = true;
    }

    void XLStreamWriter::flushWriteBuffer()
    {
        if (m_writeBuffer.empty()) return;
        sinkWrite(m_writeBuffer);
        m_writeBuffer.clear();
    }

    // ── Phase 0 ──────────────────────────────────────────────────────────────

    void XLStreamWriter::normalizeColRange(uint16_t& minCol, uint16_t& maxCol) const
    {
        if (minCol < 1 || minCol > MAX_COLS || maxCol < 1 || maxCol > MAX_COLS) {
            throw XLInputError("XLStreamWriter: column number out of range");
        }
        if (minCol > maxCol) std::swap(minCol, maxCol);
    }

    void XLStreamWriter::setColWidth(uint16_t minCol, uint16_t maxCol, double width)
    {
        checkPhase0();
        normalizeColRange(minCol, maxCol);
        if (width < 0.0 || width > kMaxColWidth)
            throw XLInputError("XLStreamWriter: column width must be in range [0, 255]");
        ColDirective d;
        d.minCol = minCol;
        d.maxCol = maxCol;
        d.width  = width;
        m_colDirectives.push_back(d);
    }

    void XLStreamWriter::setColVisible(uint16_t minCol, uint16_t maxCol, bool visible)
    {
        checkPhase0();
        normalizeColRange(minCol, maxCol);
        ColDirective d;
        d.minCol = minCol;
        d.maxCol = maxCol;
        d.hidden = !visible;
        m_colDirectives.push_back(d);
    }

    void XLStreamWriter::setColOutlineLevel(uint16_t col, uint8_t level)
    {
        checkPhase0();
        if (col < 1 || col > MAX_COLS) throw XLInputError("XLStreamWriter: column number out of range");
        if (level < 1 || level > 7) throw XLInputError("XLStreamWriter: outlineLevel must be in range [1, 7]");
        ColDirective d;
        d.minCol = d.maxCol = col;
        d.outlineLevel      = level;
        m_colDirectives.push_back(d);
    }

    void XLStreamWriter::setColStyle(uint16_t minCol, uint16_t maxCol, XLStyleIndex style)
    {
        checkPhase0();
        normalizeColRange(minCol, maxCol);
        ColDirective d;
        d.minCol     = minCol;
        d.maxCol     = maxCol;
        d.styleIndex = style;
        m_colDirectives.push_back(d);
    }

    void XLStreamWriter::setPanesXml(std::string panesXml)
    {
        checkPhase0();
        m_panesXml = std::move(panesXml);
    }

    std::string XLStreamWriter::buildPanesXml(const XLStreamPanesOptions& panes) const
    {
        std::string xml = "<sheetViews><sheetView workbookViewId=\"0\">";

        if (panes.split) {
            if (panes.xSplit == 0.0 && panes.ySplit == 0.0) {
                xml += "</sheetView></sheetViews>";
                return xml;
            }
            xml += "<pane";
            if (panes.xSplit > 0.0) {
                xml += " xSplit=\"";
                fmt::format_to(std::back_inserter(xml), "{}", panes.xSplit);
                xml += '"';
            }
            if (panes.ySplit > 0.0) {
                xml += " ySplit=\"";
                fmt::format_to(std::back_inserter(xml), "{}", panes.ySplit);
                xml += '"';
            }
            if (!panes.topLeftCell.empty()) {
                xml += " topLeftCell=\"";
                xml += panes.topLeftCell;
                xml += '"';
            }
            xml += " activePane=\"";
            xml += panes.activePane.empty() ? "bottomRight" : panes.activePane;
            xml += "\" state=\"split\"/>";
            xml += "<selection pane=\"";
            xml += panes.activePane.empty() ? "bottomRight" : panes.activePane;
            xml += "\"/>";
        }
        else {
            const uint16_t cols = panes.freezeCols;
            const uint32_t rows = panes.freezeRows;
            if (cols == 0 && rows == 0) {
                xml += "</sheetView></sheetViews>";
                return xml;
            }
            std::string topLeft = panes.topLeftCell;
            if (topLeft.empty()) {
                topLeft = XLCellReference(rows + 1, static_cast<uint16_t>(cols + 1)).address();
            }
            const char* active = (cols > 0 && rows > 0) ? "bottomRight" : (cols > 0 ? "topRight" : "bottomLeft");
            xml += "<pane";
            if (cols > 0) {
                xml += " xSplit=\"";
                xml += std::to_string(cols);
                xml += '"';
            }
            if (rows > 0) {
                xml += " ySplit=\"";
                xml += std::to_string(rows);
                xml += '"';
            }
            xml += " topLeftCell=\"";
            xml += topLeft;
            xml += "\" activePane=\"";
            xml += active;
            xml += "\" state=\"frozen\"/>";
            if (cols > 0 && rows > 0) {
                xml += "<selection pane=\"topRight\"/>";
                xml += "<selection pane=\"bottomLeft\"/>";
                xml += "<selection pane=\"bottomRight\" activeCell=\"";
                xml += topLeft;
                xml += "\" sqref=\"";
                xml += topLeft;
                xml += "\"/>";
            }
            else {
                xml += "<selection pane=\"";
                xml += active;
                xml += "\" activeCell=\"";
                xml += topLeft;
                xml += "\" sqref=\"";
                xml += topLeft;
                xml += "\"/>";
            }
        }
        xml += "</sheetView></sheetViews>";
        return xml;
    }

    void XLStreamWriter::setPanes(const XLStreamPanesOptions& panes)
    {
        checkPhase0();
        m_panesXml = buildPanesXml(panes);
    }

    void XLStreamWriter::freezePanes(const std::string& topLeftCell)
    {
        XLCellReference ref(topLeftCell);
        freezePanes(static_cast<uint16_t>(ref.column() - 1), ref.row() - 1);
    }

    void XLStreamWriter::freezePanes(uint16_t cols, uint32_t rows)
    {
        XLStreamPanesOptions opts;
        opts.freezeCols  = cols;
        opts.freezeRows  = rows;
        opts.split       = false;
        setPanes(opts);
    }

    XLStyleIndex XLStreamWriter::ensureDefaultDateStyle()
    {
        if (m_dateStyleReady) return m_cachedDateStyle;
        m_dateStyleReady = true;
        m_cachedDateStyle = XLInvalidStyleIndex;
        if (!m_worksheet) return m_cachedDateStyle;
        try {
            XLStyle style;
            style.numberFormat = m_options.defaultDateTimeFormat.empty() ? "yyyy-mm-dd hh:mm:ss" : m_options.defaultDateTimeFormat;
            m_cachedDateStyle  = m_worksheet->parentDoc().styles().findOrCreateStyle(style);
        }
        catch (...) {
            m_cachedDateStyle = XLInvalidStyleIndex;
        }
        return m_cachedDateStyle;
    }

    void XLStreamWriter::insertPageBreak(std::string_view cellRef)
    {
        checkStreamWritable();
        const XLCellReference ref {std::string(cellRef)};
        m_rowBreaks.push_back(std::to_string(ref.row()));
        m_colBreaks.push_back(std::to_string(ref.column()));
    }

    std::string XLStreamWriter::buildColsXml() const
    {
        if (m_colDirectives.empty()) return {};
        std::string xml = "<cols>";
        for (const auto& d : m_colDirectives) {
            xml += "<col min=\"";
            xml += std::to_string(d.minCol);
            xml += "\" max=\"";
            xml += std::to_string(d.maxCol);
            xml += '"';
            if (d.width) {
                xml += " width=\"";
                fmt::format_to(std::back_inserter(xml), "{}", *d.width);
                xml += "\" customWidth=\"1\"";
            }
            if (d.styleIndex && *d.styleIndex != XLDefaultCellFormat && *d.styleIndex != XLInvalidStyleIndex) {
                xml += " style=\"";
                xml += std::to_string(*d.styleIndex);
                xml += '"';
            }
            if (d.hidden && *d.hidden) xml += " hidden=\"1\"";
            if (d.outlineLevel && *d.outlineLevel > 0) {
                xml += " outlineLevel=\"";
                xml += std::to_string(*d.outlineLevel);
                xml += '"';
            }
            xml += "/>";
        }
        xml += "</cols>";
        return xml;
    }

    void XLStreamWriter::injectBeforeSheetData(std::string fragment)
    {
        if (fragment.empty()) return;
        // Materialize file-backed hybrid header once if injects are needed.
        if (m_preHeaderFileBacked && m_preContent.empty() && !m_preHeaderPath.empty()) {
            std::ifstream in(m_preHeaderPath, std::ios::binary);
            if (!in) throw XLInternalError("XLStreamWriter: failed to read hybrid header temp file");
            m_preContent.assign(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
            m_preHeaderFileBacked = false;
            std::error_code ec;
            std::filesystem::remove(m_preHeaderPath, ec);
            m_preHeaderPath.clear();
        }
        const size_t pos = m_preContent.find("<sheetData");
        if (pos != std::string::npos) {
            m_preContent.insert(pos, fragment);
        }
        else {
            m_preContent += fragment;
        }
    }

    void XLStreamWriter::streamPreHeaderToSink()
    {
        if (m_preHeaderFileBacked && !m_preHeaderPath.empty() && m_preContent.empty()) {
            // Stream large hybrid header without holding it as std::string when no injects ran.
            std::ifstream in(m_preHeaderPath, std::ios::binary);
            if (!in) throw XLInternalError("XLStreamWriter: failed to stream hybrid header temp file");
            char   buf[65536];
            while (in) {
                in.read(buf, sizeof(buf));
                const auto n = in.gcount();
                if (n > 0) sinkWrite(std::string_view(buf, static_cast<size_t>(n)));
            }
            std::error_code ec;
            std::filesystem::remove(m_preHeaderPath, ec);
            m_preHeaderPath.clear();
            m_preHeaderFileBacked = false;
            return;
        }
        sinkWrite(m_preContent);
        m_preContent.clear();
        m_preContent.shrink_to_fit();
    }

    void XLStreamWriter::installDimensionSlot()
    {
        const size_t off = replaceWithDimensionSlot(m_preContent);
        if (off == std::string::npos) {
            m_dimensionSlotSize   = 0;
            m_dimensionByteOffset = 0;
            return;
        }
        m_dimensionSlotSize   = XLStreamWriter::kDimensionSlot.size();
        m_dimensionByteOffset = m_sinkBytesWritten + static_cast<uint64_t>(off);
    }

    void XLStreamWriter::ensureSheetDataStarted()
    {
        if (m_headerWritten) {
            m_sheetDataStarted = true;
            return;
        }

        // Inject Phase 0 fragments into pre-content.
        if (!m_panesXml.empty()) {
            std::string panes = m_panesXml;
            if (panes.find("<sheetViews") == std::string::npos) {
                panes = "<sheetViews>" + panes + "</sheetViews>";
            }
            // Prefer replacing existing sheetViews if present.
            const size_t sv = m_preContent.find("<sheetViews");
            if (sv != std::string::npos) {
                const size_t end = m_preContent.find("</sheetViews>", sv);
                if (end != std::string::npos) {
                    m_preContent.replace(sv, end + 13 - sv, panes);
                }
                else {
                    injectBeforeSheetData(panes);
                }
            }
            else {
                injectBeforeSheetData(panes);
            }
        }

        const std::string colsXml = buildColsXml();
        if (!colsXml.empty()) {
            // Replace existing empty/missing cols; if <cols> already present from hybrid DOM, append our cols before sheetData.
            injectBeforeSheetData(colsXml);
        }

        if (!m_preHasOpenSheetData) {
            // Fresh path: pre has no sheetData yet.
            if (m_preContent.find("<sheetData") == std::string::npos) { m_preContent += "<sheetData>"; }
        }

        // Install fixed-width dimension before flushing header so close can O(1) patch.
        // For file-backed hybrid headers, load once if still on disk (install needs a string).
        if (m_preHeaderFileBacked && m_preContent.empty() && !m_preHeaderPath.empty()) {
            std::ifstream in(m_preHeaderPath, std::ios::binary);
            if (!in) throw XLInternalError("XLStreamWriter: failed to read hybrid header for dimension slot");
            m_preContent.assign(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
            std::error_code ec;
            std::filesystem::remove(m_preHeaderPath, ec);
            m_preHeaderPath.clear();
            m_preHeaderFileBacked = false;
        }
        installDimensionSlot();

        streamPreHeaderToSink();
        m_headerWritten    = true;
        m_sheetDataStarted = true;
    }

    // ── Phase 2 helpers ──────────────────────────────────────────────────────

    void XLStreamWriter::mergeCell(std::string_view topLeft, std::string_view bottomRight)
    {
        checkStreamWritable();
        const XLCellReference a {std::string(topLeft)};
        const XLCellReference b {std::string(bottomRight)};
        const uint32_t        r1  = std::min(a.row(), b.row());
        const uint32_t        r2  = std::max(a.row(), b.row());
        const uint16_t        c1  = std::min(a.column(), b.column());
        const uint16_t        c2  = std::max(a.column(), b.column());
        const std::string     ref = XLCellReference(r1, c1).address() + ":" + XLCellReference(r2, c2).address();
        m_mergeRefs.push_back(ref);
        if (c2 > m_maxColumn) m_maxColumn = c2;
        (void)r2;
    }

    void XLStreamWriter::addTable(const XLStreamTableOptions& table)
    {
        checkStreamWritable();
        if (m_table.has_value()) throw XLInputError("XLStreamWriter: only one table is allowed per stream writer");
        if (table.range.empty()) throw XLInputError("XLStreamWriter::addTable: range is required");
        m_table = table;
    }

    void XLStreamWriter::setAutoFilter(std::string_view range)
    {
        checkStreamWritable();
        if (range.empty()) throw XLInputError("XLStreamWriter::setAutoFilter: range is required");
        m_autoFilterRange = std::string(range);
    }

    void XLStreamWriter::addHyperlink(std::string_view cellRef, std::string_view url, std::string_view tooltip)
    {
        checkStreamWritable();
        if (cellRef.empty() || url.empty()) throw XLInputError("XLStreamWriter::addHyperlink: cellRef and url required");
        XLStreamHyperlink h;
        h.cellRef   = std::string(cellRef);
        h.target    = std::string(url);
        h.external  = true;
        h.tooltip   = std::string(tooltip);
        m_hyperlinks.push_back(std::move(h));
    }

    void XLStreamWriter::addInternalHyperlink(std::string_view cellRef, std::string_view location, std::string_view tooltip)
    {
        checkStreamWritable();
        if (cellRef.empty() || location.empty())
            throw XLInputError("XLStreamWriter::addInternalHyperlink: cellRef and location required");
        XLStreamHyperlink h;
        h.cellRef  = std::string(cellRef);
        h.target   = std::string(location);
        h.external = false;
        h.display  = std::string(location);
        h.tooltip  = std::string(tooltip);
        m_hyperlinks.push_back(std::move(h));
    }

    std::string XLStreamWriter::buildAutoFilterXml() const
    {
        if (!m_autoFilterRange.has_value() || m_autoFilterRange->empty()) return {};
        std::string xml = "<autoFilter ref=\"";
        xml += *m_autoFilterRange;
        xml += "\"/>";
        return xml;
    }

    void XLStreamWriter::writeFormulaOpenTag(const XLStreamCell* cellMeta, const std::string& formulaText)
    {
        m_writeBuffer += "<f";
        if (cellMeta) {
            if (cellMeta->formulaType == XLStreamFormulaType::Array) m_writeBuffer += R"( t="array")";
            else if (cellMeta->formulaType == XLStreamFormulaType::Shared)
                m_writeBuffer += R"( t="shared")";
            else if (cellMeta->formulaType == XLStreamFormulaType::DataTable)
                m_writeBuffer += R"( t="dataTable")";
            if (cellMeta->formulaRef.has_value() && !cellMeta->formulaRef->empty()) {
                m_writeBuffer += R"( ref=")";
                appendEscaped(m_writeBuffer, *cellMeta->formulaRef);
                m_writeBuffer += '"';
            }
            if (cellMeta->formulaSi.has_value()) {
                m_writeBuffer += R"( si=")";
                m_writeBuffer += std::to_string(*cellMeta->formulaSi);
                m_writeBuffer += '"';
            }
            if (cellMeta->formulaCa) m_writeBuffer += R"( ca="1")";
        }
        if (formulaText.empty()) {
            m_writeBuffer += "/>";
        }
        else {
            m_writeBuffer += '>';
            appendEscaped(m_writeBuffer, formulaText);
            m_writeBuffer += "</f>";
        }
    }

    std::string XLStreamWriter::buildMergeCellsXml() const
    {
        if (m_mergeRefs.empty()) return {};
        std::string xml = "<mergeCells count=\"";
        xml += std::to_string(m_mergeRefs.size());
        xml += "\">";
        for (const auto& ref : m_mergeRefs) {
            xml += "<mergeCell ref=\"";
            xml += ref;
            xml += "\"/>";
        }
        xml += "</mergeCells>";
        return xml;
    }

    std::string XLStreamWriter::buildPageBreaksXml() const
    {
        std::string xml;
        if (!m_rowBreaks.empty()) {
            xml += "<rowBreaks count=\"";
            xml += std::to_string(m_rowBreaks.size());
            xml += "\" manualBreakCount=\"";
            xml += std::to_string(m_rowBreaks.size());
            xml += "\">";
            for (const auto& r : m_rowBreaks) {
                xml += "<brk id=\"";
                xml += r;
                xml += "\" max=\"16383\" man=\"1\"/>";
            }
            xml += "</rowBreaks>";
        }
        if (!m_colBreaks.empty()) {
            xml += "<colBreaks count=\"";
            xml += std::to_string(m_colBreaks.size());
            xml += "\" manualBreakCount=\"";
            xml += std::to_string(m_colBreaks.size());
            xml += "\">";
            for (const auto& c : m_colBreaks) {
                xml += "<brk id=\"";
                xml += c;
                xml += "\" max=\"1048575\" man=\"1\"/>";
            }
            xml += "</colBreaks>";
        }
        return xml;
    }

    // ── Cell / row writing (Phase 1) ─────────────────────────────────────────

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
                std::string strVal          = value.get<std::string>();
                bool        writtenAsShared = false;
                int32_t     sstIdx          = -1;

                if (m_options.useSharedStrings && m_worksheet) {
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
                        else {
                            const bool underUniqueCap =
                                static_cast<size_t>(sstView.stringCount()) < m_options.maxUniqueStrings;
                            const bool underByteCap =
                                (m_options.maxSstBytes == 0) ||
                                (m_sstBytesUsed + strVal.size() <= m_options.maxSstBytes);
                            if (underUniqueCap && underByteCap) {
                                const auto& sst = m_worksheet->sharedStrings();
                                sstIdx          = sst.getOrCreateStringIndex(strVal);
                                m_sstBytesUsed += strVal.size();
                                if (m_localCache.size() >= kLocalCacheLimit) { m_localCache.clear(); }
                                m_localCache.emplace(sstView.getStringView(sstIdx), sstIdx);
                                writtenAsShared = true;
                            }
                            // else: fall through to inlineStr (SST budget exhausted)
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

    template<typename T>
    void XLStreamWriter::setRowImpl(uint32_t row, uint16_t startCol, const std::vector<T>& items, const XLStreamRowOpts* opts)
    {
        checkStreamWritable();
        ensureSheetDataStarted();

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

            bool                 treatAsDate = false;
            const XLStreamCell*  cellMeta    = nullptr;
            if constexpr (std::is_same_v<T, XLCellValue>) { valPtr = &item; }
            else {
                valPtr      = &item.value;
                styleIdx    = item.styleIndex;
                treatAsDate = item.treatAsDate;
                cellMeta    = &item;
                if (item.formula.has_value()) formula = &item.formula.value();
            }

            if (treatAsDate &&
                (!styleIdx.has_value() || styleIdx.value() == XLDefaultCellFormat || styleIdx.value() == XLInvalidStyleIndex)) {
                const XLStyleIndex ds = ensureDefaultDateStyle();
                if (ds != XLInvalidStyleIndex) styleIdx = ds;
            }

            const bool hasFormula =
                (formula != nullptr && !formula->empty()) ||
                (cellMeta && (cellMeta->formulaType != XLStreamFormulaType::Normal || cellMeta->formulaSi.has_value()));
            const bool hasValue = valPtr->type() != XLValueType::Empty;

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
                    m_wroteFormula                  = true;
                    const std::string formulaText = formula ? *formula : std::string{};
                    if (!hasValue) {
                        m_writeBuffer += '>';
                        writeFormulaOpenTag(cellMeta, formulaText);
                        m_writeBuffer += "</c>";
                    }
                    else {
                        switch (valPtr->type()) {
                            case XLValueType::Integer: {
                                char numBuf[24];
                                auto [numPtr, ___] = std::to_chars(numBuf, numBuf + sizeof(numBuf), valPtr->get<int64_t>());
                                m_writeBuffer += R"( t="n">)";
                                writeFormulaOpenTag(cellMeta, formulaText);
                                m_writeBuffer += "<v>";
                                m_writeBuffer.append(numBuf, numPtr);
                                m_writeBuffer += "</v></c>";
                                break;
                            }
                            case XLValueType::Float: {
                                m_writeBuffer += R"( t="n">)";
                                writeFormulaOpenTag(cellMeta, formulaText);
                                m_writeBuffer += "<v>";
                                fmt::format_to(std::back_inserter(m_writeBuffer), "{}", valPtr->get<double>());
                                m_writeBuffer += "</v></c>";
                                break;
                            }
                            case XLValueType::Boolean: {
                                m_writeBuffer += R"( t="b">)";
                                writeFormulaOpenTag(cellMeta, formulaText);
                                m_writeBuffer += "<v>";
                                m_writeBuffer += (valPtr->get<bool>() ? '1' : '0');
                                m_writeBuffer += "</v></c>";
                                break;
                            }
                            case XLValueType::String: {
                                m_writeBuffer += R"( t="str">)";
                                writeFormulaOpenTag(cellMeta, formulaText);
                                m_writeBuffer += "<v>";
                                appendEscaped(m_writeBuffer, valPtr->get<std::string>());
                                m_writeBuffer += "</v></c>";
                                break;
                            }
                            default: {
                                m_writeBuffer += '>';
                                writeFormulaOpenTag(cellMeta, formulaText);
                                m_writeBuffer += "<v>";
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

        if (m_writeBuffer.size() >= kRowFlushThreshold) flushWriteBuffer();
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

    // ── Close / dimension ────────────────────────────────────────────────────

    void XLStreamWriter::close()
    {
        if (m_active) flushSheetDataClose();
    }

    void XLStreamWriter::patchDimension()
    {
        const uint32_t endRow = (m_lastWrittenRow > 0) ? m_lastWrittenRow : m_baseLastRow;
        const uint16_t endCol = (m_maxColumn > m_baseMaxCol) ? m_maxColumn : m_baseMaxCol;

        if (m_worksheet) m_worksheet->setStreamExtents(endRow, endCol);

        // OmitPatch: leave max-range placeholder; still record extents for diagnostics.
        if (m_options.dimensionMode == XLStreamDimensionMode::OmitPatch) return;

        // No fixed slot installed (corrupt header) — nothing safe to patch without a full reload.
        if (m_dimensionSlotSize == 0) return;

        std::string dimRef = "A1";
        if (endRow > 0 && endCol > 0) {
            dimRef = "A1:" + XLCellReference::columnAsString(endCol) + std::to_string(endRow);
        }
        const std::string slot = makeDimensionSlot(dimRef);
        if (slot.size() != m_dimensionSlotSize) {
            throw XLInternalError("XLStreamWriter: dimension slot size mismatch during patch");
        }

        if (m_usingFile) {
            // In-place seek/write — O(slot) I/O, never reloads the worksheet XML (P0a).
            if (m_fileStream.is_open()) {
                m_fileStream.flush();
                m_fileStream.seekp(static_cast<std::streamoff>(m_dimensionByteOffset));
                if (!m_fileStream.good()) {
                    throw XLInternalError("XLStreamWriter: failed to seek for dimension patch");
                }
                m_fileStream.write(slot.data(), static_cast<std::streamsize>(slot.size()));
                if (!m_fileStream.good()) {
                    throw XLInternalError("XLStreamWriter: failed to write dimension patch");
                }
                m_fileStream.flush();
            }
            else {
                std::fstream file(m_tempPath, std::ios::binary | std::ios::in | std::ios::out);
                if (!file) throw XLInternalError("XLStreamWriter: failed to reopen temp file for dimension patch");
                file.seekp(static_cast<std::streamoff>(m_dimensionByteOffset));
                file.write(slot.data(), static_cast<std::streamsize>(slot.size()));
                if (!file.good()) throw XLInternalError("XLStreamWriter: failed while writing dimension patch");
            }
        }
        else {
            if (m_dimensionByteOffset + m_dimensionSlotSize > m_memBuffer.size()) {
                throw XLInternalError("XLStreamWriter: dimension offset out of range in memory buffer");
            }
            std::memcpy(m_memBuffer.data() + static_cast<size_t>(m_dimensionByteOffset), slot.data(), slot.size());
        }
    }

    void XLStreamWriter::materializeToWorksheet()
    {
        if (!m_worksheet) return;
        if (m_usingFile) {
            m_worksheet->setStreamMaterialization(m_tempPath.string(), {});
        }
        else {
            m_worksheet->setStreamMaterialization({}, std::move(m_memBuffer));
            m_memBuffer.clear();
        }
    }

    void XLStreamWriter::flushSheetDataClose()
    {
        try {
            ensureSheetDataStarted();
            flushWriteBuffer();

            // Footer: start from bottomHalf (begins with </sheetData>), inject merges / breaks / tableParts.
            std::string footer = m_bottomHalf.empty() ? std::string(kFreshBottom) : m_bottomHalf;
            m_bottomHalf.clear();

            const std::string autoFilterXml = buildAutoFilterXml();
            const std::string mergeXml      = buildMergeCellsXml();
            const std::string breaksXml     = buildPageBreaksXml();

            std::string hyperlinksXml;
            if (!m_hyperlinks.empty() && m_worksheet) {
                hyperlinksXml = m_worksheet->registerStreamHyperlinks(m_hyperlinks);
            }

            // Insert after </sheetData> in OOXML-friendly order: autoFilter, merge, hyperlinks, breaks, tableParts
            const std::string sheetDataClose = "</sheetData>";
            size_t            insertPos      = footer.find(sheetDataClose);
            std::string       inject         = autoFilterXml + mergeXml + hyperlinksXml + breaksXml;
            if (m_table.has_value() && m_worksheet) {
                try {
                    inject += m_worksheet->registerStreamTable(*m_table);
                }
                catch (...) {
                    if (m_options.strictTableRegistration) throw;
                }
            }
            if (insertPos != std::string::npos) {
                insertPos += sheetDataClose.size();
                footer.insert(insertPos, inject);
            }
            else {
                footer = sheetDataClose + inject + footer;
            }

            sinkWrite(footer);

            // Patch dimension while the file handle is still open when possible (O(1) seek).
            patchDimension();

            if (m_usingFile && m_fileStream.is_open()) m_fileStream.close();

            // Package-only materialization: zip save uses temp path / mem buffer; no DOM reload (P0b).
            materializeToWorksheet();

            if (m_wroteFormula && m_worksheet) {
                try {
                    m_worksheet->parentDoc().setFormulaNeedsRecalculation(true);
                }
                catch (...) {
                }
            }
        }
        catch (...) {
            m_active = false;
            markWriterOpen(false);
            throw;
        }

        m_active = false;
        markWriterOpen(false);
    }

    template void XLStreamWriter::setRowImpl<XLCellValue>(uint32_t, uint16_t, const std::vector<XLCellValue>&, const XLStreamRowOpts*);
    template void XLStreamWriter::setRowImpl<XLStreamCell>(uint32_t, uint16_t, const std::vector<XLStreamCell>&, const XLStreamRowOpts*);

}    // namespace OpenXLSX

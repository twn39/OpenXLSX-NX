// Stream optimization tests: Phase 0/2 APIs, Fresh open, memory spill, merge, sparse read.

#include "OpenXLSX.hpp"
#include "TestHelpers.hpp"
#include "XLStreamReader.hpp"
#include "XLStreamWriter.hpp"

#include <catch2/catch_test_macros.hpp>
#include <filesystem>
#include <string>

using namespace OpenXLSX;

namespace
{
    inline std::string uniqueXlsx(const char* tag) { return OpenXLSX::TestHelpers::getUniqueFilename(tag) + ".xlsx"; }
}    // namespace

TEST_CASE("StreamWriterPhase0_ColWidthAndVisible", "[XLStreamWriter][Phase0]")
{
    const std::string filename = uniqueXlsx("stream_phase0_cols");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    {
        XLStreamWriteOptions opts;
        opts.openMode             = XLStreamOpenMode::Fresh;
        opts.memorySpillThreshold = 8 * 1024 * 1024;
        auto sw                   = wks.streamWriter(opts);
        sw.setColWidth(1, 2, 18.5);
        sw.setColVisible(3, 3, false);
        sw.setColOutlineLevel(4, 2);
        sw.appendRow({XLCellValue("A"), XLCellValue("B"), XLCellValue("C"), XLCellValue("D")});
        sw.appendRow({XLCellValue(1), XLCellValue(2), XLCellValue(3), XLCellValue(4)});
        REQUIRE_THROWS_AS(sw.setColWidth(1, 1, 10.0), XLInputError);
        sw.close();
    }

    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    const std::string sheetXml = rdoc.archive().getEntry("xl/worksheets/sheet1.xml");
    REQUIRE(sheetXml.find("<cols>") != std::string::npos);
    REQUIRE(sheetXml.find("width=") != std::string::npos);
    REQUIRE(sheetXml.find("hidden=\"1\"") != std::string::npos);

    auto rw = rdoc.workbook().worksheet("Sheet1");
    auto rd = rw.streamReader();
    REQUIRE(rd.hasNext());
    auto row1 = rd.nextRow();
    REQUIRE(row1.size() >= 1);
    REQUIRE(row1[0].getString() == "A");
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterPhase2_MergeCell", "[XLStreamWriter][Phase2]")
{
    const std::string filename = uniqueXlsx("stream_merge");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    {
        XLStreamWriteOptions opts;
        opts.openMode = XLStreamOpenMode::Fresh;
        auto sw       = wks.streamWriter(opts);
        sw.appendRow({XLCellValue("Title"), XLCellValue(""), XLCellValue("")});
        sw.mergeCell("A1", "C1");
        sw.appendRow({XLCellValue(10), XLCellValue(20), XLCellValue(30)});
        sw.close();
    }

    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    const std::string sheetXml = rdoc.archive().getEntry("xl/worksheets/sheet1.xml");
    REQUIRE(sheetXml.find("<mergeCells") != std::string::npos);
    REQUIRE(sheetXml.find("A1:C1") != std::string::npos);
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterFreshOpen", "[XLStreamWriter][Fresh]")
{
    const std::string filename = uniqueXlsx("stream_fresh");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    {
        auto sw = wks.streamWriter();
        sw.setRow("A1", {XLCellValue("hello"), XLCellValue(42)});
        sw.close();
    }
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    auto rw = rdoc.workbook().worksheet("Sheet1");
    auto rd = rw.streamReader();
    REQUIRE(rd.hasNext());
    auto row = rd.nextRow();
    REQUIRE(row[0].getString() == "hello");
    REQUIRE(row[1].get<int64_t>() == 42);
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterMemorySpillThreshold", "[XLStreamWriter][Spill]")
{
    const std::string filename = uniqueXlsx("stream_spill");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    XLStreamWriteOptions opts;
    opts.openMode             = XLStreamOpenMode::Fresh;
    opts.memorySpillThreshold = 2048;
    {
        auto sw = wks.streamWriter(opts);
        for (int i = 0; i < 200; ++i) {
            sw.appendRow({XLCellValue(std::string(64, 'x')), XLCellValue(i)});
        }
        sw.close();
    }
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    auto rw    = rdoc.workbook().worksheet("Sheet1");
    auto rd    = rw.streamReader();
    int  count = 0;
    while (rd.hasNext()) {
        auto row = rd.nextRow();
        if (!row.empty()) ++count;
    }
    REQUIRE(count == 200);
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterAddTableMinimal", "[XLStreamWriter][Table]")
{
    const std::string filename = uniqueXlsx("stream_table");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    {
        XLStreamWriteOptions opts;
        opts.openMode = XLStreamOpenMode::Fresh;
        auto sw       = wks.streamWriter(opts);
        sw.appendRow({XLCellValue("Name"), XLCellValue("Qty")});
        sw.appendRow({XLCellValue("a"), XLCellValue(1)});
        sw.appendRow({XLCellValue("b"), XLCellValue(2)});
        XLStreamTableOptions tbl;
        tbl.range   = "A1:B3";
        tbl.name    = "Sales";
        tbl.headers = {"Name", "Qty"};
        sw.addTable(tbl);
        sw.close();
    }
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    const std::string sheetXml = rdoc.archive().getEntry("xl/worksheets/sheet1.xml");
    REQUIRE(sheetXml.find("tableParts") != std::string::npos);
    const bool hasTable =
        rdoc.archive().hasEntry("xl/tables/table1.xml") || rdoc.archive().hasEntry("xl/tables/table2.xml");
    REQUIRE(hasTable);
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamReaderSparseAndInto", "[XLStreamReader][M6]")
{
    const std::string filename = uniqueXlsx("stream_sparse");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks                   = doc.workbook().worksheet("Sheet1");
    wks.cell("A1").value()     = "x";
    wks.cell("C1").value()     = 3;
    wks.cell("A2").value()     = "y";
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    auto rw = rdoc.workbook().worksheet("Sheet1");
    auto rd = rw.streamReader();
    REQUIRE(rd.hasNext());
    auto sparse = rd.nextRowSparse();
    REQUIRE(sparse.size() >= 1);
    bool sawC = false;
    for (const auto& c : sparse) {
        if (c.column == 3) {
            sawC = true;
            REQUIRE(c.value.get<int64_t>() == 3);
        }
    }
    REQUIRE(sawC);

    std::vector<XLCellValue> buf;
    REQUIRE(rd.nextRowInto(buf));
    REQUIRE(!buf.empty());
    REQUIRE(buf[0].getString() == "y");
    REQUIRE_FALSE(rd.nextRowInto(buf));
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamReaderIndexedSst", "[XLStreamReader][SST]")
{
    const std::string filename = uniqueXlsx("stream_indexed_sst");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks               = doc.workbook().worksheet("Sheet1");
    wks.cell("A1").value() = "shared-one";
    wks.cell("A2").value() = "shared-two";
    wks.cell("B1").value() = "shared-one";
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    XLStreamReadOptions opts;
    opts.sharedStringMode = XLStreamSharedStringMode::Indexed;
    auto rw               = rdoc.workbook().worksheet("Sheet1");
    auto rd               = rw.streamReader(opts);
    REQUIRE(rd.hasNext());
    auto r1 = rd.nextRow();
    REQUIRE(r1[0].getString() == "shared-one");
    REQUIRE(rd.hasNext());
    auto r2 = rd.nextRow();
    REQUIRE(r2[0].getString() == "shared-two");
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterHybridStillWorks", "[XLStreamWriter][Hybrid]")
{
    const std::string filename = uniqueXlsx("stream_hybrid");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks               = doc.workbook().worksheet("Sheet1");
    wks.cell("A1").value() = "dom";
    {
        XLStreamWriteOptions opts;
        opts.openMode = XLStreamOpenMode::Hybrid;
        auto sw       = wks.streamWriter(opts);
        sw.appendRow({XLCellValue("streamed")});
        sw.close();
    }
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    auto rw = rdoc.workbook().worksheet("Sheet1");
    auto rd = rw.streamReader();
    REQUIRE(rd.hasNext());
    auto r1 = rd.nextRow();
    REQUIRE(r1[0].getString() == "dom");
    REQUIRE(rd.hasNext());
    auto r2 = rd.nextRow();
    REQUIRE(r2[0].getString() == "streamed");
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterDimensionInPlacePatch_Spill", "[XLStreamWriter][P0a]")
{
    // Force temp-file sink so close() must seek-patch dimension without reloading the full XML.
    const std::string filename = uniqueXlsx("stream_dim_patch");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    XLStreamWriteOptions opts;
    opts.openMode             = XLStreamOpenMode::Fresh;
    opts.memorySpillThreshold = 0;    // always spill to temp file
    {
        auto sw = wks.streamWriter(opts);
        sw.appendRow({XLCellValue("a"), XLCellValue("b"), XLCellValue("c")});
        sw.appendRow({XLCellValue(1), XLCellValue(2), XLCellValue(3)});
        sw.appendRow({XLCellValue(4), XLCellValue(5), XLCellValue(6)});
        REQUIRE(sw.hasDimensionSlot());
        sw.close();
        REQUIRE(sw.lastRow() == 3);
        REQUIRE(sw.maxColumn() == 3);
    }
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    const std::string sheetXml = rdoc.archive().getEntry("xl/worksheets/sheet1.xml");
    REQUIRE(sheetXml.find(R"(ref="A1:C3")") != std::string::npos);
    // Must not leave the max-range placeholder after FixedSlot patch.
    REQUIRE(sheetXml.find("XFD1048576") == std::string::npos);
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterSstByteBudgetFallsBackToInline", "[XLStreamWriter][P0c]")
{
    const std::string filename = uniqueXlsx("stream_sst_budget");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    XLStreamWriteOptions opts;
    opts.openMode         = XLStreamOpenMode::Fresh;
    opts.useSharedStrings = true;
    opts.maxUniqueStrings = 100000;
    opts.maxSstBytes      = 16;    // tiny budget → most unique strings become inlineStr
    {
        auto sw = wks.streamWriter(opts);
        sw.appendRow({XLCellValue(std::string("abcdefghijklmnop")), XLCellValue(std::string("qrstuvwxyz012345"))});
        sw.close();
    }
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    const std::string sheetXml = rdoc.archive().getEntry("xl/worksheets/sheet1.xml");
    REQUIRE(sheetXml.find("inlineStr") != std::string::npos);
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamReaderMaxRowBytesGuard", "[XLStreamReader][P0d]")
{
    const std::string filename = uniqueXlsx("stream_max_row");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks               = doc.workbook().worksheet("Sheet1");
    wks.cell("A1").value() = std::string(200, 'Z');
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    XLStreamReadOptions opts;
    opts.maxRowBytes = 32;    // smaller than the physical row XML
    auto rw          = rdoc.workbook().worksheet("Sheet1");
    auto rd          = rw.streamReader(opts);
    REQUIRE_THROWS_AS((void)rd.hasNext(), XLInputError);
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterStructuredFreezePanes", "[XLStreamWriter][P1]")
{
    const std::string filename = uniqueXlsx("stream_panes");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    {
        XLStreamWriteOptions opts;
        opts.openMode = XLStreamOpenMode::Fresh;
        auto sw       = wks.streamWriter(opts);
        sw.freezePanes("B2");
        sw.appendRow({XLCellValue("H1"), XLCellValue("H2")});
        sw.appendRow({XLCellValue(1), XLCellValue(2)});
        sw.close();
    }
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    const std::string sheetXml = rdoc.archive().getEntry("xl/worksheets/sheet1.xml");
    REQUIRE(sheetXml.find("state=\"frozen\"") != std::string::npos);
    REQUIRE(sheetXml.find("topLeftCell=\"B2\"") != std::string::npos);
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterDateTimeAutoStyle", "[XLStreamWriter][P1]")
{
    const std::string filename = uniqueXlsx("stream_datetime");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    {
        auto sw = wks.streamWriter();
        sw.appendRow({XLStreamCell::withDateTime(XLDateTime::fromString("2024-06-15 12:30:00"))});
        sw.close();
    }
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    const std::string sheetXml = rdoc.archive().getEntry("xl/worksheets/sheet1.xml");
    REQUIRE(sheetXml.find(" s=\"") != std::string::npos);    // style index applied
    auto rw = rdoc.workbook().worksheet("Sheet1");
    auto rd = rw.streamReader();
    REQUIRE(rd.hasNext());
    auto row = rd.nextRowDetailed();
    REQUIRE(row[0].styleIndex.has_value());
    REQUIRE(row[0].value.getDouble() > 40000.0);    // serial near 2024
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamReaderColumnProjection", "[XLStreamReader][P1]")
{
    const std::string filename = uniqueXlsx("stream_proj");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks               = doc.workbook().worksheet("Sheet1");
    wks.cell("A1").value() = "a";
    wks.cell("B1").value() = "b";
    wks.cell("C1").value() = "c";
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    XLStreamReadOptions opts;
    opts.columnFilter = {1, 3};    // A and C only
    auto rw           = rdoc.workbook().worksheet("Sheet1");
    auto rd           = rw.streamReader(opts);
    REQUIRE(rd.hasNext());
    auto proj = rd.nextRowProjected();
    REQUIRE(proj.size() == 2);
    REQUIRE(proj[0].column == 1);
    REQUIRE(proj[0].value.getString() == "a");
    REQUIRE(proj[1].column == 3);
    REQUIRE(proj[1].value.getString() == "c");
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("DocumentMaxEntryUncompressedSize", "[XLDocument][P1]")
{
    const std::string filename = uniqueXlsx("stream_entry_limit");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks               = doc.workbook().worksheet("Sheet1");
    wks.cell("A1").value() = "hello";
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    rdoc.setMaxEntryUncompressedSize(10);    // too small for sheet XML
    REQUIRE_THROWS_AS(rdoc.archive().getEntry("xl/worksheets/sheet1.xml"), XLInputError);
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterHyperlinkAndAutoFilter", "[XLStreamWriter][P2]")
{
    const std::string filename = uniqueXlsx("stream_link_filter");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    {
        XLStreamWriteOptions opts;
        opts.openMode = XLStreamOpenMode::Fresh;
        auto sw       = wks.streamWriter(opts);
        sw.appendRow({XLCellValue("Name"), XLCellValue("Link")});
        sw.appendRow({XLCellValue("a"), XLCellValue("click")});
        sw.appendRow({XLCellValue("b"), XLCellValue("x")});
        sw.setAutoFilter("A1:B3");
        sw.addHyperlink("B2", "https://example.com", "ex");
        sw.addInternalHyperlink("B3", "Sheet1!A1");
        sw.close();
    }
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    const std::string sheetXml = rdoc.archive().getEntry("xl/worksheets/sheet1.xml");
    REQUIRE(sheetXml.find("<autoFilter ref=\"A1:B3\"") != std::string::npos);
    REQUIRE(sheetXml.find("<hyperlinks>") != std::string::npos);
    REQUIRE(sheetXml.find("r:id=") != std::string::npos);
    REQUIRE(sheetXml.find("location=") != std::string::npos);
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterSharedAndArrayFormula", "[XLStreamWriter][P2]")
{
    const std::string filename = uniqueXlsx("stream_formula_meta");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    {
        auto sw = wks.streamWriter();
        sw.appendRow({
            XLStreamCell::withArrayFormula("SUM(A1:A2)", "A3"),
            XLStreamCell::withSharedFormula(0, "B1+1", "B1:B2"),
        });
        sw.appendRow({
            XLCellValue(1),
            XLStreamCell::withSharedFormula(0),    // dependent
        });
        sw.close();
    }
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    const std::string sheetXml = rdoc.archive().getEntry("xl/worksheets/sheet1.xml");
    REQUIRE(sheetXml.find(R"(t="array")") != std::string::npos);
    REQUIRE(sheetXml.find(R"(t="shared")") != std::string::npos);
    REQUIRE(sheetXml.find(R"(si="0")") != std::string::npos);

    auto rd = rdoc.workbook().worksheet("Sheet1").streamReader();
    // Keep worksheet alive
    auto w2 = rdoc.workbook().worksheet("Sheet1");
    auto r2 = w2.streamReader();
    REQUIRE(r2.hasNext());
    auto row = r2.nextRowDetailed();
    REQUIRE(row.size() >= 1);
    bool sawArray = false;
    for (const auto& c : row) {
        if (c.formulaType && *c.formulaType == "array") sawArray = true;
    }
    REQUIRE(sawArray);
    rdoc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamReaderRichTextAndColumns", "[XLStreamReader][P2]")
{
    const std::string filename = uniqueXlsx("stream_rich_cols");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    {
        auto sw = wks.streamWriter();
        XLRichText rt;
        rt.addRun(XLRichTextRun("Hi").setBold(true));
        rt.addRun(XLRichTextRun("There"));
        sw.appendRow({XLCellValue(rt), XLCellValue(10)});
        sw.appendRow({XLCellValue("x"), XLCellValue(20)});
        sw.close();
    }
    doc.save();
    doc.close();

    XLDocument rdoc;
    rdoc.open(filename);
    auto               w = rdoc.workbook().worksheet("Sheet1");
    XLStreamReadOptions opts;
    opts.parseRichText = true;
    auto rd            = w.streamReader(opts);
    REQUIRE(rd.hasNext());
    auto row = rd.nextRowDetailed();
    REQUIRE(row[0].richText.has_value());
    REQUIRE(row[0].richText->runs().size() >= 1);

    // streamColumns consumes remaining + re-open reader for columns
    auto rd2 = w.streamReader();
    auto cols = rd2.streamColumns();
    REQUIRE(cols.next());
    REQUIRE(cols.currentColumn() == 1);
    REQUIRE(cols.values().size() == 2);
    REQUIRE(cols.next());
    REQUIRE(cols.currentColumn() == 2);
    REQUIRE(cols.values()[0].get<int64_t>() == 10);
    REQUIRE(cols.values()[1].get<int64_t>() == 20);
    REQUIRE_FALSE(cols.next());
    rdoc.close();
    std::filesystem::remove(filename);
}

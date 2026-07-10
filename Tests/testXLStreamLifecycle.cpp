#include "OpenXLSX.hpp"
#include "TestHelpers.hpp"
#include "XLStreamWriter.hpp"

#include <catch2/catch_approx.hpp>
#include <catch2/catch_test_macros.hpp>
#include <filesystem>
#include <string>
#include <utility>

using namespace OpenXLSX;

namespace {
inline std::string uniqueXlsx(const char* tag)
{
    return OpenXLSX::TestHelpers::getUniqueFilename(tag) + ".xlsx";
}
}    // namespace

TEST_CASE("StreamWriterSaveWhileActiveThrows", "[XLStreamWriter][Lifecycle]")
{
    const std::string filename = uniqueXlsx("stream_save_active");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto stream = wks.streamWriter();
    stream.appendRow({XLCellValue(1), XLCellValue(2)});

    REQUIRE_THROWS_AS(doc.save(), XLInputError);

    stream.close();
    REQUIRE_NOTHROW(doc.save());
    doc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterSecondWriterThrows", "[XLStreamWriter][Lifecycle]")
{
    const std::string filename = uniqueXlsx("stream_double_writer");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto stream = wks.streamWriter();

    REQUIRE_THROWS_AS(wks.streamWriter(), XLInputError);

    stream.close();
    REQUIRE_NOTHROW(wks.streamWriter().close());
    doc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterDoubleCloseIdempotent", "[XLStreamWriter][Lifecycle]")
{
    const std::string filename = uniqueXlsx("stream_double_close");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto stream = wks.streamWriter();
    stream.appendRow({XLCellValue("a")});
    stream.close();
    REQUIRE_FALSE(stream.isStreamActive());
    REQUIRE_NOTHROW(stream.close());
    REQUIRE_THROWS_AS(stream.appendRow({XLCellValue("b")}), XLInternalError);

    doc.save();
    doc.close();
    std::filesystem::remove(filename);
}

TEST_CASE("StreamWriterDimensionAfterClose", "[XLStreamWriter][Dimension]")
{
    const std::string filename = uniqueXlsx("stream_dimension");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();
        for (int i = 0; i < 100; ++i) {
            stream.appendRow({XLCellValue("R"), XLCellValue(i), XLCellValue(1.5), XLCellValue(true), XLCellValue("E")});
        }
        REQUIRE(stream.lastRow() == 100);
        REQUIRE(stream.maxColumn() == 5);
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell(100, 1).value().get<std::string>() == "R");
        REQUIRE(wks.cell(100, 2).value().get<int>() == 99);
        REQUIRE(wks.cell(100, 5).value().get<std::string>() == "E");

        // Dimension is in the worksheet XML; verify used range via row/column counts after DOM load.
        REQUIRE(wks.rowCount() == 100);
        REQUIRE(wks.columnCount() >= 5);

        // Spot-check raw XML dimension attribute via package path content.
        auto xml = wks.xmlDocument().document_element().child("dimension");
        REQUIRE_FALSE(xml.empty());
        std::string ref = xml.attribute("ref").value();
        REQUIRE(ref.find("100") != std::string::npos);

        doc.close();
        std::filesystem::remove(filename);
    }
}

TEST_CASE("StreamWriterFormulaRoundTrip", "[XLStreamWriter][Formula]")
{
    const std::string filename = uniqueXlsx("stream_formula");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();
        stream.appendRow({
            XLStreamCell(1),
            XLStreamCell(2),
            XLStreamCell::withFormula("A1+B1", XLCellValue(3)),
        });
        stream.appendRow({
            XLStreamCell::withFormula("SUM(A1:B1)"),
        });
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell(1, 1).value().get<int>() == 1);
        REQUIRE(wks.cell(1, 2).value().get<int>() == 2);
        REQUIRE(wks.cell(1, 3).hasFormula());
        REQUIRE(wks.cell(1, 3).formula().get().find("A1+B1") != std::string::npos);
        REQUIRE(wks.cell(1, 3).value().get<int>() == 3);
        REQUIRE(wks.cell(2, 1).hasFormula());
        REQUIRE(wks.cell(2, 1).formula().get().find("SUM") != std::string::npos);
        doc.close();
        std::filesystem::remove(filename);
    }
}

TEST_CASE("StreamWriterRowOptsAndSetRow", "[XLStreamWriter][RowOpts][SetRow]")
{
    const std::string filename = uniqueXlsx("stream_setrow");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();

        XLStreamRowOpts opts;
        opts.height = 24.0;
        opts.hidden = false;

        stream.setRow("A1", {XLCellValue("H1"), XLCellValue("H2")}, opts);
        stream.setRow("C3", {XLCellValue(10), XLCellValue(20)});    // skip row 2, start at col C

        REQUIRE_THROWS_AS(stream.setRow(2, 1, std::vector<XLCellValue>{XLCellValue(1)}), XLInputError);

        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell("A1").value().get<std::string>() == "H1");
        REQUIRE(wks.cell("B1").value().get<std::string>() == "H2");
        REQUIRE(wks.row(1).height() == Catch::Approx(24.0));
        REQUIRE(wks.cell("C3").value().get<int>() == 10);
        REQUIRE(wks.cell("D3").value().get<int>() == 20);
        REQUIRE(wks.cell("A3").value().type() == XLValueType::Empty);
        doc.close();
        std::filesystem::remove(filename);
    }
}

TEST_CASE("StreamWriterMovePreservesSaveGuard", "[XLStreamWriter][Lifecycle]")
{
    const std::string filename = uniqueXlsx("stream_move_guard");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    auto stream = wks.streamWriter();
    stream.appendRow({XLCellValue(42)});
    auto moved = std::move(stream);
    REQUIRE(moved.isStreamActive());
    REQUIRE_THROWS_AS(doc.save(), XLInputError);
    moved.close();
    REQUIRE_NOTHROW(doc.save());
    doc.close();
    std::filesystem::remove(filename);
}

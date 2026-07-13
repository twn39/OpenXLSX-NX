#include "OpenXLSX.hpp"
#include "TestHelpers.hpp"
#include "XLStreamReader.hpp"
#include "XLStreamWriter.hpp"

#include <catch2/catch_approx.hpp>
#include <catch2/catch_test_macros.hpp>
#include <filesystem>
#include <string>
#include <vector>

using namespace OpenXLSX;

namespace {
std::string uniqueXlsx(const char* tag)
{
    return OpenXLSX::TestHelpers::getUniqueFilename(tag) + ".xlsx";
}
}    // namespace

TEST_CASE("StreamReaderEmitEmptyRows", "[XLStreamReader][EmptyRows]")
{
    const std::string filename = uniqueXlsx("stream_emit_empty");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks               = doc.workbook().worksheet("Sheet1");
        wks.cell("A1").value() = 1;
        wks.cell("B3").value() = 3;    // gap: no row 2
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");

        XLStreamReadOptions opts;
        opts.emptyRows = XLStreamEmptyRowPolicy::EmitEmptyRows;
        auto reader    = wks.streamReader(opts);

        REQUIRE(reader.hasNext());
        auto r1 = reader.nextRow();
        REQUIRE(reader.currentRow() == 1);
        REQUIRE(r1.size() >= 1);
        REQUIRE(r1[0].get<int>() == 1);
        REQUIRE_FALSE(reader.currentRowOpts().isSyntheticEmpty);

        REQUIRE(reader.hasNext());
        auto r2 = reader.nextRow();
        REQUIRE(reader.currentRow() == 2);
        REQUIRE(r2.empty());
        REQUIRE(reader.currentRowOpts().isSyntheticEmpty);

        REQUIRE(reader.hasNext());
        auto r3 = reader.nextRow();
        REQUIRE(reader.currentRow() == 3);
        REQUIRE(r3.size() >= 2);
        REQUIRE(r3[0].type() == XLValueType::Empty);
        REQUIRE(r3[1].get<int>() == 3);
        REQUIRE_FALSE(reader.currentRowOpts().isSyntheticEmpty);

        REQUIRE_FALSE(reader.hasNext());
        doc.close();
        std::filesystem::remove(filename);
    }
}

TEST_CASE("StreamReaderFormulaAndStyleDetailed", "[XLStreamReader][Detailed]")
{
    const std::string filename = uniqueXlsx("stream_read_formula_style");

    XLStyleIndex styleId = 0;
    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);

        XLStyle currency;
        currency.numberFormat = "$#,##0.00";
        styleId               = doc.styles().findOrCreateStyle(currency);

        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();
        stream.appendRow({
            XLStreamCell(10.5, styleId),
            XLStreamCell::withFormula("A1*2", XLCellValue(21.0)),
            XLStreamCell(XLCellValue("txt")),
        });

        XLStreamRowOpts ropts;
        ropts.height = 30.0;
        stream.appendRow({XLStreamCell(1)}, ropts);
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto reader = wks.streamReader();

        REQUIRE(reader.hasNext());
        auto row = reader.nextRowDetailed();
        REQUIRE(reader.currentRow() == 1);
        REQUIRE(row.size() == 3);

        REQUIRE(row[0].value.get<double>() == Catch::Approx(10.5));
        REQUIRE(row[0].styleIndex.has_value());
        REQUIRE(row[0].styleIndex.value() == styleId);
        REQUIRE(row[0].column == 1);

        REQUIRE(row[1].formula.has_value());
        REQUIRE(row[1].formula->find("A1*2") != std::string::npos);
        REQUIRE(row[1].value.get<double>() == Catch::Approx(21.0));
        REQUIRE(row[1].column == 2);

        REQUIRE(row[2].value.get<std::string>() == "txt");
        REQUIRE_FALSE(row[2].formula.has_value());

        REQUIRE(reader.hasNext());
        auto row2 = reader.nextRowDetailed();
        REQUIRE(reader.currentRow() == 2);
        auto opts = reader.currentRowOpts();
        REQUIRE(opts.height.has_value());
        REQUIRE(opts.height.value() == Catch::Approx(30.0));

        reader.close();
        doc.close();
        std::filesystem::remove(filename);
    }
}

TEST_CASE("StreamReaderFormattedStrings", "[XLStreamReader][Format]")
{
    const std::string filename = uniqueXlsx("stream_read_format");

    XLStyleIndex styleId = 0;
    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        XLStyle pct;
        pct.numberFormat = "0%";
        styleId          = doc.styles().findOrCreateStyle(pct);

        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();
        stream.appendRow({XLStreamCell(0.25, styleId), XLStreamCell(XLCellValue("ok"))});
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");

        XLStreamReadOptions opts;
        opts.applyNumberFormats = true;
        auto reader             = wks.streamReader(opts);

        auto strings = reader.nextRowStrings();
        REQUIRE(strings.size() == 2);
        // 0.25 with 0% → "25%" typically
        REQUIRE(strings[0].find('2') != std::string::npos);
        REQUIRE(strings[0].find('%') != std::string::npos);
        REQUIRE(strings[1] == "ok");

        // Close entry stream before document close / reopen. Leaving the zip entry open while
        // closing the archive can lock the file on Windows and break the second open below.
        reader.close();
        doc.close();

        XLDocument doc2;
        doc2.open(filename);
        auto wks2     = doc2.workbook().worksheet("Sheet1");
        auto reader2  = wks2.streamReader(opts);
        auto detailed = reader2.nextRowDetailed();
        REQUIRE(detailed.size() == 2);
        auto formatted = reader2.formattedValue(detailed[0]);
        REQUIRE(formatted.find('%') != std::string::npos);

        reader2.close();
        doc2.close();
        std::filesystem::remove(filename);
    }
}

TEST_CASE("StreamReaderSkipMissingRowsDefault", "[XLStreamReader][EmptyRows]")
{
    // Legacy behaviour: skip missing row 2
    const std::string filename = uniqueXlsx("stream_skip_empty");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks               = doc.workbook().worksheet("Sheet1");
        wks.cell("A1").value() = 1;
        wks.cell("A3").value() = 3;
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto reader = doc.workbook().worksheet("Sheet1").streamReader();
        REQUIRE(reader.nextRow()[0].get<int>() == 1);
        REQUIRE(reader.currentRow() == 1);
        REQUIRE(reader.nextRow()[0].get<int>() == 3);
        REQUIRE(reader.currentRow() == 3);
        REQUIRE_FALSE(reader.hasNext());
        doc.close();
        std::filesystem::remove(filename);
    }
}

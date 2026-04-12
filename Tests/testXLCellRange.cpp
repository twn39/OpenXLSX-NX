#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLCellRangeTests", "[XLCellRange]")
{
    XLDocument doc;
    doc.create("./testXLCellRange.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    SECTION("Constructor")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto cl : rng) {
            INFO("Writing to cell: " << cl.cellReference().address());
            cl.value() = "Value";
        }
        doc.save();

        REQUIRE(wks.cell("B2").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("C2").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("D2").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("B3").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("C3").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("D3").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("B4").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("C4").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("D4").value().get<std::string>() == "Value");
    }

    SECTION("Copy Constructor")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto cl : rng) {
            INFO("Writing to cell: " << cl.cellReference().address());
            cl.value() = "Value";
        }
        doc.save();

        XLCellRange rng2 = rng;
        for (auto cl : rng2) cl.value() = "Value2";
        doc.save();

        REQUIRE(wks.cell("B2").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("C2").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("D2").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("B3").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("C3").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("D3").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("B4").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("C4").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("D4").value().get<std::string>() == "Value2");
    }

    SECTION("Move Constructor")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto cl : rng) {
            INFO("Writing to cell: " << cl.cellReference().address());
            cl.value() = "Value";
        }
        doc.save();

        XLCellRange rng2{std::move(rng)};
        for (auto cl : rng2) cl.value() = "Value3";
        doc.save();

        REQUIRE(wks.cell("B2").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("C2").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("D2").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("B3").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("C3").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("D3").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("B4").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("C4").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("D4").value().get<std::string>() == "Value3");
    }

    SECTION("Copy Assignment")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto cl : rng) {
            INFO("Writing to cell: " << cl.cellReference().address());
            cl.value() = "Value";
        }
        doc.save();

        XLCellRange rng2 = wks.range();
        rng2             = rng;
        for (auto cl : rng2) cl.value() = "Value4";
        doc.save();

        REQUIRE(wks.cell("B2").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("C2").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("D2").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("B3").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("C3").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("D3").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("B4").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("C4").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("D4").value().get<std::string>() == "Value4");
    }

    SECTION("Move Assignment")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto cl : rng) {
            INFO("Writing to cell: " << cl.cellReference().address());
            cl.value() = "Value";
        }
        doc.save();

        XLCellRange rng2 = wks.range();
        rng2             = std::move(rng);
        for (auto cl : rng2) cl.value() = "Value5";
        doc.save();

        REQUIRE(wks.cell("B2").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("C2").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("D2").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("B3").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("C3").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("D3").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("B4").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("C4").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("D4").value().get<std::string>() == "Value5");
    }

    SECTION("Functions")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto iter = rng.begin(); iter != rng.end(); iter++) iter->value() = "Value";

        doc.save();

        REQUIRE(rng.numColumns() == 3);
        REQUIRE(rng.numRows() == 3);

        rng.clear();

        REQUIRE(wks.cell("B2").value().get<std::string>().empty());
        REQUIRE(wks.cell("C2").value().get<std::string>().empty());
        REQUIRE(wks.cell("D2").value().get<std::string>().empty());
        REQUIRE(wks.cell("B3").value().get<std::string>().empty());
        REQUIRE(wks.cell("C3").value().get<std::string>().empty());
        REQUIRE(wks.cell("D3").value().get<std::string>().empty());
        REQUIRE(wks.cell("B4").value().get<std::string>().empty());
        REQUIRE(wks.cell("C4").value().get<std::string>().empty());
        REQUIRE(wks.cell("D4").value().get<std::string>().empty());
    }

    SECTION("XLCellIterator")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto iter = rng.begin(); iter != rng.end(); iter++) iter->value() = "Value";

        auto begin = rng.begin();
        REQUIRE(begin->cellReference().address() == "B2");

        auto iter2 = ++begin;
        REQUIRE(iter2->cellReference().address() == "C2");

        XLCellIterator iter3 = std::move(++begin);
        REQUIRE(iter3->cellReference().address() == "D2");

        auto iter4 = rng.begin();
        iter4      = ++iter3;
        REQUIRE(iter4->cellReference().address() == "B3");

        auto iter5 = rng.begin();
        iter5      = std::move(++iter3);
        REQUIRE(iter5->cellReference().address() == "C3");

        auto it1 = rng.begin();
        auto it2 = rng.begin();
        auto it3 = rng.end();

        REQUIRE(it1 == it2);
        REQUIRE_FALSE(it1 != it2);
        REQUIRE(it1 != it3);
        REQUIRE_FALSE(it1 == it3);

        REQUIRE(std::distance(it1, it3) == 9);
    }

    SECTION("Intersection")
    {
        auto rng1 = wks.range(XLCellReference("B2"), XLCellReference("D4"));    // B2:D4
        auto rng2 = wks.range(XLCellReference("C3"), XLCellReference("E5"));    // C3:E5

        auto intersect = rng1.intersect(rng2);
        REQUIRE(intersect.address() == "C3:D4");
        REQUIRE(intersect.numRows() == 2);
        REQUIRE(intersect.numColumns() == 2);
        REQUIRE_FALSE(intersect.empty());

        // Identical ranges
        auto rngIdentical = rng1.intersect(rng1);
        REQUIRE(rngIdentical.address() == "B2:D4");
        REQUIRE_FALSE(rngIdentical.empty());

        // Adjacent ranges (No intersection)
        auto rngAdjacent = wks.range(XLCellReference("E2"), XLCellReference("F4"));
        auto noIntersect = rng1.intersect(rngAdjacent);
        REQUIRE(noIntersect.empty());
        REQUIRE(noIntersect.numRows() == 0);
        REQUIRE(noIntersect.numColumns() == 0);
        REQUIRE(noIntersect.address() == "");

        // No intersection at all
        auto rngFar = wks.range(XLCellReference("Z10"), XLCellReference("AA20"));
        REQUIRE(rng1.intersect(rngFar).empty());
    }
}
TEST_CASE("XLCellRangeBulkOperations", "[XLCellRange]")
{
    XLDocument doc;
    doc.create("./testXLCellRange_bulk.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    auto range = wks.range("B2:D4");

    // 1. Matrix Write
    std::vector<std::vector<int>> matrix = {{1, 2, 3}, {4, 5, 6}, {7, 8, 9}};
    range.setValue(matrix);

    // Verify
    REQUIRE(wks.cell("B2").value().get<int>() == 1);
    REQUIRE(wks.cell("D4").value().get<int>() == 9);

    // 2. Matrix Read
    auto readMatrix = range.getValue<int>();
    REQUIRE(readMatrix.size() == 3);
    REQUIRE(readMatrix[0].size() == 3);
    REQUIRE(readMatrix[2][2] == 9);

    // 3. Batch Style Apply
    XLStyle style;
    style.font.bold    = true;
    style.fill.pattern = XLPatternSolid;
    style.fill.fgColor = XLColor("FFFF00");    // Yellow
    range.applyStyle(style);

    // 4. Set Border Outline
    range.setBorderOutline(XLLineStyleThick, XLColor("0000FF"));    // Blue thick border

    doc.save();
    doc.close();

    // Reopen and verify
    XLDocument doc2;
    doc2.open("./testXLCellRange_bulk.xlsx");
    auto  wks2   = doc2.workbook().worksheet("Sheet1");
    auto& styles = doc2.styles();

    // Verify cell B2 (Top-Left corner) has thick top and left borders
    auto formatB2 = styles.cellFormats().cellFormatByIndex(wks2.cell("B2").cellFormat());
    auto borderB2 = styles.borders().borderByIndex(formatB2.borderIndex());
    REQUIRE(borderB2.top().style() == XLLineStyleThick);
    REQUIRE(borderB2.left().style() == XLLineStyleThick);
    REQUIRE(borderB2.bottom().style() != XLLineStyleThick);

    // Verify cell C3 (Center cell) has NO thick borders
    auto formatC3 = styles.cellFormats().cellFormatByIndex(wks2.cell("C3").cellFormat());
    auto borderC3 = styles.borders().borderByIndex(formatC3.borderIndex());
    REQUIRE(borderC3.top().style() != XLLineStyleThick);
    REQUIRE(borderC3.left().style() != XLLineStyleThick);

    // Check background style applied via batch
    auto fillC3 = styles.fills().fillByIndex(formatC3.fillIndex());
    REQUIRE(fillC3.patternType() == XLPatternSolid);

    doc2.close();
}

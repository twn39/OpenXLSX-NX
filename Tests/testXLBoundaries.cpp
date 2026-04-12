#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <string>

using namespace OpenXLSX;

TEST_CASE("ExcelPhysicalBoundaries", "[Boundaries]")
{
    XLDocument doc;
    doc.create("./testXLBoundaries.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    SECTION("Row and Column Limits")
    {
        // 1. Valid Max Limits
        // These should not throw and should safely create the cell.
        REQUIRE_NOTHROW(wks.cell(1048576, 16384).value() = "Bottom Right");
        REQUIRE(wks.cell("XFD1048576").value().get<std::string>() == "Bottom Right");

        // OpenXLSX should handle this sparsely without allocating 17 billion cells.
        REQUIRE(wks.rowCount() == 1048576);
        REQUIRE(wks.columnCount() == 16384);

        // 2. Invalid Bound - Row Overflow
        REQUIRE_THROWS_AS(wks.cell(1048577, 1), XLCellAddressError);
        REQUIRE_THROWS_AS(wks.cell("A1048577"), XLInputError);

        // 3. Invalid Bound - Col Overflow (XFE = 16385)
        // Note: The worksheet class itself throws XLException when validating columns
        REQUIRE_THROWS_AS(wks.cell(1, 16385), XLException);
        REQUIRE_THROWS_AS(wks.cell("XFE1"), XLInputError);

        // 4. Invalid Bound - Zero Index
        REQUIRE_THROWS_AS(wks.cell(0, 1), XLCellAddressError);
        REQUIRE_THROWS_AS(wks.cell(1, 0), XLException);
    }

    SECTION("Data Volume Limits (C++ Memory Boundaries)")
    {
        // Microsoft Excel strictly limits cell text to 32,767 characters.
        // OpenXLSX does not artificially enforce this, which allows it to process files gracefully.
        // But we must guarantee that pushing >32k characters does not corrupt memory or segfault.
        std::string hugeString(35000, 'A');

        REQUIRE_NOTHROW(wks.cell("A1").value() = hugeString);
        REQUIRE(wks.cell("A1").value().get<std::string>() == hugeString);

        // Save and reload to ensure the XML writer and libzip handle the large text block properly
        doc.save();
        doc.close();

        doc.open("./testXLBoundaries.xlsx");
        auto wks2 = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks2.cell("A1").value().get<std::string>() == hugeString);
    }

    SECTION("Worksheet Name Memory Limits")
    {
        // Excel limits names to 31 chars.
        std::string validName(31, 'W');
        REQUIRE_NOTHROW(doc.workbook().addWorksheet(validName));

        // OpenXLSX actually strictly enforces this and throws an XLInputError
        std::string overLimitName(32, 'X');
        REQUIRE_THROWS_AS(doc.workbook().addWorksheet(overLimitName), XLInputError);
    }

    SECTION("Maximum Worksheets Volume (Stress Test)")
    {
        // Excel does not have a strict hardcoded limit for sheets (memory bounded),
        // but older versions were limited to 255. Let's ensure creating 255 doesn't crash us.
        for (int i = 0; i < 255; ++i) { REQUIRE_NOTHROW(doc.workbook().addWorksheet("MassSheet_" + std::to_string(i))); }
        REQUIRE(doc.workbook().worksheetCount() >= 255);
    }
}
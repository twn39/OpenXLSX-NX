#include "OpenXLSX.hpp"
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLStreamReader_large_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLStreamReader_sax_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLStreamReader_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_3() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLStreamReader_skip_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("StreamingReaderFunctionalTests", "[XLStreamReader]")
{
    // Generate a file using StreamWriter
    {
        XLDocument doc;
        doc.create(__global_unique_file_2(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        auto writer = wks.streamWriter();
        writer.appendRow({"String", 123, 45.6, true, "Special <>&\""});
        writer.appendRow({XLCellValue(), "Skipped First"});
        writer.close();

        doc.save();
        doc.close();
    }

    // Now test StreamReader
    XLDocument doc;
    doc.open(__global_unique_file_2());
    auto wks = doc.workbook().worksheet("Sheet1");

    auto reader = wks.streamReader();

    REQUIRE(reader.hasNext() == true);
    auto row1 = reader.nextRow();
    REQUIRE(row1.size() == 5);
    REQUIRE(row1[0].get<std::string>() == "String");
    REQUIRE(row1[1].get<int>() == 123);
    REQUIRE(row1[2].get<double>() == 45.6);
    REQUIRE(row1[3].get<bool>() == true);
    REQUIRE(row1[4].get<std::string>() == "Special <>&\"");

    REQUIRE(reader.hasNext() == true);
    auto row2 = reader.nextRow();
    REQUIRE(row2.size() == 2);
    REQUIRE(row2[0].type() == XLValueType::Empty);
    REQUIRE(row2[1].get<std::string>() == "Skipped First");

    REQUIRE(reader.hasNext() == false);

    doc.close();
}
TEST_CASE("StreamingReaderSkippingEmptyRowsandCells", "[XLStreamReader]")
{
    {
        XLDocument doc;
        doc.create(__global_unique_file_3(), XLForceOverwrite);
        auto wks               = doc.workbook().worksheet("Sheet1");
        wks.cell("A1").value() = 1;
        wks.cell("C1").value() = 2;
        wks.cell("B3").value() = 3;    // Skips row 2, and skips col A in row 3
        doc.save();
        doc.close();
    }

    XLDocument doc;
    doc.open(__global_unique_file_3());
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto reader = wks.streamReader();

    REQUIRE(reader.hasNext() == true);
    auto row1 = reader.nextRow();
    REQUIRE(reader.currentRow() == 1);
    REQUIRE(row1.size() == 3);
    REQUIRE(row1[0].get<int>() == 1);
    REQUIRE(row1[1].type() == XLValueType::Empty);
    REQUIRE(row1[2].get<int>() == 2);

    REQUIRE(reader.hasNext() == true);
    auto row3 = reader.nextRow();
    REQUIRE(reader.currentRow() == 3);
    REQUIRE(row3.size() == 2);
    REQUIRE(row3[0].type() == XLValueType::Empty);
    REQUIRE(row3[1].get<int>() == 3);

    REQUIRE(reader.hasNext() == false);
}

TEST_CASE("StreamingReaderLargeFileTest", "[XLStreamReader]")
{
    {
        XLDocument doc;
        doc.create(__global_unique_file_0(), XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto writer = wks.streamWriter();
        for (int i = 0; i < 10000; i++) { writer.appendRow({"Row", i, i * 1.5, true}); }
        writer.close();
        doc.save();
        doc.close();
    }

    XLDocument doc;
    doc.open(__global_unique_file_0());
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto reader = wks.streamReader();

    int count = 0;
    while (reader.hasNext()) {
        auto row = reader.nextRow();
        REQUIRE(row.size() == 4);
        REQUIRE(row[1].get<int>() == count);
        count++;
    }
    REQUIRE(count == 10000);
    doc.close();
}

TEST_CASE("SAXParserEdgeCases", "[XLStreamReader][SAX]")
{
    // Write a file with boolean, error cell, and column-skipping patterns
    {
        XLDocument doc;
        doc.create(__global_unique_file_1(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // Row 1: bool true, bool false
        wks.cell("A1").value() = true;
        wks.cell("B1").value() = false;

        // Row 2: gap at col B (only A and C set)
        wks.cell("A2").value() = 10;
        wks.cell("C2").value() = 30;

        // Row 3: shared string via DOM path (sets a string value)
        wks.cell("A3").value() = "SAX test";

        doc.save();
        doc.close();
    }

    XLDocument doc;
    doc.open(__global_unique_file_1());
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto reader = wks.streamReader();

    // Row 1 — booleans
    REQUIRE(reader.hasNext());
    auto r1 = reader.nextRow();
    REQUIRE(reader.currentRow() == 1);
    REQUIRE(r1.size() == 2);
    REQUIRE(r1[0].get<bool>() == true);
    REQUIRE(r1[1].get<bool>() == false);

    // Row 2 — gap: B is empty
    REQUIRE(reader.hasNext());
    auto r2 = reader.nextRow();
    REQUIRE(reader.currentRow() == 2);
    REQUIRE(r2.size() == 3);
    REQUIRE(r2[0].get<int>() == 10);
    REQUIRE(r2[1].type() == XLValueType::Empty);
    REQUIRE(r2[2].get<int>() == 30);

    // Row 3 — shared string read via SAX reader
    REQUIRE(reader.hasNext());
    auto r3 = reader.nextRow();
    REQUIRE(reader.currentRow() == 3);
    REQUIRE(r3[0].get<std::string>() == "SAX test");

    REQUIRE_FALSE(reader.hasNext());
    doc.close();
}

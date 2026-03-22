#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Streaming Reader Functional Tests", "[XLStreamReader]")
{
    // Generate a file using StreamWriter
    {
        XLDocument doc;
        doc.create("./testXLStreamReader.xlsx", XLForceOverwrite);
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
    doc.open("./testXLStreamReader.xlsx");
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
TEST_CASE("Streaming Reader Skipping Empty Rows and Cells", "[XLStreamReader]")
{
    {
        XLDocument doc;
        doc.create("./testXLStreamReader_skip.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        wks.cell("A1").value() = 1;
        wks.cell("C1").value() = 2;
        wks.cell("B3").value() = 3; // Skips row 2, and skips col A in row 3
        doc.save();
        doc.close();
    }

    XLDocument doc;
    doc.open("./testXLStreamReader_skip.xlsx");
    auto reader = doc.workbook().worksheet("Sheet1").streamReader();
    
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
    doc.close();
}

TEST_CASE("Streaming Reader Large File Test", "[XLStreamReader]")
{
    {
        XLDocument doc;
        doc.create("./testXLStreamReader_large.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        auto writer = wks.streamWriter();
        for (int i = 0; i < 10000; i++) {
            writer.appendRow({"Row", i, i * 1.5, true});
        }
        writer.close();
        doc.save();
        doc.close();
    }

    XLDocument doc;
    doc.open("./testXLStreamReader_large.xlsx");
    auto reader = doc.workbook().worksheet("Sheet1").streamReader();
    
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

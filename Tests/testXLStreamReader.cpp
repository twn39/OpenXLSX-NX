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
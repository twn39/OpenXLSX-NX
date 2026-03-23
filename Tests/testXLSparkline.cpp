#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Sparkline Creation", "[XLSparkline]")
{
    // === Functionality Setup ===
    XLDocument doc;
    doc.create("./SparklineTest.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Add some sample data
    for (int col = 1; col <= 5; ++col) {
        wks.cell(1, col).value() = col * 10;
        wks.cell(2, col).value() = col * -5;
    }

    // Add multiple sparklines
    wks.addSparkline("F1", "A1:E1", XLSparklineType::Line);
    wks.addSparkline("F2", "A2:E2", XLSparklineType::Column);
    wks.addSparkline("G2", "A2:E2", XLSparklineType::Stacked);

    doc.save();
    doc.close();

    // === Read-back and OOXML Verification ===
    XLDocument doc2;
    doc2.open("./SparklineTest.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");
    
    // Check if basic data was saved correctly
    REQUIRE(wks2.cell("A1").value().get<int64_t>() == 10);
    REQUIRE(wks2.cell("A2").value().get<int64_t>() == -5);

    // Extract the raw XML from the document archive
    std::string sheetXmlStr = doc2.extractXmlFromArchive("xl/worksheets/sheet1.xml");

    // 1. Verify that 'x14' namespace was correctly appended to mc:Ignorable
    REQUIRE(sheetXmlStr.find("mc:Ignorable=\"x14ac x14\"") != std::string::npos);

    // 2. Verify that the sparklineGroups block exists under extLst
    REQUIRE(sheetXmlStr.find("<extLst>") != std::string::npos);
    REQUIRE(sheetXmlStr.find("xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"") != std::string::npos);
    REQUIRE(sheetXmlStr.find("<x14:sparklineGroups xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">") != std::string::npos);

    // 3. Verify types of sparkline groups exist
    REQUIRE(sheetXmlStr.find("type=\"line\"") != std::string::npos);
    REQUIRE(sheetXmlStr.find("type=\"column\"") != std::string::npos);
    REQUIRE(sheetXmlStr.find("type=\"stacked\"") != std::string::npos);

    // 4. Verify that formatting (like colorAxis) was created
    REQUIRE(sheetXmlStr.find("<x14:colorAxis rgb=\"FF000000\" />") != std::string::npos);

    // 5. Verify the formula logic (xm:f) and square reference (xm:sqref)
    REQUIRE(sheetXmlStr.find("<xm:f>Sheet1!A1:E1</xm:f>") != std::string::npos);
    REQUIRE(sheetXmlStr.find("<xm:sqref>F1</xm:sqref>") != std::string::npos);
    
    REQUIRE(sheetXmlStr.find("<xm:f>Sheet1!A2:E2</xm:f>") != std::string::npos);
    REQUIRE(sheetXmlStr.find("<xm:sqref>G2</xm:sqref>") != std::string::npos);

    doc2.close();
}

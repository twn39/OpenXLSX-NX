#include "TestHelpers.hpp"
#include "OpenXLSX.hpp"
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__SparklineAdvTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__SparklineTest_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("SparklineCreation", "[XLSparkline]")
{
    // === Functionality Setup ===
    XLDocument doc;
    doc.create(__global_unique_file_1(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Add some sample data
    for (uint16_t col = 1; col <= 5; ++col) {
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
    doc2.open(__global_unique_file_1());
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
    REQUIRE(sheetXmlStr.find("<x14:sparklineGroups xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">") !=
            std::string::npos);

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

TEST_CASE("SparklineAdvancedConfiguration", "[XLSparkline]")
{
    XLDocument doc;
    doc.create(__global_unique_file_0(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    for (uint16_t col = 1; col <= 5; ++col) {
        wks.cell(1, col).value() = col * 10;
    }

    XLSparklineOptions opts;
    opts.type = XLSparklineType::Column;
    opts.high = true;
    opts.low = true;
    opts.first = true;
    opts.last = true;
    opts.negative = true;
    opts.markers = true;
    opts.displayXAxis = true;
    opts.displayEmptyCellsAs = "zero";
    opts.seriesColor = "FF112233";
    opts.highMarkerColor = "FF445566";
    
    wks.addSparkline("F1", "A1:E1", opts);

    doc.save();
    doc.close();

    // Verify
    XLDocument doc2;
    doc2.open(__global_unique_file_0());
    std::string sheetXmlStr = doc2.extractXmlFromArchive("xl/worksheets/sheet1.xml");

    REQUIRE(sheetXmlStr.find("displayEmptyCellsAs=\"zero\"") != std::string::npos);
    REQUIRE(sheetXmlStr.find("markers=\"1\"") != std::string::npos);
    REQUIRE(sheetXmlStr.find("high=\"1\"") != std::string::npos);
    REQUIRE(sheetXmlStr.find("low=\"1\"") != std::string::npos);
    REQUIRE(sheetXmlStr.find("first=\"1\"") != std::string::npos);
    REQUIRE(sheetXmlStr.find("last=\"1\"") != std::string::npos);
    REQUIRE(sheetXmlStr.find("negative=\"1\"") != std::string::npos);
    REQUIRE(sheetXmlStr.find("displayXAxis=\"1\"") != std::string::npos);
    
    // Verify custom colors
    bool ok1 = sheetXmlStr.find("<x14:colorSeries rgb=\"FF112233\" />") != std::string::npos || sheetXmlStr.find("<x14:colorSeries rgb=\"FF112233\"/>") != std::string::npos;
    REQUIRE(ok1);
    bool ok2 = sheetXmlStr.find("<x14:colorHigh rgb=\"FF445566\" />") != std::string::npos || sheetXmlStr.find("<x14:colorHigh rgb=\"FF445566\"/>") != std::string::npos;
    REQUIRE(ok2);

    doc2.close();
    std::remove(__global_unique_file_0().c_str());
}

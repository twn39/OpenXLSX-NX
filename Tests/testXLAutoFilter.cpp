#include "OpenXLSX.hpp"
#include "XLDynamicFilter.hpp"
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__AutoFilterAdvancedTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__AutoFilterTest_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("AutoFilterandSortStateGeneration", "[XLAutoFilter]")
{
    // === Functionality Setup ===
    XLDocument doc;
    doc.create(__global_unique_file_1(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Add some sample data
    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Product";
    wks.cell("C1").value() = "Sales";

    wks.cell("A2").value() = "North";
    wks.cell("B2").value() = "Apples";
    wks.cell("C2").value() = 100;

    wks.cell("A3").value() = "South";
    wks.cell("B3").value() = "Bananas";
    wks.cell("C3").value() = 300;

    wks.cell("A4").value() = "North";
    wks.cell("B4").value() = "Oranges";
    wks.cell("C4").value() = 150;

    wks.cell("A5").value() = "East";
    wks.cell("B5").value() = "Apples";
    wks.cell("C5").value() = 50;

    // Set AutoFilter Range
    wks.setAutoFilter(wks.range(XLCellReference("A1"), XLCellReference("C5")));

    // Apply a filter condition to Column 0 ("Region")
    auto filterRegion = wks.autofilterObject().filterColumn(0);
    filterRegion.addFilter("North");

    // Apply a dynamic filter to Column 2 ("Sales")
    auto filterSales = wks.autofilterObject().filterColumn(2);
    filterSales.setDynamicFilter(XLDynamicFilterTypeToString(XLDynamicFilterType::AboveAverage));

    // Apply a sort condition (Sort by Sales Descending)
    wks.addSortCondition("A1:C5", 2, true);

    // Simulate auto filter application (should hide "South" and "East" rows)
    wks.applyAutoFilter();

    doc.save();
    doc.close();

    // === Read-back and OOXML Verification ===
    XLDocument doc2;
    doc2.open(__global_unique_file_1());
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    REQUIRE(wks2.hasAutoFilter() == true);
    REQUIRE(wks2.autoFilter() == "A1:C5");

    // The "South" and "East" rows (3 and 5) should be hidden by applyAutoFilter
    REQUIRE(wks2.row(3).isHidden() == true);
    REQUIRE(wks2.row(5).isHidden() == true);
    // The "North" rows (2 and 4) should be visible
    REQUIRE(wks2.row(2).isHidden() == false);
    REQUIRE(wks2.row(4).isHidden() == false);

    // Verify XML structure
    std::string sheetXmlStr = doc2.extractXmlFromArchive("xl/worksheets/sheet1.xml");

    // Check autoFilter parent node
    REQUIRE(sheetXmlStr.find("<autoFilter ref=\"A1:C5\">") != std::string::npos);

    // Check standard filter
    REQUIRE(sheetXmlStr.find("<filterColumn colId=\"0\">") != std::string::npos);
    REQUIRE(sheetXmlStr.find("<filter val=\"North\" />") != std::string::npos);

    // Check dynamic filter
    REQUIRE(sheetXmlStr.find("<filterColumn colId=\"2\">") != std::string::npos);
    REQUIRE(sheetXmlStr.find("<dynamicFilter type=\"aboveAverage\" />") != std::string::npos);

    // Check sortState as child of autoFilter, and proper row offset
    REQUIRE(sheetXmlStr.find("<sortState ref=\"A1:C5\">") != std::string::npos);
    REQUIRE(sheetXmlStr.find("<sortCondition ref=\"C2:C5\" descending=\"1\" />") != std::string::npos);

    doc2.close();
}

TEST_CASE("AutoFilterAdvancedConditions", "[XLAutoFilter][Advanced]")
{
    XLDocument doc;
    doc.create(__global_unique_file_0(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "ID";
    wks.cell("B1").value() = "Value";
    
    for(int i=2; i<=6; ++i) {
        wks.cell(i, 1).value() = i-1;
        wks.cell(i, 2).value() = (i-1) * 10;
    }

    wks.setAutoFilter(wks.range(XLCellReference("A1"), XLCellReference("B6")));
    
    // Test CustomFilter (e.g., Value > 20 AND Value < 50)
    auto filterValue = wks.autofilterObject().filterColumn(1);
    filterValue.setCustomFilter("greaterThan", "20", XLFilterLogic::And, "lessThan", "50");
    
    // Test Top10 Filter
    auto filterId = wks.autofilterObject().filterColumn(0);
    filterId.setTop10(3, false, true); // Top 3 items

    doc.save();
    doc.close();

    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open(__global_unique_file_0()));
    
    std::string sheetXmlStr = doc2.extractXmlFromArchive("xl/worksheets/sheet1.xml");

    // Verify Custom Filter XML
    REQUIRE(sheetXmlStr.find("<customFilters and=\"1\">") != std::string::npos);
    REQUIRE(sheetXmlStr.find("<customFilter operator=\"greaterThan\" val=\"20\"") != std::string::npos);
    REQUIRE(sheetXmlStr.find("<customFilter operator=\"lessThan\" val=\"50\"") != std::string::npos);

    // Verify Top10 Filter XML
    REQUIRE(sheetXmlStr.find("<top10 val=\"3\"") != std::string::npos);

    doc2.close();
    std::remove(__global_unique_file_0());
}

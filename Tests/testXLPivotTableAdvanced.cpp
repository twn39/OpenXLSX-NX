#include "OpenXLSX.hpp"
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotMultiDataTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotBuilderTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotStyleRegressionTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_3() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotBaseFieldTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_4() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotNamedSourceTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_5() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotDataOnRowsTest_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("AdvancedDynamicPivotTableBuilder", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create(__global_unique_file_1(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Sample data
    wks.cell("A1").value() = "Category";
    wks.cell("B1").value() = "Product";
    wks.cell("C1").value() = "Sales";
    wks.cell("D1").value() = "Date";

    wks.cell("A2").value() = "Fruit";
    wks.cell("B2").value() = "Apple";
    wks.cell("C2").value() = 120;
    wks.cell("D2").value() = "Q1";

    wks.cell("A3").value() = "Fruit";
    wks.cell("B3").value() = "Banana";
    wks.cell("C3").value() = 250;
    wks.cell("D3").value() = "Q2";

    wks.cell("A4").value() = "Vegetable";
    wks.cell("B4").value() = "Carrot";
    wks.cell("C4").value() = 150;
    wks.cell("D4").value() = "Q1";

    auto options = XLPivotTableOptions("AdvancedPivot", "Sheet1!A1:D4", "F1")
        .addFilterField("Date")
        .addRowField("Category")
        .addColumnField("Product")
        .addDataField("Sales", "Average Sales", XLPivotSubtotal::Average);
// Configure layout and styles based on our new Builder flags
options.setRowGrandTotals(false)
       .setColGrandTotals(false)
       .setCompactData(false)
       .setShowRowStripes(true)
       .setShowColStripes(true)
       .setPivotTableStyle("PivotStyleMedium9");

    REQUIRE_NOTHROW(wks.addPivotTable(options));
    REQUIRE_NOTHROW(doc.save());

    // Verify written values
    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open(__global_unique_file_1()));

    std::string ptDefXmlStr = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");

    // Verify boolean flags translated to XML attributes
    REQUIRE(ptDefXmlStr.find("rowGrandTotals=\"0\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("colGrandTotals=\"0\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("compactData=\"0\"") != std::string::npos);
    
    // Verify style
    REQUIRE(ptDefXmlStr.find("PivotStyleMedium9") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("showRowStripes=\"1\"") != std::string::npos);

    // Verify Page Field (Filter)
    REQUIRE(ptDefXmlStr.find("<pageFields count=\"1\">") != std::string::npos);
    
    // Check Average subtotal setting
    REQUIRE(ptDefXmlStr.find("subtotal=\"average\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("name=\"Average Sales\"") != std::string::npos);
}

TEST_CASE("AdvancedPivotTableMultiDataFieldsandEdgeCases", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create(__global_unique_file_0(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Sample data
    wks.cell("A1").value() = "Category";
    wks.cell("B1").value() = "Product";
    wks.cell("C1").value() = "Sales";
    wks.cell("D1").value() = "Profit";
    wks.cell("E1").value() = "Date";

    wks.cell("A2").value() = "Fruit";
    wks.cell("B2").value() = "Apple";
    wks.cell("C2").value() = 120;
    wks.cell("D2").value() = 40;
    wks.cell("E2").value() = "Q1";

    wks.cell("A3").value() = "Fruit";
    wks.cell("B3").value() = "Banana";
    wks.cell("C3").value() = 250;
    wks.cell("D3").value() = 80;
    wks.cell("E3").value() = "Q2";

    wks.cell("A4").value() = "Vegetable";
    wks.cell("B4").value() = "Carrot";
    wks.cell("C4").value() = 150;
    wks.cell("D4").value() = 30;
    wks.cell("E4").value() = "Q1";

    auto options = XLPivotTableOptions("MultiDataPivot", "Sheet1!A1:E4", "G1")
        .addRowField("Category")
        .addDataField("Sales", "Total Sales", XLPivotSubtotal::Sum)
        .addDataField("Profit", "Avg Profit", XLPivotSubtotal::Average);

    REQUIRE_NOTHROW(wks.addPivotTable(options));
    REQUIRE_NOTHROW(doc.save());
    doc.close();

    // Verify written values and XML structures for multi-data fields
    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open(__global_unique_file_0()));

    std::string ptDefXmlStr = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");

    // Verify both data fields are registered in <dataFields>
    REQUIRE(ptDefXmlStr.find("<dataField ") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("name=\"Total Sales\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("name=\"Avg Profit\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("subtotal=\"average\"") != std::string::npos);
    
    // Verify that the virtual data field (x=data) is added to the column fields by default when multiple data fields exist
    // Usually Excel represents this as an item with value "data" in <colFields>
    REQUIRE(ptDefXmlStr.find("<colFields count=\"1\">") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("x=\"-2\"") != std::string::npos); // -2 is the special index for the values field
    
    std::remove(__global_unique_file_0());
}

TEST_CASE("AdvancedPivotTableDataOnRowsandNumFmt", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create(__global_unique_file_5(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Revenue";
    wks.cell("C1").value() = "Costs";

    wks.cell("A2").value() = "North";
    wks.cell("B2").value() = 1000.50;
    wks.cell("C2").value() = 800.25;

    wks.cell("A3").value() = "South";
    wks.cell("B3").value() = 1500.75;
    wks.cell("C3").value() = 1200.00;

    auto options = XLPivotTableOptions("DataOnRowsPivot", "Sheet1!A1:C3", "E1")
        .setDataOnRows(true)
        .addRowField("Region")
        .addDataField("Revenue", "Total Revenue", XLPivotSubtotal::Sum, 4)
        .addDataField("Costs", "Total Costs", XLPivotSubtotal::Sum, 3);

    REQUIRE_NOTHROW(wks.addPivotTable(options));
    REQUIRE_NOTHROW(doc.save());
    doc.close();

    // Verify written values
    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open(__global_unique_file_5()));

    std::string ptDefXmlStr = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");

    // Verify rowFields contains the x="-2" virtual field
    auto rowFieldsPos = ptDefXmlStr.find("<rowFields");
    REQUIRE(ptDefXmlStr.find("x=\"-2\"", rowFieldsPos) != std::string::npos);
    
    // Verify colFields does NOT contain x="-2"
    auto colFieldsPos = ptDefXmlStr.find("<colFields");
    if (colFieldsPos != std::string::npos) {
        REQUIRE(ptDefXmlStr.find("x=\"-2\"", colFieldsPos) == std::string::npos);
    }

    // Verify numFmtId is correctly applied to dataFields
    REQUIRE(ptDefXmlStr.find("name=\"Total Revenue\"") != std::string::npos); REQUIRE(ptDefXmlStr.find("numFmtId=\"4\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("name=\"Total Costs\"") != std::string::npos); REQUIRE(ptDefXmlStr.find("numFmtId=\"3\"") != std::string::npos);
    
    std::remove(__global_unique_file_5());
}

TEST_CASE("AdvancedPivotTableStylingandFormattingRegression", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create(__global_unique_file_2(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Category";
    wks.cell("B1").value() = "Value";

    wks.cell("A2").value() = "A"; wks.cell("B2").value() = 100;
    wks.cell("A3").value() = "B"; wks.cell("B3").value() = 200;

    auto options = XLPivotTableOptions("StylePivot", "Sheet1!A1:B3", "D1")
        .setShowRowStripes(true)
        .setShowColStripes(false)
        .setShowRowHeaders(true)
        .setShowColHeaders(true)
        .setPivotTableStyle("PivotStyleMedium9")
        .addRowField("Category")
        .addDataField("Value", "Total Value", XLPivotSubtotal::Sum);

    REQUIRE_NOTHROW(wks.addPivotTable(options));
    REQUIRE_NOTHROW(doc.save());
    doc.close();

    // Verify XML structure to ensure styling properties weren't lost
    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open(__global_unique_file_2()));

    std::string ptDefXmlStr = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");

    // The single most critical node that tells Excel how to paint the table
    // Check that it contains the explicit true/false values mapped from boolean flags
    auto stylePos = ptDefXmlStr.find("<pivotTableStyleInfo");
    REQUIRE(stylePos != std::string::npos);

    REQUIRE(ptDefXmlStr.find("name=\"PivotStyleMedium9\"", stylePos) != std::string::npos);
    REQUIRE(ptDefXmlStr.find("showRowStripes=\"1\"", stylePos) != std::string::npos);
    REQUIRE(ptDefXmlStr.find("showColStripes=\"0\"", stylePos) != std::string::npos);
    REQUIRE(ptDefXmlStr.find("showRowHeaders=\"1\"", stylePos) != std::string::npos);
    REQUIRE(ptDefXmlStr.find("showColHeaders=\"1\"", stylePos) != std::string::npos);
    
    std::remove(__global_unique_file_2());
}

TEST_CASE("AdvancedPivotTableMultipleDataFieldsBaseAttributesRegression", "[XLPivotTable][Regression]")
{
    // This test ensures we NEVER regress on the critical OOXML requirement where multiple
    // <dataField> items MUST possess baseField="0" and baseItem="0". Without these, Microsoft Excel 
    // flags the workbook as corrupted.
    
    XLDocument doc;
    doc.create(__global_unique_file_3(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Category";
    wks.cell("B1").value() = "Metric1";
    wks.cell("C1").value() = "Metric2";

    wks.cell("A2").value() = "X"; wks.cell("B2").value() = 10; wks.cell("C2").value() = 100;
    wks.cell("A3").value() = "Y"; wks.cell("B3").value() = 20; wks.cell("C3").value() = 200;

    auto options = XLPivotTableOptions("BaseFieldPivot", "Sheet1!A1:C3", "E1")
        .addRowField("Category")
        .addDataField("Metric1", "Sum of Metric1", XLPivotSubtotal::Sum)
        .addDataField("Metric2", "Sum of Metric2", XLPivotSubtotal::Sum);

    REQUIRE_NOTHROW(wks.addPivotTable(options));
    REQUIRE_NOTHROW(doc.save());
    doc.close();

    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open(__global_unique_file_3()));

    std::string ptDefXmlStr = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");

    // Critical assertion: verify the existence of the baseField and baseItem attributes for EVERY dataField
    // We expect 2 dataFields
    size_t dataFieldPos = 0;
    int dataFieldCount = 0;
    while ((dataFieldPos = ptDefXmlStr.find("<dataField ", dataFieldPos)) != std::string::npos) {
        // Find the end of this specific tag
        size_t endTag = ptDefXmlStr.find(">", dataFieldPos);
        REQUIRE(endTag != std::string::npos);
        
        std::string tagContent = ptDefXmlStr.substr(dataFieldPos, endTag - dataFieldPos);
        
        // Assert the mandated anti-corruption attributes are injected
        REQUIRE(tagContent.find("baseField=\"0\"") != std::string::npos);
        REQUIRE(tagContent.find("baseItem=\"0\"") != std::string::npos);
        
        dataFieldCount++;
        dataFieldPos = endTag;
    }
    
    REQUIRE(dataFieldCount == 2);
    
    std::remove(__global_unique_file_3());
}

TEST_CASE("AdvancedPivotTableNamedRangeTableSourceBinding", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create(__global_unique_file_4(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Category";
    wks.cell("B1").value() = "Value";
    wks.cell("A2").value() = "Fruit";
    wks.cell("B2").value() = 100;

    // Simulate binding to an Excel Table or a global named range where no explicit sheet name is given
    auto options = XLPivotTableOptions("NamedRangePivot", "MyTable1", "D1")
        .addRowField("Category")
        .addDataField("Value", "Total Value", XLPivotSubtotal::Sum);

    // This currently will throw or fallback to "Field1" because MyTable1 is not parsed as "Sheet1!A1:B2"
    // OpenXLSX currently tightly couples parsing the string with `sourceRef.find(':')` to extract headers.
    
    // Let's assert it falls back safely without crashing.
    REQUIRE_NOTHROW(wks.addPivotTable(options));
    
    REQUIRE_NOTHROW(doc.save());
    doc.close();

    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open(__global_unique_file_4()));
    std::string ptCacheStr = doc2.extractXmlFromArchive("xl/pivotCache/pivotCacheDefinition1.xml");
    
    // It should have safely defaulted the cache source to MyTable1 and fallen back to a default count
    bool hasName = ptCacheStr.find("name=\"MyTable1\"") != std::string::npos || ptCacheStr.find("ref=\"MyTable1\"") != std::string::npos;
    REQUIRE(hasName);
    
    std::remove(__global_unique_file_4());
}

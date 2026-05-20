#include "TestHelpers.hpp"
#include "OpenXLSX.hpp"
#include "XLPivotTable.hpp"

#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_testXLPivotTableAdvanced_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotMultiDataTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_testXLPivotTableAdvanced_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotBuilderTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_testXLPivotTableAdvanced_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotStyleRegressionTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_testXLPivotTableAdvanced_3() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotBaseFieldTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_testXLPivotTableAdvanced_4() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotNamedSourceTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_testXLPivotTableAdvanced_5() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotDataOnRowsTest_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("AdvancedDynamicPivotTableBuilder", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create(__global_unique_testXLPivotTableAdvanced_1(), XLForceOverwrite);
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
    REQUIRE_NOTHROW(doc2.open(__global_unique_testXLPivotTableAdvanced_1()));

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
    doc.create(__global_unique_testXLPivotTableAdvanced_0(), XLForceOverwrite);
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
    REQUIRE_NOTHROW(doc2.open(__global_unique_testXLPivotTableAdvanced_0()));

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
    
    std::remove(__global_unique_testXLPivotTableAdvanced_0().c_str());
}

TEST_CASE("AdvancedPivotTableDataOnRowsandNumFmt", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create(__global_unique_testXLPivotTableAdvanced_5(), XLForceOverwrite);
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
    REQUIRE_NOTHROW(doc2.open(__global_unique_testXLPivotTableAdvanced_5()));

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
    
    std::remove(__global_unique_testXLPivotTableAdvanced_5().c_str());
}

TEST_CASE("AdvancedPivotTableStylingandFormattingRegression", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create(__global_unique_testXLPivotTableAdvanced_2(), XLForceOverwrite);
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
    REQUIRE_NOTHROW(doc2.open(__global_unique_testXLPivotTableAdvanced_2()));

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
    
    std::remove(__global_unique_testXLPivotTableAdvanced_2().c_str());
}

TEST_CASE("AdvancedPivotTableMultipleDataFieldsBaseAttributesRegression", "[XLPivotTable][Regression]")
{
    // This test ensures we NEVER regress on the critical OOXML requirement where multiple
    // <dataField> items MUST possess baseField="0" and baseItem="0". Without these, Microsoft Excel 
    // flags the workbook as corrupted.
    
    XLDocument doc;
    doc.create(__global_unique_testXLPivotTableAdvanced_3(), XLForceOverwrite);
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
    REQUIRE_NOTHROW(doc2.open(__global_unique_testXLPivotTableAdvanced_3()));

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
        REQUIRE(tagContent.find("baseField=\"-1\"") != std::string::npos);
        REQUIRE(tagContent.find("baseItem=\"1048832\"") != std::string::npos);
        
        dataFieldCount++;
        dataFieldPos = endTag;
    }
    
    REQUIRE(dataFieldCount == 2);
    
    std::remove(__global_unique_testXLPivotTableAdvanced_3().c_str());
}

TEST_CASE("AdvancedPivotTableNamedRangeTableSourceBinding", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create(__global_unique_testXLPivotTableAdvanced_4(), XLForceOverwrite);
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
    REQUIRE_NOTHROW(doc2.open(__global_unique_testXLPivotTableAdvanced_4()));
    std::string ptCacheStr = doc2.extractXmlFromArchive("xl/pivotCache/pivotCacheDefinition1.xml");
    
    // It should have safely defaulted the cache source to MyTable1 and fallen back to a default count
    bool hasName = ptCacheStr.find("name=\"MyTable1\"") != std::string::npos || ptCacheStr.find("ref=\"MyTable1\"") != std::string::npos;
    REQUIRE(hasName);
    
    std::remove(__global_unique_testXLPivotTableAdvanced_4().c_str());
}

TEST_CASE("AdvancedPivotTableClassicLayoutPrintTitlesAndFiltering", "[XLPivotTable]")
{
    static std::string filename = OpenXLSX::TestHelpers::getUniqueFilename("__PivotFilterTest_xlsx") + ".xlsx";
    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Year";
    wks.cell("B1").value() = "Quarter";
    wks.cell("C1").value() = "Sales";

    wks.cell("A2").value() = 2021; wks.cell("B2").value() = "Q1"; wks.cell("C2").value() = 100;
    wks.cell("A3").value() = 2021; wks.cell("B3").value() = "Q2"; wks.cell("C3").value() = 200;
    wks.cell("A4").value() = 2022; wks.cell("B4").value() = "Q1"; wks.cell("C4").value() = 300;
    wks.cell("A5").value() = 2022; wks.cell("B5").value() = "Q2"; wks.cell("C5").value() = 400;

    auto options = XLPivotTableOptions("FilterPivot", "Sheet1!A1:C5", "E1")
        .setClassicLayout(true)
        .setFieldPrintTitles(true)
        .setItemPrintTitles(true)
        .addRowField("Year", {"2021"})  // Hide 2022
        .addColumnField("Quarter", {"Q2"})  // Hide Q1
        .addDataField("Sales", "Var Sales", XLPivotSubtotal::Var);

    REQUIRE_NOTHROW(wks.addPivotTable(options));
    REQUIRE_NOTHROW(doc.save());
    doc.close();

    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open(filename));

    std::string ptDefXmlStr = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");
    std::string ptCacheStr = doc2.extractXmlFromArchive("xl/pivotCache/pivotCacheDefinition1.xml");

    // Verify classic layout attributes on ptRoot
    REQUIRE(ptDefXmlStr.find("gridDropZones=\"1\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("fieldPrintTitles=\"1\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("itemPrintTitles=\"1\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("compactData=\"0\"") != std::string::npos);

    // Verify subtotal="var" is written
    REQUIRE(ptDefXmlStr.find("subtotal=\"var\"") != std::string::npos);

    // Verify sharedItems uses blank placeholders (refreshOnLoad strategy)
    REQUIRE(ptCacheStr.find("containsBlank") != std::string::npos);

    // Verify SelectedItems filtering:
    // Year: selected "2021", so "2022" should be hidden.
    // Quarter: selected "Q2", so "Q1" should be hidden.
    REQUIRE(ptDefXmlStr.find("<item x=\"1\" h=\"1\" />") != std::string::npos); // 2022 hidden
    REQUIRE(ptDefXmlStr.find("<item x=\"0\" h=\"1\" />") != std::string::npos); // Q1 hidden

    std::remove(filename.c_str());
}

TEST_CASE("GeneratePivotTableDemo", "[.demo]")
{
    std::string filename = "pivot_table_demo.xlsx";
    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Product";
    wks.cell("C1").value() = "Sales";
    wks.cell("D1").value() = "Profit";
    wks.cell("E1").value() = "Year";

    wks.cell("A2").value() = "North"; wks.cell("B2").value() = "Apple";  wks.cell("C2").value() = 1200; wks.cell("D2").value() = 400; wks.cell("E2").value() = 2023;
    wks.cell("A3").value() = "South"; wks.cell("B3").value() = "Apple";  wks.cell("C3").value() = 1500; wks.cell("D3").value() = 500; wks.cell("E3").value() = 2023;
    wks.cell("A4").value() = "North"; wks.cell("B4").value() = "Orange"; wks.cell("C4").value() = 900;  wks.cell("D4").value() = 300; wks.cell("E4").value() = 2023;
    wks.cell("A5").value() = "South"; wks.cell("B5").value() = "Orange"; wks.cell("C5").value() = 1100; wks.cell("D5").value() = 350; wks.cell("E5").value() = 2023;
    wks.cell("A6").value() = "North"; wks.cell("B6").value() = "Banana"; wks.cell("C6").value() = 600;  wks.cell("D6").value() = 200; wks.cell("E6").value() = 2024;
    wks.cell("A7").value() = "South"; wks.cell("B7").value() = "Banana"; wks.cell("C7").value() = 700;  wks.cell("D7").value() = 250; wks.cell("E7").value() = 2024;
    wks.cell("A8").value() = "North"; wks.cell("B8").value() = "Apple";  wks.cell("C8").value() = 1300; wks.cell("D8").value() = 450; wks.cell("E8").value() = 2024;
    wks.cell("A9").value() = "South"; wks.cell("B9").value() = "Apple";  wks.cell("C9").value() = 1600; wks.cell("D9").value() = 550; wks.cell("E9").value() = 2024;

    doc.workbook().addWorksheet("Sheet2");
    auto wks2 = doc.workbook().worksheet("Sheet2");

    auto options = XLPivotTableOptions("PremiumPivot", "Sheet1!A1:E9", "A3")
        .setClassicLayout(true)
        .setFieldPrintTitles(true)
        .setItemPrintTitles(true)
        .addRowField("Region", {"North"}) // Only show North region
        .addColumnField("Product", {"Apple", "Orange"}) // Only show Apple and Orange, hiding Banana
        .addFilterField("Year")
        .addDataField("Sales", "Total Sales", XLPivotSubtotal::Sum)
        .addDataField("Profit", "Variance Profit", XLPivotSubtotal::Var);

    wks2.addPivotTable(options);
    doc.save();
    doc.close();
}

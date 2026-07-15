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
        
        // Assert the correct baseField/baseItem attributes (0/0 matches Excel's repaired format)
        REQUIRE(tagContent.find("baseField=\"0\"") != std::string::npos);
        REQUIRE(tagContent.find("baseItem=\"0\"") != std::string::npos);
        
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
    wks.cell("A3").value() = "Veg";
    wks.cell("B3").value() = 50;

    // Real Excel Table as pivot source (worksheetSource/@name).
    REQUIRE_NOTHROW(wks.tables().add("MyTable1", "A1:B3"));

    auto options = XLPivotTableOptions("NamedRangePivot", "MyTable1", "D1")
        .addRowField("Category")
        .addDataField("Value", "Total Value", XLPivotSubtotal::Sum);

    REQUIRE_NOTHROW(wks.addPivotTable(options));
    REQUIRE_NOTHROW(doc.save());
    doc.close();

    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open(__global_unique_testXLPivotTableAdvanced_4()));
    std::string ptCacheStr = doc2.extractXmlFromArchive("xl/pivotCache/pivotCacheDefinition1.xml");
    std::string ptDefXmlStr = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");

    REQUIRE(ptCacheStr.find("name=\"MyTable1\"") != std::string::npos);
    // Headers must be resolved from the table range (not Field1 fallback).
    REQUIRE(ptCacheStr.find("name=\"Category\"") != std::string::npos);
    REQUIRE(ptCacheStr.find("name=\"Value\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("axis=\"axisRow\"") != std::string::npos);

    // Unknown table / name must hard-fail (no silent Field1).
    XLDocument doc3;
    doc3.create(__global_unique_testXLPivotTableAdvanced_4() + ".bad.xlsx", XLForceOverwrite);
    auto w3 = doc3.workbook().worksheet("Sheet1");
    w3.cell("A1").value() = "H";
    w3.cell("B1").value() = "V";
    w3.cell("A2").value() = "x";
    w3.cell("B2").value() = 1;
    REQUIRE_THROWS_AS(
        w3.addPivotTable(XLPivotTableOptions("Bad", "NoSuchTable", "D1").addRowField("H").addDataField("V")),
        OpenXLSX::XLInputError);
    doc3.close();
    std::remove((__global_unique_testXLPivotTableAdvanced_4() + ".bad.xlsx").c_str());

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

    // Verify sharedItems has actual distinct values (containsString/Number present)
    bool hasCacheData = ptCacheStr.find("containsString") != std::string::npos ||
                        ptCacheStr.find("containsNumber") != std::string::npos;
    REQUIRE(hasCacheData);

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

TEST_CASE("PivotTableShowValuesAsAndFieldLayout", "[XLPivotTable][P1]")
{
    static std::string filename = OpenXLSX::TestHelpers::getUniqueFilename("__PivotShowValuesAs_xlsx") + ".xlsx";
    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Month";
    wks.cell("C1").value() = "Revenue";
    wks.cell("A2").value() = "East";
    wks.cell("B2").value() = "Jan";
    wks.cell("C2").value() = 100;
    wks.cell("A3").value() = "West";
    wks.cell("B3").value() = "Jan";
    wks.cell("C3").value() = 200;
    wks.cell("A4").value() = "East";
    wks.cell("B4").value() = "Feb";
    wks.cell("C4").value() = 150;

    XLPivotField rowFld;
    rowFld.name = "Region";
    rowFld.compact = true;
    rowFld.outline = true;
    rowFld.insertBlankRow = true;

    XLPivotShowValuesAsOptions sva;
    sva.type = XLPivotShowValuesAs::PercentOf;
    sva.baseField = "Region";
    sva.baseItem = "East";

    auto options = XLPivotTableOptions("SVAPivot", "Sheet1!A1:C4", "E1")
                       .addRowField(rowFld)
                       .addColumnField("Month")
                       .addDataField("Revenue", "Pct of East", XLPivotSubtotal::Sum, 0, sva)
                       .setLocationRange("E1:K20");

    REQUIRE_NOTHROW(wks.addPivotTable(options));
    REQUIRE_NOTHROW(doc.save());
    doc.close();

    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open(filename));
    std::string ptXml = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");

    REQUIRE(ptXml.find("showDataAs=\"percent\"") != std::string::npos);
    REQUIRE(ptXml.find("baseField=\"0\"") != std::string::npos); // Region is field 0
    REQUIRE(ptXml.find("insertBlankRow=\"1\"") != std::string::npos);
    REQUIRE(ptXml.find("compact=\"1\"") != std::string::npos);
    REQUIRE(ptXml.find("ref=\"E1:K20\"") != std::string::npos);

    // options() round-trip (best effort)
    auto pts = doc2.workbook().worksheet("Sheet1").pivotTables();
    REQUIRE(pts.size() == 1);
    auto recovered = pts[0].options();
    REQUIRE(recovered.name() == "SVAPivot");
    REQUIRE(recovered.rows().size() == 1);
    REQUIRE(recovered.rows()[0].name == "Region");
    REQUIRE(recovered.data().size() == 1);
    REQUIRE(recovered.data()[0].showValuesAs.type == XLPivotShowValuesAs::PercentOf);
    REQUIRE(recovered.data()[0].showValuesAs.baseField == "Region");
    REQUIRE(recovered.locationRange() == "E1:K20");

    // x14 ShowValuesAs
    XLDocument doc3;
    doc3.create(filename + ".x14.xlsx", XLForceOverwrite);
    auto w3 = doc3.workbook().worksheet("Sheet1");
    w3.cell("A1").value() = "Year";
    w3.cell("B1").value() = "Sales";
    w3.cell("A2").value() = 2021;
    w3.cell("B2").value() = 10;
    w3.cell("A3").value() = 2022;
    w3.cell("B3").value() = 20;
    XLPivotShowValuesAsOptions sva2;
    sva2.type = XLPivotShowValuesAs::RunningTotalIn;
    sva2.baseField = "Year";
    // Running total needs base field only; ensure Year has deep shared items via selected empty but we need base field deep:
    // force deep by using base field requirement — deepSharedFields includes baseField.
    REQUIRE_NOTHROW(w3.addPivotTable(XLPivotTableOptions("RunTotal", "Sheet1!A1:B3", "D1")
                                         .addRowField("Year")
                                         .addDataField("Sales", "RT", XLPivotSubtotal::Sum, 0, sva2)));
    doc3.save();
    std::string pt3 = doc3.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");
    // RunningTotalIn is classic showDataAs runTotal (not x14)
    REQUIRE(pt3.find("showDataAs=\"runTotal\"") != std::string::npos);
    doc3.close();
    std::remove((filename + ".x14.xlsx").c_str());

    // x14 Index
    XLDocument doc4;
    doc4.create(filename + ".idx.xlsx", XLForceOverwrite);
    auto w4 = doc4.workbook().worksheet("Sheet1");
    w4.cell("A1").value() = "Cat";
    w4.cell("B1").value() = "Val";
    w4.cell("A2").value() = "A";
    w4.cell("B2").value() = 1;
    XLPivotShowValuesAsOptions sva3;
    sva3.type = XLPivotShowValuesAs::Index;
    REQUIRE_NOTHROW(w4.addPivotTable(XLPivotTableOptions("Idx", "Sheet1!A1:B2", "D1")
                                         .addRowField("Cat")
                                         .addDataField("Val", "I", XLPivotSubtotal::Sum, 0, sva3)));
    doc4.save();
    std::string pt4 = doc4.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");
    REQUIRE(pt4.find("pivotShowAs=\"index\"") != std::string::npos);
    doc4.close();
    std::remove((filename + ".idx.xlsx").c_str());

    std::remove(filename.c_str());
}

TEST_CASE("PivotTableShowValuesAsValidation", "[XLPivotTable][P1]")
{
    XLDocument doc;
    std::string filename = OpenXLSX::TestHelpers::getUniqueFilename("__PivotSVAVal_xlsx") + ".xlsx";
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    wks.cell("A1").value() = "R";
    wks.cell("B1").value() = "V";
    wks.cell("A2").value() = "x";
    wks.cell("B2").value() = 1;

    XLPivotShowValuesAsOptions sva;
    sva.type = XLPivotShowValuesAs::PercentOf;
    // missing baseField
    REQUIRE_THROWS_AS(wks.addPivotTable(XLPivotTableOptions("T", "Sheet1!A1:B2", "D1")
                                            .addRowField("R")
                                            .addDataField("V", "P", XLPivotSubtotal::Sum, 0, sva)),
                      XLInputError);
    sva.baseField = "R";
    // missing baseItem
    REQUIRE_THROWS_AS(wks.addPivotTable(XLPivotTableOptions("T", "Sheet1!A1:B2", "D1")
                                            .addRowField("R")
                                            .addDataField("V", "P", XLPivotSubtotal::Sum, 0, sva)),
                      XLInputError);
    doc.close();
    std::remove(filename.c_str());
}

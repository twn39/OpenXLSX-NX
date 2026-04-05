#include "OpenXLSX.hpp"
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Advanced Dynamic Pivot Table Builder", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create("./PivotBuilderTest.xlsx", XLForceOverwrite);
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

    XLPivotTableOptions options;
    options.name        = "AdvancedPivot";
    options.sourceRange = "Sheet1!A1:D4";
    options.targetCell  = "F1";

    // Adding dynamic fields
    options.filters.push_back({"Date", XLPivotSubtotal::Sum, ""}); // Page Field
    options.rows.push_back({"Category", XLPivotSubtotal::Sum, ""}); // Row Field
    options.columns.push_back({"Product", XLPivotSubtotal::Sum, ""}); // Col Field
    options.data.push_back({"Sales", XLPivotSubtotal::Average, "Average Sales"}); // Data Field with Average

    // Configure layout and styles based on our new Builder flags
    options.rowGrandTotals = false;
    options.colGrandTotals = false;
    options.compactData = false;
    options.showRowStripes = true;
    options.pivotTableStyleName = "PivotStyleMedium9";

    REQUIRE_NOTHROW(wks.addPivotTable(options));
    REQUIRE_NOTHROW(doc.save());

    // Verify written values
    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open("./PivotBuilderTest.xlsx"));

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

TEST_CASE("Advanced Pivot Table Multi-Data Fields and Edge Cases", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create("./PivotMultiDataTest.xlsx", XLForceOverwrite);
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

    XLPivotTableOptions options;
    options.name        = "MultiDataPivot";
    options.sourceRange = "Sheet1!A1:E4";
    options.targetCell  = "G1";

    // Adding dynamic fields
    options.rows.push_back({"Category", XLPivotSubtotal::Sum, ""}); 
    
    // Injecting multiple data fields!
    // This forces Excel to create a virtual "\x{03A3} Values" field in the columns/rows area
    options.data.push_back({"Sales", XLPivotSubtotal::Sum, "Total Sales"});
    options.data.push_back({"Profit", XLPivotSubtotal::Average, "Avg Profit"});

    REQUIRE_NOTHROW(wks.addPivotTable(options));
    REQUIRE_NOTHROW(doc.save());
    doc.close();

    // Verify written values and XML structures for multi-data fields
    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open("./PivotMultiDataTest.xlsx"));

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
    
    std::remove("./PivotMultiDataTest.xlsx");
}

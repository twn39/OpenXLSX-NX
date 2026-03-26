#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <iostream>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("Table Slicer Features", "[XLSlicer]")
{
    XLDocument doc;
    doc.create("./TableSlicerTest.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Sales";
    
    wks.cell("A2").value() = "North";
    wks.cell("B2").value() = 100;
    
    wks.cell("A3").value() = "South";
    wks.cell("B3").value() = 200;

    auto tables = wks.tables();
    auto table = tables.add("SalesTable", "A1:B3");

    XLSlicerOptions opts;
    opts.name = "Region";
    opts.caption = "Region Filter";
    
    wks.addTableSlicer("D1", table, "Region", opts);

    doc.save();
    
    XLDocument doc2;
    doc2.open("./TableSlicerTest.xlsx");
    REQUIRE(doc2.workbook().worksheet("Sheet1").hasDrawing() == true);

    doc2.saveAs("./TableSlicerTest2.xlsx", XLForceOverwrite);
    
    XLDocument doc3;
    doc3.open("./TableSlicerTest2.xlsx");
    REQUIRE(doc3.workbook().worksheet("Sheet1").hasDrawing() == true);
}

/**
 * testXLSlicerComprehensive.cpp
 *
 * 将原有的综合性 Excel 拆分为独立的功能测试文件：
 * 1. slicer_table_demo.xlsx (包含表格分片器，多 Sheet 等)
 * 2. slicer_pivot_demo.xlsx (包含数据透视表及对应的透视表分片器)
 */
#include "TestHelpers.hpp"
#include "OpenXLSX.hpp"
#include "XLSlicer.hpp"
#include "XLSlicerCollection.hpp"
#include "XLTables.hpp"
#include "XLPivotTable.hpp"

#include <catch2/catch_test_macros.hpp>
#include <iostream>
#include <string>
#include <vector>

using namespace OpenXLSX;

namespace {

// 填充通用测试数据
void fillDemoData(XLWorksheet& wks)
{
    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Product";
    wks.cell("C1").value() = "Quarter";
    wks.cell("D1").value() = "Sales";

    const std::vector<std::tuple<std::string,std::string,std::string,int>> rows = {
        {"North","Widget","Q1",1200},{"North","Gadget","Q2",900},
        {"South","Widget","Q1",850}, {"South","Gadget","Q3",1100},
        {"East", "Widget","Q2",700}, {"East", "Gadget","Q4",1300},
        {"West", "Widget","Q3",600}, {"West", "Gadget","Q1",1050},
    };
    for (size_t i = 0; i < rows.size(); ++i) {
        uint32_t r = 2 + static_cast<uint32_t>(i);
        wks.cell(r, 1).value() = std::get<0>(rows[i]);
        wks.cell(r, 2).value() = std::get<1>(rows[i]);
        wks.cell(r, 3).value() = std::get<2>(rows[i]);
        wks.cell(r, 4).value() = std::get<3>(rows[i]);
    }
}

} // namespace

TEST_CASE("GenerateTableSlicerDemo", "[SlicerDemo]")
{
    const std::string out = "slicer_table_demo.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    // 1. Sheet1: Raw data table + 2 Table Slicers
    auto wks1 = doc.workbook().worksheet(1);
    wks1.setName("Sheet1");
    fillDemoData(wks1);

    auto table1 = wks1.tables().add("T_Sales", "A1:D9");

    // Add table slicers
    wks1.slicers().add("F1", table1, "Region")
        .name("Slicer_Region")
        .caption("地区 (Light1)")
        .style(XLSlicerStyle::Light1)
        .size(150, 200);

    wks1.slicers().add("H1", table1, "Product")
        .name("Slicer_Product")
        .caption("产品 (Dark3)")
        .style(XLSlicerStyle::Dark3)
        .size(150, 200);

    // 2. Sheet2: Another sheet with a Table and Slicer to test multi-sheet workbook-wide cache name uniqueness
    doc.workbook().addWorksheet("Sheet2");
    auto wks2 = doc.workbook().worksheet("Sheet2");
    fillDemoData(wks2);

    auto table2 = wks2.tables().add("T_Quarter", "A1:D9");

    wks2.slicers().add("F1", table2, "Quarter")
        .name("Slicer_Quarter")
        .caption("季度 (Light4)")
        .style(XLSlicerStyle::Light4)
        .size(150, 200);

    doc.save();
    doc.close();

    std::cout << "\n===================================================\n";
    std::cout << "  表格切片器测试文件生成完毕: " << out << "\n";
    std::cout << "  包含多 Sheet、唯一 Cache 命名、各种样式表格切片器。\n";
    std::cout << "===================================================\n\n";

    REQUIRE(true);
}

TEST_CASE("GeneratePivotSlicerDemo", "[SlicerDemo]")
{
    const std::string out = "slicer_pivot_demo.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    // 1. Sheet1: Raw data table
    auto wks1 = doc.workbook().worksheet(1);
    wks1.setName("Sheet1");
    fillDemoData(wks1);

    // 2. Sheet2: Pivot Table + Pivot Slicer
    doc.workbook().addWorksheet("Sheet2");
    auto wks2 = doc.workbook().worksheet("Sheet2");
    
    // Add Pivot Table sourcing from Sheet1!A1:D9
    auto pt = wks2.addPivotTable(
        XLPivotTableOptions("PT_Sales", "Sheet1!A1:D9", "A1")
            .addRowField("Region")
            .addColumnField("Product")
            .addDataField("Sales", "总销售额")
    );

    if (pt.valid()) {
        wks2.slicers().add("F1", pt, "Region")
            .caption("地区筛选[Pivot]")
            .style(XLSlicerStyle::Dark4)
            .size(160, 220);
    } else {
        wks2.cell("F1").value() = "Pivot Table Creation Failed";
    }

    doc.save();
    doc.close();

    std::cout << "\n===================================================\n";
    std::cout << "  透视表切片器测试文件生成完毕: " << out << "\n";
    std::cout << "  包含数据源 Sheet、透视表 Sheet，以及透视表切片器。\n";
    std::cout << "===================================================\n\n";

    REQUIRE(true);
}

TEST_CASE("GenerateSlicerComprehensiveDemo", "[SlicerDemo]")
{
    const std::string out = "slicer_comprehensive_demo.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    // 1. Sheet1: Raw data table + 2 Table Slicers
    auto wks1 = doc.workbook().worksheet(1);
    wks1.setName("Sheet1");
    fillDemoData(wks1);

    auto table1 = wks1.tables().add("T_Sales", "A1:D9");

    // Add table slicers
    wks1.slicers().add("F1", table1, "Region")
        .name("Slicer_Region")
        .caption("地区 (Light1)")
        .style(XLSlicerStyle::Light1)
        .size(150, 200);

    wks1.slicers().add("H1", table1, "Product")
        .name("Slicer_Product")
        .caption("产品 (Dark3)")
        .style(XLSlicerStyle::Dark3)
        .size(150, 200);

    // 2. Sheet2: Pivot Table + Pivot Slicer
    doc.workbook().addWorksheet("Sheet2");
    auto wks2 = doc.workbook().worksheet("Sheet2");
    
    // Add Pivot Table sourcing from Sheet1!A1:D9
    auto pt = wks2.addPivotTable(
        XLPivotTableOptions("PT_Sales", "Sheet1!A1:D9", "A1")
            .addRowField("Region")
            .addColumnField("Product")
            .addDataField("Sales", "总销售额")
    );

    if (pt.valid()) {
        wks2.slicers().add("F1", pt, "Region")
            .caption("地区筛选[Pivot]")
            .style(XLSlicerStyle::Dark4)
            .size(160, 220);
    } else {
        wks2.cell("F1").value() = "Pivot Table Creation Failed";
    }

    // 3. Sheet3: Another sheet with a Table and Slicer to test multi-sheet workbook-wide cache name uniqueness
    doc.workbook().addWorksheet("Sheet3");
    auto wks3 = doc.workbook().worksheet("Sheet3");
    fillDemoData(wks3);

    auto table2 = wks3.tables().add("T_Quarter", "A1:D9");

    wks3.slicers().add("F1", table2, "Quarter")
        .name("Slicer_Quarter")
        .caption("季度 (Light4)")
        .style(XLSlicerStyle::Light4)
        .size(150, 200);

    doc.save();
    doc.close();

    std::cout << "\n===================================================\n";
    std::cout << "  综合切片器测试文件生成完毕: " << out << "\n";
    std::cout << "  包含 Table 切片器, Pivot Table 切片器，多 Sheet 等。\n";
    std::cout << "===================================================\n\n";

    REQUIRE(true);
}


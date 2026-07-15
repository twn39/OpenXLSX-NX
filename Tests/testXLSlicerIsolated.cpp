/**
 * testXLSlicerIsolated.cpp
 *
 * 将切片器功能拆分为多个独立 xlsx 文件，逐一人工校验哪个损坏。
 * 运行: ./OpenXLSXTests "[SlicerIsolated]" --reporter compact
 * 输出: build/ 目录下生成如下文件:
 *   slicer_iso_01_minimal.xlsx       — 最简单: 1张表 + 1个切片器
 *   slicer_iso_02_two_slicers.xlsx   — 1张表 + 2个切片器
 *   slicer_iso_03_style.xlsx         — 切片器带不同样式
 *   slicer_iso_04_prefilter.xlsx     — 切片器预筛选项目
 *   slicer_iso_05_multi_sheet.xlsx   — 多 Sheet 各自有切片器
 *   slicer_iso_06_pivot.xlsx         — Pivot Table 切片器
 *   slicer_iso_07_delete.xlsx        — 创建后删除
 *   slicer_iso_08_modify.xlsx        — 创建后修改属性
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

// 填充通用测试数据（4列，8行数据）
void fillData(XLWorksheet& wks)
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

void printResult(const std::string& name, bool ok) {
    std::cout << (ok ? "  [OK] " : "  [!!] ") << name << "\n";
}

} // namespace

// ─── TEST 01: 最简单 — 1表 + 1切片器 ────────────────────────────────────────
TEST_CASE("SlicerIso01_Minimal", "[SlicerIsolated]")
{
    const std::string out = "slicer_iso_01_minimal.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    auto wks = doc.workbook().worksheet(1);
    wks.setName("Sheet1");
    fillData(wks);

    auto table = wks.tables().add("T1", "A1:D9");

    // 只加1个切片器，用最简单的方式
    wks.slicers().add("F1", table, "Region")
        .name("Slicer_Region")
        .caption("地区");

    doc.save();
    doc.close();
    std::cout << "\n[ISO-01] 1表 + 1切片器 → " << out << "\n";
    REQUIRE(true);
}

// ─── TEST 02: 1表 + 2个切片器 ────────────────────────────────────────────────
TEST_CASE("SlicerIso02_TwoSlicers", "[SlicerIsolated]")
{
    const std::string out = "slicer_iso_02_two_slicers.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    auto wks = doc.workbook().worksheet(1);
    wks.setName("Sheet1");
    fillData(wks);

    auto table = wks.tables().add("T1", "A1:D9");

    wks.slicers().add("F1", table, "Region")
        .name("Slicer_Region").caption("地区");
    wks.slicers().add("H1", table, "Product")
        .name("Slicer_Product").caption("产品");

    doc.save();
    doc.close();
    std::cout << "\n[ISO-02] 1表 + 2切片器 → " << out << "\n";
    REQUIRE(true);
}

// ─── TEST 03: 切片器带样式 ────────────────────────────────────────────────────
TEST_CASE("SlicerIso03_Style", "[SlicerIsolated]")
{
    const std::string out = "slicer_iso_03_style.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    auto wks = doc.workbook().worksheet(1);
    wks.setName("Sheet1");
    fillData(wks);

    auto table = wks.tables().add("T1", "A1:D9");

    wks.slicers().add("F1", table, "Region")
        .name("Slicer_Region").caption("地区").style(XLSlicerStyle::Light1).size(144, 200);
    wks.slicers().add("H1", table, "Product")
        .name("Slicer_Product").caption("产品").style(XLSlicerStyle::Dark3).size(160, 200);
    wks.slicers().add("J1", table, "Quarter")
        .name("Slicer_Quarter").caption("季度").style(XLSlicerStyle::Light3).size(144, 200);

    doc.save();
    doc.close();
    std::cout << "\n[ISO-03] 3种样式切片器 → " << out << "\n";
    REQUIRE(true);
}

// ─── TEST 04: 预筛选 (showOnly) ──────────────────────────────────────────────
TEST_CASE("SlicerIso04_PreFilter", "[SlicerIsolated]")
{
    const std::string out = "slicer_iso_04_prefilter.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    auto wks = doc.workbook().worksheet(1);
    wks.setName("Sheet1");
    fillData(wks);

    auto table = wks.tables().add("T1", "A1:D9");

    wks.slicers().add("F1", table, "Quarter")
        .name("Slicer_Quarter").caption("季度[预选Q1+Q2]")
        .style(XLSlicerStyle::Light1).size(144, 200)
        .showOnly({"Q1", "Q2"});

    doc.save();
    doc.close();
    std::cout << "\n[ISO-04] 预筛选切片器 → " << out << "\n";
    REQUIRE(true);
}

// ─── TEST 05: 多Sheet各有切片器 ─────────────────────────────────────────────
TEST_CASE("SlicerIso05_MultiSheet", "[SlicerIsolated]")
{
    const std::string out = "slicer_iso_05_multi_sheet.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    // Sheet1
    {
        auto wks = doc.workbook().worksheet(1);
        wks.setName("Sheet1");
        fillData(wks);
        auto table = wks.tables().add("T1", "A1:D9");
        wks.slicers().add("F1", table, "Region")
            .name("Slicer_Region").caption("地区").style(XLSlicerStyle::Light1);
    }

    // Sheet2
    {
        doc.workbook().addWorksheet("Sheet2");
        auto wks = doc.workbook().worksheet("Sheet2");
        fillData(wks);
        auto table = wks.tables().add("T2", "A1:D9");
        wks.slicers().add("F1", table, "Product")
            .name("Slicer_Product").caption("产品").style(XLSlicerStyle::Dark2);
    }

    // Sheet3
    {
        doc.workbook().addWorksheet("Sheet3");
        auto wks = doc.workbook().worksheet("Sheet3");
        fillData(wks);
        auto table = wks.tables().add("T3", "A1:D9");
        wks.slicers().add("F1", table, "Quarter")
            .name("Slicer_Quarter").caption("季度").style(XLSlicerStyle::Light4);
    }

    doc.save();
    doc.close();
    std::cout << "\n[ISO-05] 3个Sheet各1个切片器 → " << out << "\n";
    REQUIRE(true);
}

// ─── TEST 06: 透视表切片器 ───────────────────────────────────────────────────
TEST_CASE("SlicerIso06_Pivot", "[SlicerIsolated]")
{
    const std::string out = "slicer_iso_06_pivot.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    auto wks = doc.workbook().worksheet(1);
    wks.setName("Sheet1");
    fillData(wks);

    auto pt = wks.addPivotTable(
        XLPivotTableOptions("PT_Sales", "Sheet1!A1:D9", "F1")
            .addRowField("Region")
            .addDataField("Sales", "总销售额")
    );

    if (pt.valid()) {
        wks.slicers().add("M1", pt, "Region")
            .caption("地区筛选[Pivot]").style(XLSlicerStyle::Dark4).size(160, 220);
        std::cout << "\n[ISO-06] Pivot切片器 → " << out << "\n";
    } else {
        wks.cell("M1").value() = "Pivot 创建失败";
        std::cout << "\n[ISO-06] Pivot创建失败，跳过切片器 → " << out << "\n";
    }

    doc.save();
    doc.close();
    REQUIRE(true);
}

// ─── TEST 07: 创建后删除 ─────────────────────────────────────────────────────
TEST_CASE("SlicerIso07_Delete", "[SlicerIsolated]")
{
    const std::string out = "slicer_iso_07_delete.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    auto wks = doc.workbook().worksheet(1);
    wks.setName("Sheet1");
    fillData(wks);

    auto table = wks.tables().add("T1", "A1:D9");

    wks.slicers().add("F1", table, "Region")
        .name("Del_Region").caption("将被删除").style(XLSlicerStyle::Light1);
    wks.slicers().add("H1", table, "Product")
        .name("Keep_Product").caption("保留的切片器").style(XLSlicerStyle::Dark3);

    // 删除第一个
    wks.slicers().reload();
    REQUIRE(wks.slicers().count() == 2);
    wks.deleteSlicer("Del_Region");
    REQUIRE(wks.slicers().count() == 1);

    wks.cell("J1").value() = "F列切片器已删除，H列保留";

    doc.save();
    doc.close();
    std::cout << "\n[ISO-07] 创建2个删除1个 → " << out << "\n";
    REQUIRE(true);
}

// ─── TEST 08: 创建后修改属性 ─────────────────────────────────────────────────
TEST_CASE("SlicerIso08_Modify", "[SlicerIsolated]")
{
    const std::string out = "slicer_iso_08_modify.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    auto wks = doc.workbook().worksheet(1);
    wks.setName("Sheet1");
    fillData(wks);

    auto table = wks.tables().add("T1", "A1:D9");

    wks.slicers().add("F1", table, "Region")
        .name("Mod_Region").caption("原标题 Light1")
        .style(XLSlicerStyle::Light1).size(150, 220);

    // 修改属性
    wks.slicers().reload();
    auto& s = wks.slicers()["Mod_Region"];
    s.setCaption("修改后标题 Dark5 ✓").setStyle(XLSlicerStyle::Dark5);
    REQUIRE(s.caption() == "修改后标题 Dark5 ✓");
    REQUIRE(s.style() == XLSlicerStyle::Dark5);

    // 再加一个带columnCount
    wks.slicers().add("H1", table, "Product")
        .name("Mod_Product").caption("产品[2列]")
        .style(XLSlicerStyle::Light2).size(240, 130).columnCount(2);

    doc.save();
    doc.close();
    std::cout << "\n[ISO-08] 创建后修改属性 → " << out << "\n";
    REQUIRE(true);
}

// ─── 汇总报告 ─────────────────────────────────────────────────────────────────
TEST_CASE("SlicerIso_Summary", "[SlicerIsolated]")
{
    std::cout << "\n";
    std::cout << "═══════════════════════════════════════════════════\n";
    std::cout << "  切片器隔离测试文件生成完毕，请逐一在 Excel 打开:\n";
    std::cout << "═══════════════════════════════════════════════════\n";
    std::cout << "  slicer_iso_01_minimal.xlsx     — 1表+1切片器 (最简)\n";
    std::cout << "  slicer_iso_02_two_slicers.xlsx — 1表+2切片器\n";
    std::cout << "  slicer_iso_03_style.xlsx       — 切片器带样式\n";
    std::cout << "  slicer_iso_04_prefilter.xlsx   — 切片器预筛选\n";
    std::cout << "  slicer_iso_05_multi_sheet.xlsx — 多Sheet各有切片器\n";
    std::cout << "  slicer_iso_06_pivot.xlsx       — Pivot 切片器\n";
    std::cout << "  slicer_iso_07_delete.xlsx      — 创建后删除\n";
    std::cout << "  slicer_iso_08_modify.xlsx      — 创建后修改属性\n";
    std::cout << "═══════════════════════════════════════════════════\n";
    std::cout << "  若 01 损坏 → 切片器最基础结构有问题\n";
    std::cout << "  若 01 正常 02 损坏 → 多切片器共享 slicer.xml 有问题\n";
    std::cout << "  若 05 损坏 → 多 Sheet 切片器关系注册有问题\n";
    std::cout << "  若 06 损坏 → Pivot 切片器有问题\n";
    std::cout << "═══════════════════════════════════════════════════\n";
    REQUIRE(true);
}

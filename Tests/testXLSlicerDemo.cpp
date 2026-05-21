/**
 * testXLSlicerDemo.cpp
 *
 * 生成供人工校验的切片器综合演示 Excel 文件。
 * 运行: ./OpenXLSXTests "[SlicerDemo]" --reporter compact
 * 输出: slicer_comprehensive_demo.xlsx (在当前工作目录)
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

// ─── 填充通用数据 ─────────────────────────────────────────────────────────────
namespace {

void fillData(XLWorksheet& wks, uint32_t startRow = 1, uint16_t startCol = 1)
{
    wks.cell(startRow, startCol+0).value() = "Region";
    wks.cell(startRow, startCol+1).value() = "Product";
    wks.cell(startRow, startCol+2).value() = "Quarter";
    wks.cell(startRow, startCol+3).value() = "Sales";

    const std::vector<std::tuple<std::string,std::string,std::string,int>> rows = {
        {"North","Widget","Q1",1200},{"North","Gadget","Q2",900},
        {"South","Widget","Q1",850}, {"South","Gadget","Q3",1100},
        {"East", "Widget","Q2",700}, {"East", "Gadget","Q4",1300},
        {"West", "Widget","Q3",600}, {"West", "Gadget","Q1",1050},
    };
    for (size_t i = 0; i < rows.size(); ++i) {
        uint32_t r = startRow + 1 + static_cast<uint32_t>(i);
        wks.cell(r, startCol+0).value() = std::get<0>(rows[i]);
        wks.cell(r, startCol+1).value() = std::get<1>(rows[i]);
        wks.cell(r, startCol+2).value() = std::get<2>(rows[i]);
        wks.cell(r, startCol+3).value() = std::get<3>(rows[i]);
    }
}

} // namespace

// ─── 测试用例 ─────────────────────────────────────────────────────────────────
TEST_CASE("SlicerComprehensiveDemo", "[SlicerDemo]")
{
    const std::string outPath = "slicer_comprehensive_demo.xlsx";
    XLDocument doc;
    doc.create(outPath, XLForceOverwrite);

    // ─────────────────────────────────────────────────────────────────────────
    // Sheet 1: 基础 Table 切片器
    //   验证点: 3 个切片器放置于 F1/H1/J1，样式各异，尺寸精确
    // ─────────────────────────────────────────────────────────────────────────
    {
        auto wks = doc.workbook().worksheet(1);
        wks.setName("Sheet1_基础切片器");
        fillData(wks);

        auto table = wks.tables().add("T1_Data", "A1:D9");

        // ① Region — Light1, 默认尺寸 144×200
        wks.slicers().add("F1", table, "Region")
            .name("Slicer_Region")
            .caption("地区 (Light1)")
            .style(XLSlicerStyle::Light1)
            .size(144, 200);

        // ② Product — Dark3, 较宽
        wks.slicers().add("H1", table, "Product")
            .name("Slicer_Product")
            .caption("产品 (Dark3)")
            .style(XLSlicerStyle::Dark3)
            .size(160, 200);

        // ③ Quarter — Light3, 预筛选 Q1/Q2
        wks.slicers().add("J1", table, "Quarter")
            .name("Slicer_Quarter")
            .caption("季度 [预选 Q1+Q2]")
            .style(XLSlicerStyle::Light3)
            .size(144, 200)
            .showOnly({"Q1", "Q2"});

        std::cout << "[Sheet1] 3 个基础切片器 (Light1/Dark3/Light3)\n";
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Sheet 2: Light 系列 6 种样式
    //   验证点: Light1~Light6 外观有明显差异
    // ─────────────────────────────────────────────────────────────────────────
    {
        doc.workbook().addWorksheet("Sheet2_Light样式");
        auto wks = doc.workbook().worksheet("Sheet2_Light样式");
        fillData(wks);
        auto table = wks.tables().add("T2_Data", "A1:D9");

        const XLSlicerStyle lights[] = {
            XLSlicerStyle::Light1, XLSlicerStyle::Light2, XLSlicerStyle::Light3,
            XLSlicerStyle::Light4, XLSlicerStyle::Light5, XLSlicerStyle::Light6,
        };
        for (int i = 0; i < 6; ++i) {
            // F1, H1, J1, L1, N1, P1
            uint16_t col = static_cast<uint16_t>(6 + i * 2);
            XLCellReference ref(1u, col);
            std::string nm = "L" + std::to_string(i+1) + "_Region";
            wks.slicers().add(ref.address(), table, "Region")
                .name(nm)
                .caption("Light" + std::to_string(i+1))
                .style(lights[i])
                .size(120, 160);
        }
        std::cout << "[Sheet2] Light1~Light6 共 6 种切片器样式\n";
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Sheet 3: Dark 系列 6 种样式
    // ─────────────────────────────────────────────────────────────────────────
    {
        doc.workbook().addWorksheet("Sheet3_Dark样式");
        auto wks = doc.workbook().worksheet("Sheet3_Dark样式");
        fillData(wks);
        auto table = wks.tables().add("T3_Data", "A1:D9");

        const XLSlicerStyle darks[] = {
            XLSlicerStyle::Dark1, XLSlicerStyle::Dark2, XLSlicerStyle::Dark3,
            XLSlicerStyle::Dark4, XLSlicerStyle::Dark5, XLSlicerStyle::Dark6,
        };
        for (int i = 0; i < 6; ++i) {
            uint16_t col = static_cast<uint16_t>(6 + i * 2);
            XLCellReference ref(1u, col);
            std::string nm = "D" + std::to_string(i+1) + "_Region";
            wks.slicers().add(ref.address(), table, "Region")
                .name(nm)
                .caption("Dark" + std::to_string(i+1))
                .style(darks[i])
                .size(120, 160);
        }
        std::cout << "[Sheet3] Dark1~Dark6 共 6 种切片器样式\n";
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Sheet 4: 创建后修改 — setCaption / setStyle / setColumnCount / setLockedPosition
    //   验证点: 切片器标题、样式是否按修改后生效
    // ─────────────────────────────────────────────────────────────────────────
    {
        doc.workbook().addWorksheet("Sheet4_创建后修改");
        auto wks = doc.workbook().worksheet("Sheet4_创建后修改");
        fillData(wks);
        auto table = wks.tables().add("T4_Data", "A1:D9");

        // 创建时用 Light1，后面改成 Dark5
        wks.slicers().add("F1", table, "Region")
            .name("Mod_Region")
            .caption("原始标题(Light1)")
            .style(XLSlicerStyle::Light1)
            .size(150, 220);

        // 创建后立即修改
        wks.slicers().reload();
        auto& coll = wks.slicers();

        REQUIRE(coll.contains("Mod_Region"));
        auto& s = coll["Mod_Region"];
        s.setCaption("修改后标题 [Dark5] ✓")
         .setStyle(XLSlicerStyle::Dark5);

        REQUIRE(s.styleRaw() == "SlicerStyleDark5");
        REQUIRE(s.caption() == "修改后标题 [Dark5] ✓");

        // 2列布局切片器
        wks.slicers().add("H1", table, "Product")
            .name("Mod_Product")
            .caption("产品 [2列布局]")
            .style(XLSlicerStyle::Light2)
            .size(240, 130)
            .columnCount(2);

        // 降序 + 锁定位置
        wks.slicers().add("J1", table, "Quarter")
            .name("Mod_Quarter")
            .caption("季度 [降序+锁定]")
            .style(XLSlicerStyle::Dark2)
            .size(150, 200)
            .sortDescending()
            .lockedPosition();

        // 带像素偏移
        wks.slicers().add("L1", table, "Sales")
            .name("Mod_Sales")
            .caption("销售额 [+10px偏移]")
            .style(XLSlicerStyle::Light4)
            .size(150, 200)
            .offset(10, 10);

        std::cout << "[Sheet4] 创建后修改: Dark5样式, 2列, 降序, 偏移\n";
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Sheet 5: 多切片器 + XLSlicerCollection 集合操作
    //   验证点: count() / contains() / find() / operator[] / range-for
    // ─────────────────────────────────────────────────────────────────────────
    {
        doc.workbook().addWorksheet("Sheet5_集合操作");
        auto wks = doc.workbook().worksheet("Sheet5_集合操作");
        fillData(wks);
        auto table = wks.tables().add("T5_Data", "A1:D9");

        // 创建 4 个切片器
        wks.slicers().add("F1", table, "Region")
            .name("C_Region").caption("地区").style(XLSlicerStyle::Light2).size(140, 190);
        wks.slicers().add("H1", table, "Product")
            .name("C_Product").caption("产品").style(XLSlicerStyle::Dark1).size(140, 190);
        wks.slicers().add("J1", table, "Quarter")
            .name("C_Quarter").caption("季度").style(XLSlicerStyle::Light5).size(140, 190);
        wks.slicers().add("L1", table, "Sales")
            .name("C_Sales").caption("销售额").style(XLSlicerStyle::Dark4).size(140, 190);

        // 用 collection 验证并写报告
        wks.slicers().reload();
        auto& coll = wks.slicers();

        REQUIRE(coll.count() == 4);
        REQUIRE(coll.contains("C_Region"));
        REQUIRE(coll.contains("C_Product"));
        REQUIRE(!coll.contains("NonExistent"));

        // 按名访问
        REQUIRE(coll["C_Region"].name() == "C_Region");

        // find() — 安全访问（不抛异常）
        auto found = coll.find("C_Quarter");
        REQUIRE(found.valid());
        REQUIRE(found.name() == "C_Quarter");

        auto notFound = coll.find("Nope");
        REQUIRE(!notFound.valid());

        // 将集合信息写入 N 列（用于人眼校验）
        wks.cell("N1").value() = "切片器集合报告";
        wks.cell("N2").value() = "count=" + std::to_string(coll.count());
        wks.cell("N3").value() = "contains(C_Region)=" + std::string(coll.contains("C_Region") ? "true" : "false");
        wks.cell("N4").value() = "contains(NonExistent)=" + std::string(coll.contains("NonExistent") ? "true" : "false");

        int row = 5;
        for (auto& s : coll) {
            wks.cell(static_cast<uint32_t>(row), 14).value() =
                s.name() + " | " + s.caption() + " | " + s.styleRaw();
            ++row;
        }
        std::cout << "[Sheet5] " << coll.count() << " 个切片器，集合报告写入 N 列\n";
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Sheet 6: Pivot Table 切片器
    //   验证点: Pivot 字段切片器可正常筛选数据
    // ─────────────────────────────────────────────────────────────────────────
    {
        doc.workbook().addWorksheet("Sheet6_Pivot切片器");
        auto wks = doc.workbook().worksheet("Sheet6_Pivot切片器");
        fillData(wks);

        auto pt = wks.addPivotTable(
            XLPivotTableOptions("PT_Sales", "A1:D9", "F1")
                .addRowField("Region")
                .addRowField("Product")
                .addColumnField("Quarter")
                .addDataField("Sales", "总销售额")
        );

        if (pt.valid()) {
            // Region 切片器
            wks.slicers().add("M1", pt, "Region")
                .caption("地区筛选 [Pivot]")
                .style(XLSlicerStyle::Dark4)
                .size(160, 220);

            // Product 切片器
            wks.slicers().add("O1", pt, "Product")
                .caption("产品筛选 [Pivot]")
                .style(XLSlicerStyle::Light6)
                .size(160, 220);

            std::cout << "[Sheet6] Pivot Table 切片器已创建\n";
        } else {
            wks.cell("M1").value() = "注: 此环境暂不支持 Pivot Table，切片器跳过";
            std::cout << "[Sheet6] Pivot Table 创建失败，已跳过\n";
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Sheet 7: 删除操作验证
    //   验证点: deleteSlicer() 后切片器消失，关联 Drawing/Cache 也被清理
    // ─────────────────────────────────────────────────────────────────────────
    {
        doc.workbook().addWorksheet("Sheet7_删除验证");
        auto wks = doc.workbook().worksheet("Sheet7_删除验证");
        fillData(wks);
        auto table = wks.tables().add("T7_Data", "A1:D9");

        // 创建 2 个，然后删除 1 个
        wks.slicers().add("F1", table, "Region")
            .name("Del_Region").caption("将被删除").style(XLSlicerStyle::Light1).size(144, 200);
        wks.slicers().add("H1", table, "Product")
            .name("Del_Product").caption("保留的切片器").style(XLSlicerStyle::Dark3).size(144, 200);

        wks.slicers().reload();
        REQUIRE(wks.slicers().count() == 2);

        wks.deleteSlicer("Del_Region");
        REQUIRE(wks.slicers().count() == 1);
        REQUIRE(!wks.slicers().contains("Del_Region"));
        REQUIRE(wks.slicers().contains("Del_Product"));

        wks.cell("J1").value() = "校验说明";
        wks.cell("J2").value() = "此 Sheet 原有 2 个切片器";
        wks.cell("J3").value() = "Del_Region 已被删除 → F 列应无切片器";
        wks.cell("J4").value() = "Del_Product 保留 → H 列应有切片器";

        std::cout << "[Sheet7] 删除 1 个切片器，剩余 " << wks.slicers().count() << " 个\n";
    }

    doc.save();

    std::cout << "\n════════════════════════════════════════\n";
    std::cout << "✓ 生成文件: " << outPath << "\n";
    std::cout << "\n人工校验要点:\n";
    std::cout << "  Sheet1  基础切片器  — F/H/J 列有 3 个切片器; Q季度预选Q1+Q2\n";
    std::cout << "  Sheet2  Light样式   — F~P 列有 6 个切片器，颜色从浅到深\n";
    std::cout << "  Sheet3  Dark样式    — F~P 列有 6 个深色切片器\n";
    std::cout << "  Sheet4  创建后修改  — Region 标题='修改后标题[Dark5]✓'; Product 2列布局;\n";
    std::cout << "                        Quarter 降序+锁定; Sales 带偏移\n";
    std::cout << "  Sheet5  集合操作    — 4 个切片器; N 列有集合信息报告\n";
    std::cout << "  Sheet6  Pivot切片器 — M/O 列有 Pivot 字段切片器 (若 Pivot 支持)\n";
    std::cout << "  Sheet7  删除验证    — F 列无切片器 (已删), H 列有切片器 (保留)\n";
    std::cout << "════════════════════════════════════════\n";
}

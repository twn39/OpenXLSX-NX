/**
 * slicer_demo.cpp
 *
 * 综合演示 OpenXLSX-NX 切片器新 API 的完整功能：
 *
 *  Sheet1  — 基础 Table 切片器（单列、多列、不同样式、预筛选）
 *  Sheet2  — 多切片器 + XLSlicerStyle 全枚举展示
 *  Sheet3  — 创建后修改（resize / moveTo / setCaption / setStyle）
 *  Sheet4  — Pivot Table 切片器
 *  Sheet5  — XLSlicerBuilder 链式 API
 *
 * 生成文件: slicer_comprehensive_demo.xlsx
 */

#include "OpenXLSX.hpp"
#include "XLSlicer.hpp"
#include "XLSlicerCollection.hpp"
#include "XLTables.hpp"
#include "XLPivotTable.hpp"

#include <iostream>
#include <string>
#include <vector>

using namespace OpenXLSX;

// ─── 辅助函数 ───────────────────────────────────────────────────────────────

/// 填充一个带表头的数据区域，返回引用 range
void fillSalesData(XLWorksheet& wks, const std::string& origin = "A1")
{
    // 表头
    wks.cell(origin).value() = "Region";
    // 从 origin 向右填充其他列
    XLCellReference ref(origin);
    uint32_t r = ref.row();
    uint16_t c = ref.column();

    wks.cell(r, c+0).value() = "Region";
    wks.cell(r, c+1).value() = "Product";
    wks.cell(r, c+2).value() = "Quarter";
    wks.cell(r, c+3).value() = "Sales";

    // 数据行
    const std::vector<std::tuple<std::string,std::string,std::string,int>> data = {
        {"North", "Widget", "Q1", 1200},
        {"North", "Gadget", "Q2", 900},
        {"South", "Widget", "Q1", 850},
        {"South", "Gadget", "Q3", 1100},
        {"East",  "Widget", "Q2", 700},
        {"East",  "Gadget", "Q4", 1300},
        {"West",  "Widget", "Q3", 600},
        {"West",  "Gadget", "Q1", 1050},
    };
    for (size_t i = 0; i < data.size(); ++i) {
        uint32_t row = static_cast<uint32_t>(r + 1 + i);
        wks.cell(row, c+0).value() = std::get<0>(data[i]);
        wks.cell(row, c+1).value() = std::get<1>(data[i]);
        wks.cell(row, c+2).value() = std::get<2>(data[i]);
        wks.cell(row, c+3).value() = std::get<3>(data[i]);
    }
}

// ─── Sheet1: 基础 Table 切片器 ──────────────────────────────────────────────

void buildSheet1(XLDocument& doc)
{
    auto wks = doc.workbook().worksheet("Sheet1");
    wks.setName("基础切片器");
    fillSalesData(wks, "A1");

    auto& tables = wks.tables();
    auto table = tables.add("T_Sheet1", "A1:D9");

    // 切片器 1: Region — Light1 样式，默认尺寸
    wks.slicers().add("F1", table, "Region")
        .caption("地区 (Light1)")
        .style(XLSlicerStyle::Light1)
        .size(144, 200);

    // 切片器 2: Product — Dark3 样式
    wks.slicers().add("H1", table, "Product")
        .caption("产品 (Dark3)")
        .style(XLSlicerStyle::Dark3)
        .size(144, 200);

    // 切片器 3: Quarter — 预选只显示 Q1 和 Q2
    wks.slicers().add("J1", table, "Quarter")
        .caption("季度 [Q1,Q2 预筛选]")
        .style(XLSlicerStyle::Light3)
        .size(144, 200)
        .showOnly({"Q1", "Q2"});

    std::cout << "[Sheet1] 3 个 Table 切片器已创建\n";
}

// ─── Sheet2: 多切片器 + 样式枚举 ────────────────────────────────────────────

void buildSheet2(XLDocument& doc)
{
    doc.workbook().addWorksheet("样式展示");
    auto wks = doc.workbook().worksheet("样式展示");
    fillSalesData(wks, "A1");
    auto& tables = wks.tables();
    auto table = tables.add("T_Sheet2", "A1:D9");

    // 展示 Light1…Light6
    const std::vector<XLSlicerStyle> lightStyles = {
        XLSlicerStyle::Light1, XLSlicerStyle::Light2, XLSlicerStyle::Light3,
        XLSlicerStyle::Light4, XLSlicerStyle::Light5, XLSlicerStyle::Light6,
    };
    const std::vector<XLSlicerStyle> darkStyles = {
        XLSlicerStyle::Dark1, XLSlicerStyle::Dark2, XLSlicerStyle::Dark3,
        XLSlicerStyle::Dark4, XLSlicerStyle::Dark5, XLSlicerStyle::Dark6,
    };

    // 用 Region 列展示 Light 系列 (F列开始)
    for (size_t i = 0; i < lightStyles.size(); ++i) {
        uint16_t col = static_cast<uint16_t>(6 + i * 2);  // F, H, J, L, N, P
        XLCellReference ref(1, col);
        std::string slicerName = "Light" + std::to_string(i+1) + "_Region";
        wks.slicers().add(ref.address(), table, "Region")
            .name(slicerName)
            .caption("Light" + std::to_string(i+1))
            .style(lightStyles[i])
            .size(120, 180);
    }

    // Dark 系列放在第 11 行
    for (size_t i = 0; i < darkStyles.size(); ++i) {
        uint16_t col = static_cast<uint16_t>(6 + i * 2);
        XLCellReference ref(11, col);
        std::string slicerName = "Dark" + std::to_string(i+1) + "_Region";
        wks.slicers().add(ref.address(), table, "Region")
            .name(slicerName)
            .caption("Dark" + std::to_string(i+1))
            .style(darkStyles[i])
            .size(120, 180);
    }

    std::cout << "[Sheet2] 12 种样式切片器已创建\n";
}

// ─── Sheet3: 创建后修改 (resize / moveTo / setCaption / setStyle / columnCount) ─

void buildSheet3(XLDocument& doc)
{
    doc.workbook().addWorksheet("创建后修改");
    auto wks = doc.workbook().worksheet("创建后修改");
    fillSalesData(wks, "A1");
    auto& tables = wks.tables();
    auto table = tables.add("T_Sheet3", "A1:D9");

    // Step 1: 用默认选项创建，然后修改
    wks.slicers().add("F1", table, "Region")
        .name("Region_Mod")
        .caption("原始标题")
        .style(XLSlicerStyle::Light1)
        .size(100, 150);

    // Step 2: 重新加载并修改
    auto& coll = wks.slicers();
    coll.reload();  // 触发重新加载（模拟实际使用场景）

    if (coll.contains("Region_Mod")) {
        auto& s = coll["Region_Mod"];
        s.setCaption("修改后标题 ✓")
         .setStyle(XLSlicerStyle::Dark5)
         .setColumnCount(2)
         .setShowCaption(true);
        std::cout << "[Sheet3] Region_Mod 修改后: caption='"
                  << s.caption() << "' styleRaw='" << s.styleRaw() << "'\n";
    }

    // Step 3: 创建带锁定位置和降序排序的切片器
    wks.slicers().add("H1", table, "Product")
        .caption("产品 [降序+锁定]")
        .style(XLSlicerStyle::Other1)
        .size(160, 220)
        .sortDescending()
        .lockedPosition();

    // Step 4: 多列布局
    wks.slicers().add("J1", table, "Quarter")
        .caption("季度 [2列布局]")
        .style(XLSlicerStyle::Dark2)
        .size(240, 120)
        .columnCount(2);

    std::cout << "[Sheet3] 修改、降序、多列切片器已创建\n";
}

// ─── Sheet4: Pivot Table 切片器 ─────────────────────────────────────────────

void buildSheet4(XLDocument& doc)
{
    doc.workbook().addWorksheet("透视表切片器");
    auto wks = doc.workbook().worksheet("透视表切片器");

    // 数据源 (A1:D9)
    fillSalesData(wks, "A1");

    // 建立一个 PivotTable
    XLPivotTableOptions ptOpts;
    ptOpts.name      = "PT_SalesAnalysis";
    ptOpts.sourceRef = "A1:D9";
    ptOpts.targetRef = "F1";

    XLPivotTable pt = wks.addPivotTable(ptOpts);
    if (pt.valid()) {
        // 添加行字段
        pt.addRowField("Region");
        pt.addRowField("Product");
        // 添加值字段
        pt.addValueField("Sales", XLPivotAggregation::Sum, "总销售额");
        // 设置筛选字段 (用于切片器)
        pt.addColumnField("Quarter");
    }

    // Pivot Table 切片器（Region 字段）
    wks.slicers().add("L1", pt, "Region")
        .caption("地区筛选 (Pivot)")
        .style(XLSlicerStyle::Dark4)
        .size(160, 220);

    // Pivot Table 切片器（Product 字段）
    wks.slicers().add("N1", pt, "Product")
        .caption("产品筛选 (Pivot)")
        .style(XLSlicerStyle::Light4)
        .size(160, 220);

    std::cout << "[Sheet4] 透视表切片器已创建\n";
}

// ─── Sheet5: XLSlicerCollection 完整遍历展示 ─────────────────────────────────

void buildSheet5(XLDocument& doc)
{
    doc.workbook().addWorksheet("集合遍历验证");
    auto wks = doc.workbook().worksheet("集合遍历验证");
    fillSalesData(wks, "A1");
    auto& tables = wks.tables();
    auto table = tables.add("T_Sheet5", "A1:D9");

    // 创建 3 个切片器
    wks.slicers().add("F1", table, "Region")
        .name("Slicer_Region5")
        .caption("地区").style(XLSlicerStyle::Light2).size(140, 190);
    wks.slicers().add("H1", table, "Product")
        .name("Slicer_Product5")
        .caption("产品").style(XLSlicerStyle::Dark1).size(140, 190);
    wks.slicers().add("J1", table, "Quarter")
        .name("Slicer_Quarter5")
        .caption("季度").style(XLSlicerStyle::Light5).size(140, 190);

    // 用报告行记录集合信息（写到 M 列）
    wks.slicers().reload();  // 确保集合已加载

    auto& coll = wks.slicers();
    wks.cell("M1").value() = "切片器信息报告";
    wks.cell("M2").value() = "总数: " + std::to_string(coll.count());
    wks.cell("M3").value() = "contains(Slicer_Region5): " +
                              std::string(coll.contains("Slicer_Region5") ? "true" : "false");

    int row = 4;
    for (auto& s : coll) {
        wks.cell(static_cast<uint32_t>(row), 13).value() =
            "名称=" + s.name() + " 标题=" + s.caption() +
            " 样式=" + s.styleRaw();
        ++row;
    }

    std::cout << "[Sheet5] " << coll.count() << " 个切片器，集合遍历写入 M 列\n";
}

// ─── main ───────────────────────────────────────────────────────────────────

int main()
{
    const std::string outPath = "slicer_comprehensive_demo.xlsx";
    std::cout << "OpenXLSX-NX 切片器 API 综合演示\n";
    std::cout << "输出文件: " << outPath << "\n\n";

    try {
        XLDocument doc;
        doc.create(outPath, XLForceOverwrite);

        // Sheet1 已存在
        buildSheet1(doc);
        buildSheet2(doc);
        buildSheet3(doc);

        // Sheet4: Pivot 切片器（可能因 PivotTable 依赖而跳过）
        try {
            buildSheet4(doc);
        } catch (const std::exception& e) {
            std::cout << "[Sheet4] Pivot 切片器跳过 (需要完整 Pivot 支持): " << e.what() << "\n";
            // 仍然保留 sheet 但无切片器
        }

        buildSheet5(doc);

        doc.save();
        std::cout << "\n✓ 文件已保存: " << outPath << "\n";
        std::cout << "\n手动校验要点:\n";
        std::cout << "  Sheet1 '基础切片器': 3 个切片器，Light1/Dark3/Light3 样式\n";
        std::cout << "  Sheet2 '样式展示': Light1-6 + Dark1-6 共 12 种样式\n";
        std::cout << "  Sheet3 '创建后修改': 标题修改✓，Dark5样式，2列布局，降序，锁定\n";
        std::cout << "  Sheet4 '透视表切片器': Pivot 字段切片器\n";
        std::cout << "  Sheet5 '集合遍历验证': M 列有切片器信息报告\n";
        std::cout << "\n  切片器锚定精度验证 (旧 Bug 修复):\n";
        std::cout << "    - 144x200px → cx=1371600 cy=1904850 EMU (oneCellAnchor)\n";

    } catch (const std::exception& e) {
        std::cerr << "错误: " << e.what() << "\n";
        return 1;
    }

    return 0;
}

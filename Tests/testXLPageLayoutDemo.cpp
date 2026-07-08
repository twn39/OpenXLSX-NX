#include "TestHelpers.hpp"
#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>
#include <iostream>
#include <string>

using namespace OpenXLSX;

TEST_CASE("GeneratePageLayoutDemo", "[PageLayoutDemo]")
{
    const std::string out = "pagelayout_comprehensive_demo.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    auto wks = doc.workbook().worksheet(1);
    wks.setName("Page Layout Demo");

    // Adjust column widths for better layout
    wks.column(1).setWidth(10);  // ID
    wks.column(2).setWidth(20);  // Task Name
    wks.column(3).setWidth(15);  // Department
    wks.column(4).setWidth(15);  // Priority
    wks.column(5).setWidth(15);  // Status
    wks.column(6).setWidth(25);  // Notes

    // Title / Header Row (Row 1)
    wks.cell("A1").value() = "Task ID";
    wks.cell("B1").value() = "Task Name";
    wks.cell("C1").value() = "Department";
    wks.cell("D1").value() = "Priority";
    wks.cell("E1").value() = "Status";
    wks.cell("F1").value() = "Notes/Remarks";

    // Fill 50 rows of dummy data so that page breaks and repeating titles are visible
    for (uint32_t r = 2; r <= 50; ++r) {
        wks.cell(r, 1).value() = static_cast<int>(r - 1);
        wks.cell(r, 2).value() = "Task #" + std::to_string(r - 1);
        
        // Alternate departments and priorities
        if (r % 3 == 0) {
            wks.cell(r, 3).value() = "Engineering";
            wks.cell(r, 4).value() = "High";
        } else if (r % 3 == 1) {
            wks.cell(r, 3).value() = "Marketing";
            wks.cell(r, 4).value() = "Medium";
        } else {
            wks.cell(r, 3).value() = "Sales";
            wks.cell(r, 4).value() = "Low";
        }

        wks.cell(r, 5).value() = (r % 2 == 0) ? "Completed" : "In Progress";
        wks.cell(r, 6).value() = "Verification row " + std::to_string(r) + " for page layout layout testing.";
    }

    // 1. Configure Page Setup (Orientation, Paper Size, Scaling)
    auto ps = wks.pageSetup();
    ps.setOrientation(XLPageOrientation::Landscape); // Landscape orientation
    ps.setPaperSize(9);                              // A4 Paper (Excel ID = 9)
    ps.setPageOrder("downThenOver");                 // Down then Over
    ps.setFirstPageNumber(1);
    ps.setUseFirstPageNumber(true);

    // 2. Adjust Page Margins (in inches)
    auto pm = wks.pageMargins();
    pm.setTop(1.2);    // 1.2 inches top margin
    pm.setBottom(1.2); // 1.2 inches bottom margin
    pm.setLeft(0.8);   // 0.8 inch left margin
    pm.setRight(0.8);  // 0.8 inch right margin
    pm.setHeader(0.4); // 0.4 inch header margin
    pm.setFooter(0.4); // 0.4 inch footer margin

    // 3. Print Options
    auto po = wks.printOptions();
    po.setGridLines(true);          // Print gridlines
    po.setHeadings(true);           // Print row/column headings (A, B, C / 1, 2, 3)
    po.setHorizontalCentered(true); // Center horizontally on printout
    po.setVerticalCentered(false);  // Align to top vertically

    // 4. Header & Footer Configurations
    auto hf = wks.headerFooter();
    hf.setDifferentFirst(true); // Different header/footer on first page
    
    // First Page Header & Footer
    hf.setFirstHeader("&C&\"Arial,Bold\"&16 PAGE LAYOUT TEST - COVER PAGE");
    hf.setFirstFooter("&C&10 Confidential - First Page Verification");

    // Standard Odd/Even Page Header & Footer (Page 2 onwards)
    hf.setOddHeader("&L&A &CHidden Corp &R&D &T"); 
    hf.setOddFooter("&CPage &P of &N");

    // 5. Print Titles and Print Area
    wks.setPrintTitleRows(1, 1);    // Repeat Row 1 at the top of every printed page
    wks.setPrintArea("A1:F50");     // Restrict printable area to A1:F50

    // 6. Manual Page Breaks
    wks.insertRowBreak(25);         // Force page break immediately after row 25
    wks.insertColBreak(3);          // Force page break immediately after column 3 (Column C)

    doc.save();
    doc.close();

    std::cout << "\n===================================================\n";
    std::cout << "  页面布局测试文件生成完毕: " << out << "\n";
    std::cout << "  您可以打开此文件并查看“打印预览”，人工校验以下项：\n";
    std::cout << "  1. 页面方向为【横向 (Landscape)】，纸张大小为【A4】\n";
    std::cout << "  2. 第一页页眉显示加粗居中标题，第二页以后页眉显示【Hidden Corp】和日期时间\n";
    std::cout << "  3. 打印网格线、行号列标已启用，且内容水平居中\n";
    std::cout << "  4. 第一行（标题行）在每一页上重复打印\n";
    std::cout << "  5. 第 25 行后与 C 列后有强制的分页线\n";
    std::cout << "===================================================\n\n";

    REQUIRE(true);
}

#include "TestHelpers.hpp"
#include <OpenXLSX.hpp>
#include "XLComments.hpp"

#include <catch2/catch_test_macros.hpp>
#include <iostream>
#include <string>
#include <vector>

using namespace OpenXLSX;

TEST_CASE("GenerateRichTextDemo", "[RichTextDemo]")
{
    const std::string out = "richtext_comprehensive_demo.xlsx";
    XLDocument doc;
    doc.create(out, XLForceOverwrite);

    auto wks = doc.workbook().worksheet(1);
    wks.setName("Rich Text Demo");

    // Adjust column widths for better visual readability
    wks.column(1).setWidth(25); // Feature Category
    wks.column(2).setWidth(40); // Description
    wks.column(3).setWidth(60); // Demonstration

    // Set title row
    wks.row(1).setHeight(30);
    XLRichText titleRt;
    titleRt.addRun("OpenXLSX Rich Text Feature Validation Demo")
        .setFontName("Segoe UI")
        .setFontSize(16)
        .setBold(true)
        .setFontColor(XLColor("FF002060")); // Dark Blue
    wks.cell("A1").value() = titleRt;

    // Table Headers
    wks.row(3).setHeight(20);
    wks.cell("A3").value() = "Feature Category";
    wks.cell("B3").value() = "Description";
    wks.cell("C3").value() = "Rich Text Demonstration";

    struct DemoItem {
        std::string category;
        std::string desc;
        XLRichText  richText;
    };

    std::vector<DemoItem> items;

    // 1. Basic font weight/posture/line effects
    XLRichText rtBasic;
    rtBasic.addRun("This is normal, ");
    rtBasic.addRun("this is bold, ").setBold(true);
    rtBasic.addRun("this is italic, ").setItalic(true);
    rtBasic.addRun("this is bold & italic, ").setBold(true).setItalic(true);
    rtBasic.addRun("and this has a strikethrough.").setStrikethrough(true);
    items.push_back({"Basic Font Styles", "Combinations of bold, italic, and strikethrough effects.", rtBasic});

    // 2. Underline variants
    XLRichText rtUnderlines;
    rtUnderlines.addRun("None").setUnderlineStyle(XLUnderlineNone);
    rtUnderlines.addRun(" | ");
    rtUnderlines.addRun("Single").setUnderlineStyle(XLUnderlineSingle);
    rtUnderlines.addRun(" | ");
    rtUnderlines.addRun("Double").setUnderlineStyle(XLUnderlineDouble);
    rtUnderlines.addRun(" | ");
    rtUnderlines.addRun("Single Accounting").setUnderlineStyle(XLUnderlineSingleAccounting);
    rtUnderlines.addRun(" | ");
    rtUnderlines.addRun("Double Accounting").setUnderlineStyle(XLUnderlineDoubleAccounting);
    items.push_back({"Underline Styles", "All Excel-supported underline variants (standard & accounting).", rtUnderlines});

    // 3. Vertical alignment (Script styles)
    XLRichText rtScript;
    rtScript.addRun("Water: H");
    rtScript.addRun("2").setSubscript(true);
    rtScript.addRun("O.   Einstein: E = mc");
    rtScript.addRun("2").setSuperscript(true);
    rtScript.addRun(".   Baseline text continues here.");
    items.push_back({"Script Alignment", "Superscript (e.g. math exponents) and Subscript (e.g. chemical formulas).", rtScript});

    // 4. Color capabilities
    XLRichText rtColors;
    rtColors.addRun("Crimson Red ").setFontColor(XLColor("FFFF0000"));
    rtColors.addRun("Forest Green ").setFontColor(XLColor("FF228B22"));
    rtColors.addRun("Navy Blue ").setFontColor(XLColor("FF000080"));
    rtColors.addRun("Orange ").setFontColor(XLColor("FFFFA500"));
    rtColors.addRun("Deep Purple ").setFontColor(XLColor("FF800080"));
    items.push_back({"Font Colors", "Vibrant colors defined via RGB Hex values.", rtColors});

    // 5. Font naming and sizing
    XLRichText rtFonts;
    rtFonts.addRun("Arial 18pt ").setFontName("Arial").setFontSize(18);
    rtFonts.addRun("Courier New 10pt ").setFontName("Courier New").setFontSize(10);
    rtFonts.addRun("Times New Roman 14pt ").setFontName("Times New Roman").setFontSize(14).setItalic(true);
    items.push_back({"Font Families & Sizes", "Different font families and sizes mixed inside a single cell.", rtFonts});

    // 6. Ultimate combination
    XLRichText rtUltimate;
    rtUltimate.addRun("Bold Red").setBold(true).setFontColor(XLColor("FFFF0000"));
    rtUltimate.addRun(" meets ");
    rtUltimate.addRun("Italic Green Double Underline").setItalic(true).setUnderlineStyle(XLUnderlineDouble).setFontColor(XLColor("FF008000"));
    rtUltimate.addRun(" meets ");
    rtUltimate.addRun("Superscript Blue").setSuperscript(true).setFontColor(XLColor("FF0000FF"));
    rtUltimate.addRun(" and ");
    rtUltimate.addRun("Strikethrough Orange").setStrikethrough(true).setFontColor(XLColor("FFFFA500")).setFontSize(14);
    items.push_back({"Ultimate Combination", "Overlapping styles combining multiple colors, weights, sizes, and decorations.", rtUltimate});

    // Write all items to the sheet
    for (size_t i = 0; i < items.size(); ++i) {
        uint32_t r = 4 + static_cast<uint32_t>(i);
        wks.row(r).setHeight(25);
        wks.cell(r, 1).value() = items[i].category;
        wks.cell(r, 2).value() = items[i].desc;
        wks.cell(r, 3).value() = items[i].richText;
    }

    // 7. Add cell with a rich-text comment
    uint32_t commentRow = 4 + static_cast<uint32_t>(items.size()) + 1;
    wks.row(commentRow).setHeight(25);
    wks.cell(commentRow, 1).value() = "Comments Validation";
    wks.cell(commentRow, 2).value() = "Hover over the demonstration cell to verify the rich text comment.";

    XLRichText rtComment;
    rtComment.addRun("Attention: ").setBold(true).setFontColor(XLColor("FFFF0000"));
    rtComment.addRun("This is a ").setItalic(true);
    rtComment.addRun("Rich Text Comment").setBold(true).setItalic(true).setFontColor(XLColor("FF0000FF")).setUnderlineStyle(XLUnderlineDouble);
    rtComment.addRun(" containing multiple formatted segments and styles.");

    wks.cell(commentRow, 3).value() = "Hover over this cell";
    wks.comments().setRichText(XLCellReference(commentRow, 3).address(), rtComment, "OpenXLSX Validator", 4, 6);

    doc.save();
    doc.close();

    std::cout << "\n===================================================\n";
    std::cout << "  富文本测试文件生成完毕: " << out << "\n";
    std::cout << "  可以通过人工审核校验该文件中的全部富文本样式。\n";
    std::cout << "===================================================\n\n";

    REQUIRE(true);
}

#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLStyleFacade_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("XLStylesIntegrationTests", "[XLStyles]")
{
    SECTION("Font Style Round-trip")
    {
        const std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();

        // 1. Create and apply style
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks    = doc.workbook().worksheet("Sheet1");
            auto styles = doc.styles();

            // Create a new font
            size_t fontIdx = styles.fonts().create();
            auto   font    = styles.fonts()[fontIdx];
            font.setFontName("Courier New");
            font.setFontSize(16);
            font.setBold(true);

            // Create a cell format (xf) that uses this font
            size_t xfIdx = styles.cellFormats().create();
            styles.cellFormats()[xfIdx].setFontIndex(fontIdx);
            styles.cellFormats()[xfIdx].setApplyFont(true);

            // Apply format to A1
            wks.cell("A1").value() = "Styled Text";
            wks.cell("A1").setCellFormat(xfIdx);

            doc.save();
            doc.close();
        }

        // 2. Re-open and verify
        {
            XLDocument doc;
            doc.open(filename);
            auto wks    = doc.workbook().worksheet("Sheet1");
            auto styles = doc.styles();

            auto   cell  = wks.cell("A1");
            size_t xfIdx = cell.cellFormat();
            REQUIRE(xfIdx > 0);    // Should not be default 0

            auto   xf      = styles.cellFormats()[xfIdx];
            size_t fontIdx = xf.fontIndex();

            auto font = styles.fonts()[fontIdx];
            REQUIRE(font.fontName() == "Courier New");
            REQUIRE(font.fontSize() == 16);
            REQUIRE(font.bold() == true);

            doc.close();
        }
    }

    SECTION("Fill Style Round-trip")
    {
        const std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();

        // 1. Create and apply fill
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks    = doc.workbook().worksheet("Sheet1");
            auto styles = doc.styles();

            size_t fillIdx = styles.fills().create();
            auto   fill    = styles.fills()[fillIdx];
            fill.setPatternType(XLPatternSolid);
            fill.setColor(XLColor(255, 255, 0));    // Yellow

            size_t xfIdx = styles.cellFormats().create();
            styles.cellFormats()[xfIdx].setFillIndex(fillIdx);
            styles.cellFormats()[xfIdx].setApplyFill(true);

            wks.cell("B2").setCellFormat(xfIdx);

            doc.save();
            doc.close();
        }

        // 2. Re-open and verify
        {
            XLDocument doc;
            doc.open(filename);
            auto wks    = doc.workbook().worksheet("Sheet1");
            auto styles = doc.styles();

            auto   cell    = wks.cell("B2");
            size_t xfIdx   = cell.cellFormat();
            auto   fillIdx = styles.cellFormats()[xfIdx].fillIndex();

            auto fill = styles.fills()[fillIdx];
            REQUIRE(fill.patternType() == XLPatternSolid);
            // REQUIRE(fill.color().hex() == "ffffff00");

            doc.close();
        }
    }
}

TEST_CASE("HighLevelXLStyleFacadeIntegration", "[XLStyle]")
{
    XLDocument doc;
    doc.create(__global_unique_file_0(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    XLStyle style;
    style.font.bold   = true;
    style.font.italic = true;
    style.font.color  = XLColor("FF0000");
    style.font.size   = 14;

    style.fill.pattern = XLPatternSolid;
    style.fill.fgColor = XLColor("00FF00");    // green

    style.alignment.horizontal = XLAlignCenter;
    style.alignment.wrapText   = true;

    style.border.left.style = XLLineStyleThin;
    style.border.left.color = XLColor("0000FF");

    style.numberFormat = "#,##0.00";

    wks.cell("A1").value() = 1234.56;
    wks.cell("A1").setStyle(style);

    doc.save();
    doc.close();

    // Verify it was correctly persisted
    XLDocument doc2;
    doc2.open(__global_unique_file_0());
    auto wks2 = doc2.workbook().worksheet("Sheet1");
    auto cell = wks2.cell("A1");

    auto& styles = doc2.styles();
    auto  format = styles.cellFormats().cellFormatByIndex(cell.cellFormat());

    // Verify Font
    REQUIRE(format.applyFont() == true);
    auto font = styles.fonts().fontByIndex(format.fontIndex());
    REQUIRE(font.bold() == true);
    REQUIRE(font.italic() == true);
    REQUIRE(font.fontSize() == 14);
    REQUIRE(font.fontColor().hex() == "FFFF0000");

    // Verify Fill
    REQUIRE(format.applyFill() == true);
    auto fill = styles.fills().fillByIndex(format.fillIndex());
    REQUIRE(fill.patternType() == XLPatternSolid);
    REQUIRE(fill.color().hex() == "FF00FF00");

    // Verify Alignment
    REQUIRE(format.applyAlignment() == true);
    REQUIRE(format.alignment().horizontal() == XLAlignCenter);
    REQUIRE(format.alignment().wrapText() == true);

    // Verify Border
    REQUIRE(format.applyBorder() == true);
    auto border = styles.borders().borderByIndex(format.borderIndex());
    REQUIRE(border.left().style() == XLLineStyleThin);
    REQUIRE(border.left().color().rgb().hex() == "FF0000FF");

    // Verify NumberFormat
    REQUIRE(format.applyNumberFormat() == true);
    auto numFmt = styles.numberFormats().numberFormatById(format.numberFormatId());
    REQUIRE(numFmt.formatCode() == "#,##0.00");
}

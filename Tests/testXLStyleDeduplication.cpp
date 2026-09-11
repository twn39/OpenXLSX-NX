#include <catch2/catch_test_macros.hpp>
#include <OpenXLSX.hpp>
#include <chrono>
#include <filesystem>

using namespace OpenXLSX;

TEST_CASE("Style Pool Deduplication & Normalization", "[XLStyles]")
{
    const std::string testFile = "test_style_dedup.xlsx";
    if (std::filesystem::exists(testFile)) {
        std::filesystem::remove(testFile);
    }

    SECTION("1. Top-level XLStyle zero-allocation deduplication")
    {
        XLDocument doc;
        doc.create(testFile, XLForceOverwrite);
        auto& styles = doc.styles();
        const size_t initialXfCount = styles.cellFormats().count();

        XLStyle customStyle;
        customStyle.font.name = "Arial";
        customStyle.font.size = 14;
        customStyle.font.bold = true;
        customStyle.font.color = XLColor("FFFF0000");
        customStyle.fill.pattern = XLPatternSolid;
        customStyle.fill.fgColor = XLColor("FFFFFF00");
        customStyle.numberFormat = "0.00";

        // First call creates the style
        XLStyleIndex idx1 = styles.findOrCreateStyle(customStyle);
        REQUIRE(idx1 > 0);
        REQUIRE(styles.cellFormats().count() == initialXfCount + 1);

        // Subsequent 20,000 calls must hit the top-level cache and return identical index
        auto start = std::chrono::steady_clock::now();
        for (int i = 0; i < 20000; ++i) {
            XLStyleIndex idx = styles.findOrCreateStyle(customStyle);
            REQUIRE(idx == idx1);
        }
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start);

        // Count must not increase
        REQUIRE(styles.cellFormats().count() == initialXfCount + 1);
        // 20k lookups should take less than 100ms
        CHECK(duration.count() < 100);

        doc.close();
    }

    SECTION("2. ECMA-376 Standard Number Format Registry")
    {
        XLDocument doc;
        doc.create(testFile, XLForceOverwrite);
        auto& styles = doc.styles();

        const size_t initialNumFmtCount = styles.numberFormats().count();

        // Built-in standard formats
        uint32_t idGeneral = styles.createNumberFormat("General");
        CHECK(idGeneral == 0);

        uint32_t idDecimal = styles.createNumberFormat("0.00");
        CHECK(idDecimal == 2);

        uint32_t idThousands = styles.createNumberFormat("#,##0");
        CHECK(idThousands == 3);

        uint32_t idPercent = styles.createNumberFormat("0%");
        CHECK(idPercent == 9);

        uint32_t idText = styles.createNumberFormat("@");
        CHECK(idText == 49);

        // Standard formats must NOT generate XML <numFmt> child entries
        CHECK(styles.numberFormats().count() == initialNumFmtCount);

        // Safe lookup of built-in format by ID should not throw
        REQUIRE_NOTHROW(styles.numberFormats().numberFormatById(2));
        CHECK(styles.numberFormats().numberFormatById(2).formatCode() == "0.00");

        // Custom format must generate a new entry with ID >= 164
        uint32_t idCustom = styles.createNumberFormat("[$$-409]#,##0.00;[RED]-[$$-409]#,##0.00");
        CHECK(idCustom >= 164);
        CHECK(styles.numberFormats().count() == initialNumFmtCount + 1);

        doc.close();
    }

    SECTION("3. ECMA-376 Invariants (fills 0/1, cellXfs 0)")
    {
        XLDocument doc;
        doc.create(testFile, XLForceOverwrite);
        auto& styles = doc.styles();

        // Fills: index 0 is None, index 1 is Gray125
        REQUIRE(styles.fills().count() >= 2);
        CHECK(styles.fills()[0].patternType() == XLPatternNone);
        CHECK(styles.fills()[1].patternType() == XLPatternGray125);

        // cellXfs: index 0 is Normal default
        REQUIRE(styles.cellFormats().count() >= 1);

        doc.close();
    }

    SECTION("4. Full workbook compactStyles with row, col and cell remapping")
    {
        {
            XLDocument doc;
            doc.create(testFile, XLForceOverwrite);
            auto& styles = doc.styles();
            auto wks1 = doc.workbook().worksheet("Sheet1");
            doc.workbook().addWorksheet("Sheet2");
            auto wks2 = doc.workbook().worksheet("Sheet2");

            XLStyle styleCol;
            styleCol.font.bold = true;
            styleCol.font.color = XLColor("FF0000FF"); // Blue
            XLStyleIndex colXf = styles.findOrCreateStyle(styleCol);

            XLStyle styleRow;
            styleRow.fill.pattern = XLPatternSolid;
            styleRow.fill.fgColor = XLColor("FFFFFF00"); // Yellow
            XLStyleIndex rowXf = styles.findOrCreateStyle(styleRow);

            XLStyle styleCell;
            styleCell.numberFormat = "0.00";
            XLStyleIndex cellXf = styles.findOrCreateStyle(styleCell);

            // Create some unreferenced orphan styles
            XLStyle orphanStyle1;
            orphanStyle1.font.italic = true;
            orphanStyle1.font.size = 20;
            styles.findOrCreateStyle(orphanStyle1);

            XLStyle orphanStyle2;
            orphanStyle2.fill.pattern = XLPatternSolid;
            orphanStyle2.fill.fgColor = XLColor("FFFF00FF");
            styles.findOrCreateStyle(orphanStyle2);

            // Apply formats across sheet elements
            // Col formatting
            wks1.column(2).setFormat(colXf);
            // Row formatting
            wks1.row(3).setFormat(rowXf);
            // Cell formatting
            wks1.cell("B5").setCellFormat(cellXf);
            wks1.cell("C5").setCellFormat(cellXf); // duplicate usage
            wks1.cell("D5").setCellFormat(colXf);  // sharing col style
            wks2.cell("A1").setCellFormat(rowXf);  // sharing row style on sheet2

            size_t beforeCount = styles.cellFormats().count();
            REQUIRE(beforeCount >= 6); // Normal(0) + 3 used + 2 orphans

            // Execute style compaction
            doc.compactStyles();

            size_t afterCount = styles.cellFormats().count();
            // Should keep Normal(0) + 3 distinct active styles = 4
            REQUIRE(afterCount == 4);

            // Verify styles are still consistent on sheet elements
            XLStyleIndex newColXf = wks1.column(2).format();
            XLStyleIndex newRowXf = wks1.row(3).format();
            XLStyleIndex newCellXf = wks1.cell("B5").cellFormat();

            CHECK(newColXf > 0);
            CHECK(newRowXf > 0);
            CHECK(newCellXf > 0);
            CHECK(newColXf != newRowXf);
            CHECK(newRowXf != newCellXf);

            CHECK(wks1.cell("C5").cellFormat() == newCellXf);
            CHECK(wks1.cell("D5").cellFormat() == newColXf);
            CHECK(wks2.cell("A1").cellFormat() == newRowXf);

            doc.save();
            doc.close();
        }

        // Reopen saved file and verify persistence and validity
        {
            XLDocument doc;
            doc.open(testFile);
            auto& styles = doc.styles();
            auto wks1 = doc.workbook().worksheet("Sheet1");
            auto wks2 = doc.workbook().worksheet("Sheet2");

            REQUIRE(styles.cellFormats().count() == 4);
            XLStyleIndex colXf = wks1.column(2).format();
            XLStyleIndex rowXf = wks1.row(3).format();
            XLStyleIndex cellXf = wks1.cell("B5").cellFormat();

            CHECK(colXf > 0);
            CHECK(rowXf > 0);
            CHECK(cellXf > 0);
            CHECK(wks1.cell("C5").cellFormat() == cellXf);
            CHECK(wks1.cell("D5").cellFormat() == colXf);
            CHECK(wks2.cell("A1").cellFormat() == rowXf);

            // Verify fills ECMA invariants remain intact
            REQUIRE(styles.fills().count() >= 2);
            CHECK(styles.fills()[0].patternType() == XLPatternNone);
            CHECK(styles.fills()[1].patternType() == XLPatternGray125);

            doc.close();
        }
    }

    if (std::filesystem::exists(testFile)) {
        std::filesystem::remove(testFile);
    }
}

#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <filesystem>

using namespace OpenXLSX;

TEST_CASE("XLStylesTests", "[XLStyles]")
{
    SECTION("Number Formats")
    {
        XLDocument doc;
        doc.create("./testXLStylesNumFmt.xlsx", XLForceOverwrite);
        auto styles = doc.styles();

        auto   numFmts      = styles.numberFormats();
        size_t initialCount = numFmts.count();

        XLNumberFormat nf;
        size_t         idx = numFmts.create(nf);
        numFmts[idx].setNumberFormatId(100);
        numFmts[idx].setFormatCode("0.00%");

        REQUIRE(numFmts.count() == initialCount + 1);
        REQUIRE(numFmts[idx].numberFormatId() == 100);
        REQUIRE(numFmts[idx].formatCode() == "0.00%");

        doc.close();
    }

    SECTION("Custom Number Formats")
    {
        XLDocument doc;
        doc.create("./testXLStylesCustomNumFmt.xlsx", XLForceOverwrite);
        auto styles = doc.styles();

        size_t initialCount = styles.numberFormats().count();

        // Add a new custom number format
        uint32_t nf1 = styles.createNumberFormat("YYYY-MM-DD");
        REQUIRE(nf1 >= 164);    // Excel custom formats typically start at 164

        // Add another one
        uint32_t nf2 = styles.createNumberFormat("0.000%");
        REQUIRE(nf2 == nf1 + 1);

        // Add a duplicate to test deduplication
        uint32_t nf3 = styles.createNumberFormat("YYYY-MM-DD");
        REQUIRE(nf3 == nf1);

        REQUIRE(styles.numberFormats().count() == initialCount + 2);

        doc.save();
        doc.close();

        // Open it again to verify persistence
        XLDocument doc2;
        doc2.open("./testXLStylesCustomNumFmt.xlsx");
        REQUIRE(doc2.styles().numberFormats().count() == initialCount + 2);
        doc2.close();

        std::filesystem::remove("./testXLStylesCustomNumFmt.xlsx");
    }

    SECTION("Fonts")
    {
        XLDocument doc;
        doc.create("./testXLStylesFonts.xlsx", XLForceOverwrite);
        auto styles = doc.styles();
        auto fonts  = styles.fonts();

        size_t idx  = fonts.create();
        auto   font = fonts[idx];

        font.setFontName("Courier New");
        font.setFontSize(14);
        font.setBold(true);
        font.setItalic(true);
        font.setFontColor(XLColor(255, 128, 64, 32));    // ARGB
        font.setUnderline(XLUnderlineDouble);

        REQUIRE(font.fontName() == "Courier New");
        REQUIRE(font.fontSize() == 14);
        REQUIRE(font.bold() == true);
        REQUIRE(font.italic() == true);
        REQUIRE(font.fontColor().hex() == "FF804020");
        REQUIRE(font.underline() == XLUnderlineDouble);

        doc.close();
    }

    SECTION("Fills")
    {
        XLDocument doc;
        doc.create("./testXLStylesFills.xlsx", XLForceOverwrite);
        auto styles = doc.styles();
        auto fills  = styles.fills();

        size_t idx  = fills.create();
        auto   fill = fills[idx];

        fill.setPatternType(XLPatternSolid);
        fill.setColor(XLColor(255, 255, 0, 0));                  // Red
        fill.setBackgroundColor(XLColor(255, 255, 255, 255));    // White

        REQUIRE(fill.patternType() == XLPatternSolid);
        REQUIRE(fill.color().hex() == "FFFF0000");
        REQUIRE(fill.backgroundColor().hex() == "FFFFFFFF");

        doc.close();
    }

    SECTION("Borders")
    {
        XLDocument doc;
        doc.create("./testXLStylesBorders.xlsx", XLForceOverwrite);
        auto styles  = doc.styles();
        auto borders = styles.borders();

        size_t idx    = borders.create();
        auto   border = borders[idx];

        border.setLeft(XLLineStyleThick, XLColor(255, 0, 255, 0));
        border.setRight(XLLineStyleDashed, XLColor(255, 255, 0, 0));
        border.setTop(XLLineStyleDouble, XLColor(255, 0, 0, 255));
        border.setBottom(XLLineStyleThin, XLColor(255, 0, 0, 0));

        REQUIRE(border.left().style() == XLLineStyleThick);
        REQUIRE(border.left().color().rgb().hex() == "FF00FF00");
        REQUIRE(border.right().style() == XLLineStyleDashed);
        REQUIRE(border.top().style() == XLLineStyleDouble);
        REQUIRE(border.bottom().style() == XLLineStyleThin);

        doc.close();
    }

    SECTION("Cell Formats (Xf) & Alignment")
    {
        XLDocument doc;
        doc.create("./testXLStylesXf.xlsx", XLForceOverwrite);
        auto styles = doc.styles();

        auto   cellFormats = styles.cellFormats();
        size_t idx         = cellFormats.create();
        auto   xf          = cellFormats[idx];

        xf.setFontIndex(1);
        xf.setFillIndex(1);
        xf.setBorderIndex(1);
        xf.setApplyFont(true);
        xf.setApplyFill(true);
        xf.setApplyBorder(true);

        auto align = xf.alignment(true);
        align.setHorizontal(XLAlignCenter);
        align.setVertical(XLAlignCenter);
        align.setTextRotation(90);
        xf.setApplyAlignment(true);

        REQUIRE(xf.fontIndex() == 1);
        REQUIRE(xf.fillIndex() == 1);
        REQUIRE(xf.borderIndex() == 1);
        REQUIRE(xf.applyFont() == true);
        REQUIRE(xf.alignment().horizontal() == XLAlignCenter);
        REQUIRE(xf.alignment().textRotation() == 90);

        doc.close();
    }

    SECTION("Cell Styles")
    {
        XLDocument doc;
        doc.create("./testXLStylesCellStyles.xlsx", XLForceOverwrite);
        auto styles = doc.styles();

        auto   cellStyles   = styles.cellStyles();
        size_t initialCount = cellStyles.count();

        size_t idx       = cellStyles.create();
        auto   cellStyle = cellStyles[idx];

        cellStyle.setName("My Custom Style");
        cellStyle.setXfId(0);
        cellStyle.setBuiltinId(0);
        cellStyle.setOutlineStyle(2);
        cellStyle.setHidden(true);
        cellStyle.setCustomBuiltin(true);

        REQUIRE(cellStyles.count() == initialCount + 1);
        REQUIRE(cellStyle.name() == "My Custom Style");
        REQUIRE(cellStyle.xfId() == 0);
        REQUIRE(cellStyle.builtinId() == 0);
        REQUIRE(cellStyle.outlineStyle() == 2);
        REQUIRE(cellStyle.hidden() == true);
        REQUIRE(cellStyle.customBuiltin() == true);

        doc.close();
        std::filesystem::remove("./testXLStylesCellStyles.xlsx");
    }

    SECTION("Diff Cell Formats")
    {
        XLDocument doc;
        doc.create("./testXLStylesDiffCellFormats.xlsx", XLForceOverwrite);
        auto styles = doc.styles();

        auto   diffCellFormats = styles.diffCellFormats();
        size_t initialCount    = diffCellFormats.count();

        size_t idx            = diffCellFormats.create();
        auto   diffCellFormat = diffCellFormats[idx];

        auto font = diffCellFormat.font();
        font.setBold(true);
        font.setFontColor(XLColor(255, 255, 0, 0));    // Red

        auto fill = diffCellFormat.fill();
        fill.setPatternType(XLPatternSolid);
        fill.setBackgroundColor(XLColor(255, 0, 255, 0));    // Green

        REQUIRE(diffCellFormats.count() == initialCount + 1);
        REQUIRE(diffCellFormat.font().bold() == true);
        REQUIRE(diffCellFormat.font().fontColor().hex() == "FFFF0000");
        REQUIRE(diffCellFormat.fill().patternType() == XLPatternSolid);
        REQUIRE(diffCellFormat.fill().backgroundColor().hex() == "FF00FF00");

        doc.close();
        std::filesystem::remove("./testXLStylesDiffCellFormats.xlsx");
    }

    SECTION("Style Deduplication - findOrCreate on sub-pools")
    {
        XLDocument doc;
        doc.create("./testXLStylesDedup.xlsx", XLForceOverwrite);
        auto  styles   = doc.styles();
        auto& fontsRef = styles.fonts();

        // ===== Fonts: identical descriptors must return the same index =====
        size_t idxA = fontsRef.create();
        fontsRef[idxA].setFontName("Helvetica").setFontSize(14).setBold(true);

        size_t idxB = fontsRef.findOrCreate(fontsRef[idxA]);    // same content as idxA
        REQUIRE(idxB == idxA);                                  // deduplication must succeed
        REQUIRE(fontsRef.count() == idxA + 1);                  // no extra node was added

        // ===== Fonts: different descriptor must get a new index =====
        size_t idxC = fontsRef.create();
        fontsRef[idxC].setFontName("Courier").setFontSize(12);
        size_t idxD = fontsRef.findOrCreate(fontsRef[idxC]);
        REQUIRE(idxD == idxC);    // Courier 12 is distinct from Helvetica 14
        REQUIRE(idxD != idxA);

        // ===== Cell Formats: identical xf must deduplicate =====
        auto&  cellFmts = styles.cellFormats();
        size_t xfA      = cellFmts.create();
        cellFmts[xfA].setFontIndex(idxA).setApplyFont(true);

        size_t xfB = cellFmts.findOrCreate(cellFmts[xfA]);
        REQUIRE(xfB == xfA);                     // same structure → same index
        REQUIRE(cellFmts.count() == xfA + 1);    // pool size is stable

        doc.close();
        std::filesystem::remove("./testXLStylesDedup.xlsx");
    }

    SECTION("Style Deduplication - findOrCreateStyle facade")
    {
        XLDocument doc;
        doc.create("./testXLStylesFacadeDedup.xlsx", XLForceOverwrite);
        auto styles = doc.styles();

        XLStyle s;
        s.font.name    = "Arial";
        s.font.bold    = true;
        s.fill.pattern = XLPatternSolid;
        s.fill.fgColor = XLColor("FFFF0000");    // red

        // Repeated calls with the identical descriptor must return the same index
        XLStyleIndex idx1 = styles.findOrCreateStyle(s);
        XLStyleIndex idx2 = styles.findOrCreateStyle(s);
        REQUIRE(idx1 == idx2);
        REQUIRE(idx1 != XLDefaultCellFormat);    // style was created (not the reserved default)

        // A different descriptor must produce a distinct index
        XLStyle s2;
        s2.font.name      = "Courier New";
        s2.font.italic    = true;
        XLStyleIndex idx3 = styles.findOrCreateStyle(s2);
        REQUIRE(idx3 != idx1);

        doc.save();
        doc.close();
        std::filesystem::remove("./testXLStylesFacadeDedup.xlsx");
    }
}

#include <catch2/catch_test_macros.hpp>
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("XLDxf Tests", "[Dxf]") {
    SECTION("Standalone XLDxf Properties") {
        XLDxf dxf;
        dxf.font().setFontColor(XLColor(255, 0, 0)); // Red Font
        dxf.fill().setPatternType(XLPatternSolid);
        dxf.fill().setColor(XLColor(255, 255, 0)); // Yellow Fill

        REQUIRE(dxf.hasFont());
        REQUIRE(dxf.hasFill());
        REQUIRE_FALSE(dxf.hasBorder());
        REQUIRE(dxf.font().fontColor() == XLColor(255, 0, 0));
        REQUIRE(dxf.fill().patternType() == XLPatternSolid);
    }

    SECTION("Integration: Apply DXF to Conditional Formatting") {
        XLDocument doc;
        doc.create("DXFTest.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // 1. Create a DXF (Red font, bold)
        XLDxf dxf;
        dxf.font().setFontColor(XLColor(255, 0, 0));
        dxf.font().setBold(true);
        auto dxfIdx = doc.styles().addDxf(dxf);

        // 2. Add Conditional Formatting
        auto cfIdx = wks.conditionalFormats().create();
        auto cf = wks.conditionalFormats().conditionalFormatByIndex(cfIdx);
        cf.setSqref("A1:A10");
        
        auto ruleIdx = cf.cfRules().create();
        auto rule = cf.cfRules().cfRuleByIndex(ruleIdx);
        rule.setType(XLCfType::CellIs);
        rule.setOperator(XLCfOperator::GreaterThan);
        rule.setFormula("50");
        rule.setDxfId(dxfIdx);

        // 3. Verify
        REQUIRE(rule.dxfId() == dxfIdx);
        
        // Check if styles correctly stored the DXF
        auto storedDxf = doc.styles().dxfs()[dxfIdx];
        REQUIRE(storedDxf.hasFont());
        REQUIRE(storedDxf.font().bold());
        REQUIRE(storedDxf.font().fontColor() == XLColor(255, 0, 0));

        doc.save();
        doc.close();
    }
}

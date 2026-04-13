#include "TestHelpers.hpp"
#include <OpenXLSX.hpp>
#include "XLConditionalFormatting.hpp"

#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_testXLConditionalFormattingFinal_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("CF_OOXML_Validation_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_testXLConditionalFormattingFinal_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("CF_Advanced_Deletion_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("ConditionalFormattingOOXMLStructureValidation", "[ConditionalFormatting][OOXML]")
{
    XLDocument doc;
    doc.create(__global_unique_testXLConditionalFormattingFinal_0(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // 1. Color Scale Rule
    auto csRule = XLColorScaleRule(XLColor("FFFF0000"), XLColor("FF00FF00"));
    wks.addConditionalFormatting("A1:A10", csRule);

    // 2. Data Bar Rule (Should not output empty val="")
    auto dbRule = XLDataBarRule(XLColor("FF0000FF"), true);
    wks.addConditionalFormatting("B1:B10", dbRule);

    // 3. Cell Is Rule (Checks uppercase colors and fgColor/bgColor ordering)
    XLDxf dxfC;
    dxfC.font().setFontColor(XLColor("FFFF0000"));
    dxfC.fill().setFillType(XLPatternFill);
    dxfC.fill().setPatternType(XLPatternSolid);
    dxfC.fill().setColor(XLColor("FFFFFF00"));
    dxfC.fill().setBackgroundColor(XLColor("FFFFFF00"));
    auto cellIsRule = XLCellIsRule(">=", "5");
    wks.addConditionalFormatting("C1:C10", cellIsRule, dxfC);

    // 4. Formula Rule (Checks boolean val="1")
    XLDxf dxfD;
    dxfD.font().setBold(true);
    dxfD.font().setUnderline(XLUnderlineSingle);
    auto formulaRule = XLFormulaRule("MOD(ROW(),2)=0");
    wks.addConditionalFormatting(wks.range("D1:D10"), formulaRule, dxfD);

    doc.save();

    // Now validate the underlying XML structures directly
    auto cfList = wks.conditionalFormats();
    REQUIRE(cfList.count() == 4);

    // Validate Priority assignments (should be global 1, 2, 3, 4)
    REQUIRE(cfList[0].cfRules()[0].priority() == 1);
    REQUIRE(cfList[1].cfRules()[0].priority() == 2);
    REQUIRE(cfList[2].cfRules()[0].priority() == 3);
    REQUIRE(cfList[3].cfRules()[0].priority() == 4);

    // Validate Data Bar empty value absence
    auto dataBarNode = cfList[1].cfRules()[0].node().child("dataBar");
    auto cfvoMinNode = dataBarNode.child("cfvo");
    REQUIRE(cfvoMinNode.attribute("val").empty() == true);    // Should be completely omitted, not val=""

    // Validate DXF configurations in Styles
    auto styles = doc.styles();    // styles is on doc, not workbook
    REQUIRE(styles.dxfs().count() == 2);

    // First DXF (Color and Fill)
    auto dxf1 = styles.dxfs()[0].node();
    REQUIRE(std::string(dxf1.child("font").child("color").attribute("rgb").value()) == "FFFF0000");    // Must be uppercase
    auto fillNode = dxf1.child("fill").child("patternFill");
    REQUIRE(std::string(fillNode.first_child().name()) == "fgColor");    // fgColor MUST precede bgColor in sequence
    REQUIRE(std::string(fillNode.first_child().next_sibling().name()) == "bgColor");

    // Second DXF (Font booleans)
    auto dxf2 = styles.dxfs()[1].node();
    REQUIRE(std::string(dxf2.child("font").child("b").attribute("val").value()) == "1");    // Boolean must be "1" or "0" in this context

    doc.close();
}

TEST_CASE("ConditionalFormattingAdvancedRulesandDeletion", "[ConditionalFormatting][OOXML]")
{
    XLDocument doc;
    doc.create(__global_unique_testXLConditionalFormattingFinal_1(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.addConditionalFormatting("A1:A10", XLIconSetRule("3Arrows"));
    wks.addConditionalFormatting("B1:B10", XLTop10Rule(5, true, false));
    wks.addConditionalFormatting("C1:C10", XLAboveAverageRule(true, true, 1));
    wks.addConditionalFormatting("D1:D10", XLDuplicateValuesRule());
    wks.addConditionalFormatting("E1:E10", XLContainsTextRule("Error"));

    auto cfList = wks.conditionalFormats();
    REQUIRE(cfList.count() == 5);

    // Test Deletion
    wks.removeConditionalFormatting("C1:C10");
    cfList = wks.conditionalFormats();
    REQUIRE(cfList.count() == 4);

    doc.save();

    wks.clearAllConditionalFormatting();
    cfList = wks.conditionalFormats();
    REQUIRE(cfList.count() == 0);

    doc.close();
}

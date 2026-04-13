#include "TestHelpers.hpp"
#include <OpenXLSX.hpp>
#include "XLConditionalFormatting.hpp"

#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_testXLConditionalFormatting_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("CFBuilderTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_testXLConditionalFormatting_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("ConditionalFormatting_Verification_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_testXLConditionalFormatting_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("CFTest_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("XLConditionalFormattingTests", "[ConditionalFormatting]")
{
    SECTION("XLCfvo Property Tests")
    {
        XLCfvo cfvo;
        cfvo.setType(XLCfvoType::Percent);
        cfvo.setValue("50");
        cfvo.setGte(false);

        REQUIRE(cfvo.type() == XLCfvoType::Percent);
        REQUIRE(cfvo.value() == "50");
        REQUIRE(cfvo.gte() == false);
    }

    SECTION("XLCfDataBar Property Tests")
    {
        XLCfDataBar dataBar;
        dataBar.setMin(XLCfvoType::Min, "0");
        dataBar.setMax(XLCfvoType::Max, "0");
        dataBar.setColor(XLColor(255, 0, 0));
        dataBar.setShowValue(false);

        REQUIRE(dataBar.min().type() == XLCfvoType::Min);
        REQUIRE(dataBar.max().type() == XLCfvoType::Max);
        REQUIRE(dataBar.color() == XLColor(255, 0, 0));
        REQUIRE(dataBar.showValue() == false);
    }

    SECTION("XLCfColorScale Property Tests")
    {
        XLCfColorScale colorScale;
        colorScale.addValue(XLCfvoType::Min, "0", XLColor(255, 0, 0));
        colorScale.addValue(XLCfvoType::Max, "0", XLColor(0, 255, 0));

        auto cfvos  = colorScale.cfvos();
        auto colors = colorScale.colors();

        REQUIRE(cfvos.size() == 2);
        REQUIRE(colors.size() == 2);
        REQUIRE(cfvos[0].type() == XLCfvoType::Min);
        REQUIRE(colors[0] == XLColor(255, 0, 0));
        REQUIRE(colors[1] == XLColor(0, 255, 0));
    }

    SECTION("XLCfIconSet Property Tests")
    {
        XLCfIconSet iconSet;
        iconSet.setIconSet("3Arrows");
        iconSet.addValue(XLCfvoType::Percent, "0");
        iconSet.addValue(XLCfvoType::Percent, "33");
        iconSet.setShowValue(false);
        iconSet.setReverse(true);

        REQUIRE(iconSet.iconSet() == "3Arrows");
        REQUIRE(iconSet.cfvos().size() == 2);
        REQUIRE(iconSet.showValue() == false);
        REQUIRE(iconSet.reverse() == true);
    }

    SECTION("Multi-Formula Support")
    {
        XLCfRule rule;
        rule.setType(XLCfType::CellIs);
        rule.setOperator(XLCfOperator::Between);
        rule.addFormula("10");
        rule.addFormula("20");

        auto formulas = rule.formulas();
        REQUIRE(formulas.size() == 2);
        REQUIRE(formulas[0] == "10");
        REQUIRE(formulas[1] == "20");

        // OOXML Validation
        int count = 0;
        for (auto& node : rule.node().children("formula")) {
            if (count == 0) REQUIRE(std::string(node.text().get()) == "10");
            if (count == 1) REQUIRE(std::string(node.text().get()) == "20");
            count++;
        }
        REQUIRE(count == 2);

        rule.clearFormulas();
        REQUIRE(rule.formulas().empty());
    }

    SECTION("Worksheet Integration and OOXML Validation")
    {
        XLDocument doc;
        doc.create(__global_unique_testXLConditionalFormatting_2(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        auto cfIdx = wks.conditionalFormats().create();
        auto cf    = wks.conditionalFormats().conditionalFormatByIndex(cfIdx);
        cf.setSqref("A1:A10");

        auto ruleIdx = cf.cfRules().create();
        auto rule    = cf.cfRules().cfRuleByIndex(ruleIdx);
        rule.setType(XLCfType::DataBar);

        XLCfDataBar db;
        db.setColor(XLColor(0, 0, 255));
        db.setMin(XLCfvoType::Number, "10");
        db.setMax(XLCfvoType::Number, "90");
        rule.setDataBar(db);

        // Check if XML is correctly generated before saving
        XMLNode dbNode = rule.dataBar().node();
        REQUIRE_FALSE(dbNode.empty());
        REQUIRE(std::string(dbNode.name()) == "dataBar");

        auto cfvoNodes = dbNode.children("cfvo");
        int  count     = 0;
        for (auto& n : cfvoNodes) {
            if (count == 0) {
                REQUIRE(std::string(n.attribute("type").value()) == "num");
                REQUIRE(std::string(n.attribute("val").value()) == "10");
            }
            count++;
        }
        REQUIRE(count == 2);
        REQUIRE(std::string(dbNode.child("color").attribute("rgb").value()) == "FF0000FF");

        doc.save();
        doc.close();
    }

    SECTION("High-level Builder Functions")
    {
        XLDocument doc;
        doc.create(__global_unique_testXLConditionalFormatting_0(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // Color Scale Rule
        auto csRule = XLColorScaleRule(XLColor("FFFF0000"), XLColor("FF00FF00"));
        wks.addConditionalFormatting("A1:A10", csRule);

        // Data Bar Rule
        auto dbRule = XLDataBarRule(XLColor("FF0000FF"), true);
        wks.addConditionalFormatting("B1:B10", dbRule);

        // Cell Is Rule with Style
        XLDxf dxf;
        dxf.font().setFontColor(XLColor("FFFF00FF"));
        auto cellIsRule = XLCellIsRule(">=", "50");
        wks.addConditionalFormatting("C1:C10", cellIsRule, dxf);

        // Formula Rule
        auto formulaRule = XLFormulaRule("MOD(ROW(),2)=0");
        wks.addConditionalFormatting(wks.range("D1:D10"), formulaRule);

        // Validation
        auto cfList = wks.conditionalFormats();
        REQUIRE(cfList.count() == 4);

        REQUIRE(cfList[0].sqref() == "A1:A10");
        REQUIRE(cfList[0].cfRules().count() == 1);
        REQUIRE(cfList[0].cfRules()[0].type() == XLCfType::ColorScale);

        REQUIRE(cfList[1].sqref() == "B1:B10");
        REQUIRE(cfList[1].cfRules().count() == 1);
        REQUIRE(cfList[1].cfRules()[0].type() == XLCfType::DataBar);

        REQUIRE(cfList[2].sqref() == "C1:C10");
        REQUIRE(cfList[2].cfRules().count() == 1);
        REQUIRE(cfList[2].cfRules()[0].type() == XLCfType::CellIs);
        REQUIRE(cfList[2].cfRules()[0].Operator() == XLCfOperator::GreaterThanOrEqual);
        REQUIRE(cfList[2].cfRules()[0].formula() == "50");
        REQUIRE(cfList[2].cfRules()[0].dxfId() != XLInvalidStyleIndex);

        REQUIRE(cfList[3].sqref() == "D1:D10");
        REQUIRE(cfList[3].cfRules().count() == 1);
        REQUIRE(cfList[3].cfRules()[0].type() == XLCfType::Expression);
        REQUIRE(cfList[3].cfRules()[0].formula() == "MOD(ROW(),2)=0");

        doc.save();
        doc.close();
    }
}

TEST_CASE("ConditionalFormattingExcelGeneration", "[ConditionalFormattingGen]")
{
    XLDocument doc;
    doc.create(__global_unique_testXLConditionalFormatting_1(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // ----- 准备一些测试数据 -----
    // A列：用于测试双色渐变 (1-10)
    // B列：用于测试数据条 (10-100)
    // C列：用于测试单元格数值判断 (1-10，遇到 >= 5 会高亮)
    // D列：用于测试公式判断 (偶数行会高亮)
    for (int i = 1; i <= 10; ++i) {
        wks.cell("A" + std::to_string(i)).value() = i;
        wks.cell("B" + std::to_string(i)).value() = i * 10;
        wks.cell("C" + std::to_string(i)).value() = i;
        wks.cell("D" + std::to_string(i)).value() = "Row " + std::to_string(i);
    }

    // 1. Color Scale Rule (A列: 最小值为红，最大值为绿)
    auto csRule = XLColorScaleRule(XLColor("FFFF0000"), XLColor("FF00FF00"));
    wks.addConditionalFormatting("A1:A10", csRule);

    // 2. Data Bar Rule (B列: 蓝色数据条)
    auto dbRule = XLDataBarRule(XLColor("FF0000FF"), true);
    wks.addConditionalFormatting("B1:B10", dbRule);

    // 3. Cell Is Rule (C列: 值 >= 5 时背景变黄，文字变红)
    XLDxf dxfC;
    dxfC.font().setFontColor(XLColor("FFFF0000"));    // 红字
    dxfC.fill().setFillType(XLPatternFill);
    dxfC.fill().setPatternType(XLPatternSolid);
    dxfC.fill().setBackgroundColor(XLColor("FFFFFF00"));    // 黄底
    dxfC.fill().setColor(XLColor("FFFFFF00"));              // fgColor 也设为黄
    auto cellIsRule = XLCellIsRule(">=", "5");
    wks.addConditionalFormatting("C1:C10", cellIsRule, dxfC);

    // 4. Formula Rule (D列: 偶数行字体加粗并带下划线)
    XLDxf dxfD;
    dxfD.font().setBold(true);
    dxfD.font().setUnderline(XLUnderlineSingle);
    auto formulaRule = XLFormulaRule("MOD(ROW(),2)=0");
    wks.addConditionalFormatting(wks.range("D1:D10"), formulaRule, dxfD);

    doc.save();
    doc.close();
    // No success output needed for quiet tests

    REQUIRE(true);    // Dummy requirement to make Catch2 happy
}

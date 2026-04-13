#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("CF_Advanced_OOXML_Test_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("CF_DataBar_Advanced_Test_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("ConditionalFormattingAdvancedFeaturesandOOXMLValidation", "[ConditionalFormatting][OOXML]")
{
    XLDocument doc;
    doc.create(__global_unique_file_0(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // 1. Setup Test Rules and Styles
    XLDxf dxfRed;
    dxfRed.font().setFontColor(XLColor("FFFF0000"));

    // IconSet Rule (should generate 3 cfvo children automatically)
    auto iconRule = XLIconSetRule("3TrafficLights1", false, true);    // hide value, reverse
    wks.addConditionalFormatting("A1:A10", iconRule);

    // Top10 Rule (rank=5, percent=true, bottom=false)
    auto top10Rule = XLTop10Rule(5, true, false);
    wks.addConditionalFormatting("B1:B10", top10Rule, dxfRed);

    // AboveAverage Rule (below average, equal=true, stdDev=1)
    auto avgRule = XLAboveAverageRule(false, true, 1);
    wks.addConditionalFormatting("C1:C10", avgRule, dxfRed);

    // Duplicate/Unique Values (testing Unique)
    auto uniqueRule = XLDuplicateValuesRule(true);
    wks.addConditionalFormatting("D1:D10", uniqueRule);

    // Contains/NotContains Text
    auto containsRule = XLContainsTextRule("Error");
    containsRule.addFormula("NOT(ISERROR(SEARCH(\"Error\",E1)))");    // Explicit formula
    wks.addConditionalFormatting("E1:E10", containsRule);

    auto notContainsRule = XLNotContainsTextRule("Success");
    wks.addConditionalFormatting("F1:F10", notContainsRule);

    // Blanks/Errors
    wks.addConditionalFormatting("G1:G10", XLContainsBlanksRule());
    wks.addConditionalFormatting("H1:H10", XLNotContainsErrorsRule());

    doc.save();

    // 2. Validate Logical Abstractions
    auto cfList = wks.conditionalFormats();
    REQUIRE(cfList.count() == 8);

    // Validate Priority assignments (should be global 1 to 8)
    for (size_t i = 0; i < cfList.count(); ++i) { REQUIRE(cfList[i].cfRules()[0].priority() == i + 1); }

    // 3. Validate OOXML Structural Integrity
    // IconSet checks
    auto iconSetNode = cfList[0].cfRules()[0].node().child("iconSet");
    REQUIRE(std::string(iconSetNode.attribute("iconSet").value()) == "3TrafficLights1");
    REQUIRE(std::string(iconSetNode.attribute("showValue").value()) == "0");
    REQUIRE(std::string(iconSetNode.attribute("reverse").value()) == "1");
    // Ensure exactly 3 cfvo nodes were generated
    size_t cfvoCount = 0;
    for (auto node : iconSetNode.children("cfvo")) {
        REQUIRE(std::string(node.attribute("type").value()) == "percent");
        cfvoCount++;
    }
    REQUIRE(cfvoCount == 3);

    // Top10 checks
    auto top10Node = cfList[1].cfRules()[0].node();
    REQUIRE(std::string(top10Node.attribute("type").value()) == "top10");
    REQUIRE(std::string(top10Node.attribute("rank").value()) == "5");
    REQUIRE(std::string(top10Node.attribute("percent").value()) == "1");
    REQUIRE(top10Node.attribute("bottom").empty() == true);    // Should be omitted since it's false

    // AboveAverage checks
    auto avgNode = cfList[2].cfRules()[0].node();
    REQUIRE(std::string(avgNode.attribute("type").value()) == "aboveAverage");
    REQUIRE(std::string(avgNode.attribute("aboveAverage").value()) == "0");
    REQUIRE(std::string(avgNode.attribute("equalAverage").value()) == "1");
    REQUIRE(std::string(avgNode.attribute("stdDev").value()) == "1");

    // UniqueValues checks
    auto uniqueNode = cfList[3].cfRules()[0].node();
    REQUIRE(std::string(uniqueNode.attribute("type").value()) == "uniqueValues");

    // ContainsText checks
    auto containsNode = cfList[4].cfRules()[0].node();
    REQUIRE(std::string(containsNode.attribute("type").value()) == "containsText");
    REQUIRE(std::string(containsNode.attribute("text").value()) == "Error");
    REQUIRE(std::string(containsNode.attribute("operator").value()) == "containsText");
    REQUIRE(std::string(containsNode.child("formula").text().get()) == "NOT(ISERROR(SEARCH(\"Error\",E1)))");

    // 4. Test Deletion Operations
    wks.removeConditionalFormatting("E1:E10");
    cfList = wks.conditionalFormats();
    REQUIRE(cfList.count() == 7);

    // Ensure the remaining rules are the correct ones
    REQUIRE(cfList[4].sqref() == "F1:F10");
    REQUIRE(cfList[4].cfRules()[0].type() == XLCfType::NotContainsText);

    wks.clearAllConditionalFormatting();
    cfList = wks.conditionalFormats();
    REQUIRE(cfList.count() == 0);

    // Ensure the XML actually doesn't have the elements anymore
    REQUIRE(wks.xmlDocument().document_element().child("conditionalFormatting").empty() == true);

    doc.close();
}

TEST_CASE("ConditionalFormattingAdvancedDataBarExtensions", "[ConditionalFormatting][OOXML]")
{
    XLDocument doc;
    doc.create(__global_unique_file_1(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    auto rule = XLDataBarRule(XLColor(0, 255, 0));
    auto dataBar = rule.dataBar();
    dataBar.setGradient(false);
    dataBar.setBorder(true);
    dataBar.setBorderColor(XLColor(0, 0, 255));
    dataBar.setDirection(XLDataBarDirection::RightToLeft);
    dataBar.setMinLength(15);
    dataBar.setMaxLength(85);
    dataBar.setNegativeFillColor(XLColor(255, 0, 0));
    dataBar.setNegativeBorderColor(XLColor(255, 255, 0));
    dataBar.setAxisPosition(XLDataBarAxisPosition::Middle);
    dataBar.setAxisColor(XLColor(128, 128, 128));

    wks.addConditionalFormatting("A1:A10", rule);
    
    // Now verify
    auto cfList = wks.conditionalFormats();
    REQUIRE(cfList.count() == 1);
    auto ruleNode = cfList[0].cfRules()[0].node();
    
    // Check local extLst (fake GUID insertion)
    auto localExtLst = ruleNode.child("extLst");
    REQUIRE(localExtLst.empty() == false);
    auto localExt = localExtLst.child("ext");
    REQUIRE(std::string(localExt.attribute("uri").value()) == "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}");
    auto x14id = localExt.child("x14:id");
    REQUIRE(x14id.empty() == false);
    std::string assignedId = x14id.text().get();
    
    // Check worksheet-level extLst
    auto wsNode = wks.xmlDocument().document_element();
    auto wsExtLst = wsNode.child("extLst");
    REQUIRE(wsExtLst.empty() == false);
    
    XMLNode wsExt;
    for (auto e = wsExtLst.child("ext"); e; e = e.next_sibling("ext")) {
        if (std::string(e.attribute("uri").value()) == "{78C0D931-6437-407d-A8EE-F0AAD7539E65}") {
            wsExt = e;
            break;
        }
    }
    REQUIRE(wsExt.empty() == false);
    
    auto condFmts = wsExt.child("x14:conditionalFormattings");
    auto condFmt = condFmts.child("x14:conditionalFormatting");
    REQUIRE(condFmt.empty() == false);
    REQUIRE(std::string(condFmt.child("xm:sqref").text().get()) == "A1:A10");
    
    auto x14cfRule = condFmt.child("x14:cfRule");
    REQUIRE(std::string(x14cfRule.attribute("id").value()) == assignedId);
    
    auto x14DataBar = x14cfRule.child("x14:dataBar");
    REQUIRE(x14DataBar.empty() == false);
    REQUIRE(std::string(x14DataBar.attribute("gradient").value()) == "0");
    REQUIRE(std::string(x14DataBar.attribute("border").value()) == "1");
    REQUIRE(std::string(x14DataBar.attribute("direction").value()) == "rightToLeft");
    REQUIRE(x14DataBar.attribute("minLength").as_int() == 15);
    REQUIRE(x14DataBar.attribute("maxLength").as_int() == 85);
    REQUIRE(std::string(x14DataBar.attribute("axisPosition").value()) == "middle");
    
    auto borderColor = x14DataBar.child("x14:borderColor");
    REQUIRE(std::string(borderColor.attribute("rgb").value()) == "FF0000FF");
    
    auto negativeFill = x14DataBar.child("x14:negativeFillColor");
    REQUIRE(std::string(negativeFill.attribute("rgb").value()) == "FFFF0000");

    auto negativeBorder = x14DataBar.child("x14:negativeBorderColor");
    REQUIRE(std::string(negativeBorder.attribute("rgb").value()) == "FFFFFF00");

    auto axisColor = x14DataBar.child("x14:axisColor");
    REQUIRE(std::string(axisColor.attribute("rgb").value()) == "FF808080");

    doc.save();
    doc.close();
}

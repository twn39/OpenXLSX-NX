/**
 * @file testXLRobustnessCycles.cpp
 * @brief Unit tests for architectural decoupling, call-cycle elimination, and AST depth / cycle guards.
 */

#include <OpenXLSX.hpp>
#include <XLDataValidation.hpp>
#include <XLEvaluationContext.hpp>
#include <XLFormulaEngine.hpp>
#include <catch2/catch_all.hpp>
#include <functional>
#include <string>

using namespace OpenXLSX;

TEST_CASE("Architectural Decoupling: XLDataValidation range and cell operations", "[Robustness][DataValidation]")
{
    XLDocument doc;
    doc.create("TestDVDecouple.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    auto dv = wks.dataValidations().append();
    dv.setSqref("A1");
    REQUIRE(dv.sqref() == "A1");

    SECTION("Add range and cells independently without recursion cycles")
    {
        dv.addRange("B1:B5");
        REQUIRE(dv.sqref().find("B1:B5") != std::string::npos);

        dv.addCell("C1");
        REQUIRE(dv.sqref().find("C1") != std::string::npos);

        // Add single cell as range string
        dv.addRange("D1");
        REQUIRE(dv.sqref().find("D1") != std::string::npos);

        // Remove cell and range without recursion cycles
        dv.removeCell("C1");
        REQUIRE(dv.sqref().find("C1") == std::string::npos);

        dv.removeRange("D1");
        REQUIRE(dv.sqref().find("D1") == std::string::npos);
    }
}

TEST_CASE("Robustness: Formula Parser Nesting Depth Guard", "[Robustness][ParserDepth]")
{
    XLFormulaEngine eng;

    SECTION("Normal nesting parses and evaluates correctly")
    {
        std::string normalFormula = "=( ( ( 1 + 2 ) * 3 ) )";
        auto val = eng.evaluate(normalFormula);
        REQUIRE(val.type() == XLValueType::Float);
        REQUIRE(val.get<double>() == Catch::Approx(9.0));
    }

    SECTION("Deeply nested parentheses trigger #DEPTH! safely without stack overflow")
    {
        // 300 levels of nested parentheses
        std::string deepFormula = "=";
        for (int i = 0; i < 300; ++i) deepFormula += "(";
        deepFormula += "42";
        for (int i = 0; i < 300; ++i) deepFormula += ")";

        auto val = eng.evaluate(deepFormula);
        REQUIRE(val.type() == XLValueType::Error);
        REQUIRE(val.getString() == "#DEPTH!");
    }

    SECTION("Deeply nested unary operations trigger #DEPTH! safely")
    {
        std::string deepUnary = "=";
        for (int i = 0; i < 300; ++i) deepUnary += "-";
        deepUnary += "1";

        auto val = eng.evaluate(deepUnary);
        REQUIRE(val.type() == XLValueType::Error);
        REQUIRE(val.getString() == "#DEPTH!");
    }
}

TEST_CASE("Robustness: Formula Circular Reference Detection", "[Robustness][CircularRef]")
{
    XLFormulaEngine eng;

    SECTION("Direct self-referencing cell stops recursion and returns error")
    {
        XLEvalSession session;
        std::function<XLCellValue(std::string_view)> cyclicResolver = [&](std::string_view ref) -> XLCellValue {
            if (ref == "A1") {
                return eng.evaluate("=A1", session);
            }
            return XLCellValue{};
        };
        session.setResolver(cyclicResolver);

        auto val = eng.evaluate("=A1", session);
        REQUIRE(val.type() == XLValueType::Error);
        REQUIRE(val.getString() == "#CIRC!");
    }

    SECTION("Mutual two-cell circular reference stops recursion and returns error")
    {
        XLEvalSession session;
        std::function<XLCellValue(std::string_view)> mutualResolver = [&](std::string_view ref) -> XLCellValue {
            if (ref == "A1") {
                return eng.evaluate("=B1 + 1", session);
            }
            if (ref == "B1") {
                return eng.evaluate("=A1 * 2", session);
            }
            return XLCellValue{};
        };
        session.setResolver(mutualResolver);

        auto val = eng.evaluate("=A1", session);
        REQUIRE(val.type() == XLValueType::Error);
        REQUIRE(val.getString() == "#CIRC!");
    }
}

TEST_CASE("Architectural Decoupling: Document Query Sheet Index", "[Robustness][QuerySheetIndex]")
{
    XLDocument doc;
    doc.create("TestQueryIndex.xlsx", XLForceOverwrite);
    doc.workbook().addWorksheet("Alpha");
    doc.workbook().addWorksheet("Beta");
    doc.workbook().addWorksheet("Gamma");

    REQUIRE(doc.workbook().indexOfSheet("Sheet1") == 1);
    REQUIRE(doc.workbook().indexOfSheet("Alpha") == 2);
    REQUIRE(doc.workbook().indexOfSheet("Beta") == 3);
    REQUIRE(doc.workbook().indexOfSheet("Gamma") == 4);

    // Verify worksheet index lookup runs through optimized indexOfSheet
    auto betaIndex = doc.workbook().indexOfSheet("Beta");
    REQUIRE(betaIndex == 3);
}

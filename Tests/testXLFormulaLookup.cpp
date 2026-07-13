/**
 * @file testXLFormulaLookup.cpp
 * @brief Lookup function edge cases: VLOOKUP/HLOOKUP approx, MATCH types, XLOOKUP wildcards.
 */

#include <XLEvaluationContext.hpp>
#include <XLFormulaEngine.hpp>
#include <catch2/catch_all.hpp>

using namespace OpenXLSX;

TEST_CASE("VLOOKUP approximate and exact", "[XLFormulaLookup][VLOOKUP]")
{
    XLFormulaEngine eng;
    // Ascending key column: 1,3,5,7 → values a,b,c,d
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(1.0)},
        {"A2", XLCellValue(3.0)},
        {"A3", XLCellValue(5.0)},
        {"A4", XLCellValue(7.0)},
        {"B1", XLCellValue(std::string("a"))},
        {"B2", XLCellValue(std::string("b"))},
        {"B3", XLCellValue(std::string("c"))},
        {"B4", XLCellValue(std::string("d"))},
    });
    auto r = XLFormulaEngine::makeResolver(ctx);

    SECTION("exact")
    {
        REQUIRE(eng.evaluate("=VLOOKUP(5,A1:B4,2,FALSE)", r).get<std::string>() == "c");
        REQUIRE(eng.evaluate("=VLOOKUP(99,A1:B4,2,0)", r).type() == XLValueType::Error);
    }

    SECTION("approximate largest <=")
    {
        // 4 → key 3 → b; 7 → d; 0 → #N/A
        REQUIRE(eng.evaluate("=VLOOKUP(4,A1:B4,2,TRUE)", r).get<std::string>() == "b");
        REQUIRE(eng.evaluate("=VLOOKUP(7,A1:B4,2,TRUE)", r).get<std::string>() == "d");
        REQUIRE(eng.evaluate("=VLOOKUP(0,A1:B4,2,TRUE)", r).type() == XLValueType::Error);
    }

    SECTION("wildcard exact")
    {
        XLMapEvaluationContext ctx2({
            {"A1", XLCellValue(std::string("apple"))},
            {"A2", XLCellValue(std::string("apricot"))},
            {"B1", XLCellValue(10.0)},
            {"B2", XLCellValue(20.0)},
        });
        auto r2 = XLFormulaEngine::makeResolver(ctx2);
        REQUIRE(eng.evaluate("=VLOOKUP(\"ap*\",A1:B2,2,FALSE)", r2).get<double>() == Catch::Approx(10.0));
    }
}

TEST_CASE("MATCH types", "[XLFormulaLookup][MATCH]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(10.0)},
        {"A2", XLCellValue(20.0)},
        {"A3", XLCellValue(30.0)},
        {"A4", XLCellValue(40.0)},
        {"B1", XLCellValue(std::string("cat"))},
        {"B2", XLCellValue(std::string("dog"))},
        {"B3", XLCellValue(std::string("fox"))},
    });
    auto r = XLFormulaEngine::makeResolver(ctx);

    SECTION("exact and wildcard")
    {
        REQUIRE(eng.evaluate("=MATCH(30,A1:A4,0)", r).get<int64_t>() == 3);
        REQUIRE(eng.evaluate("=MATCH(\"d*\",B1:B3,0)", r).get<int64_t>() == 2);
    }

    SECTION("type 1 largest <=")
    {
        REQUIRE(eng.evaluate("=MATCH(25,A1:A4,1)", r).get<int64_t>() == 2);    // 20
        REQUIRE(eng.evaluate("=MATCH(40,A1:A4,1)", r).get<int64_t>() == 4);
        REQUIRE(eng.evaluate("=MATCH(5,A1:A4,1)", r).type() == XLValueType::Error);
    }

    SECTION("type -1 smallest >=")
    {
        REQUIRE(eng.evaluate("=MATCH(25,A1:A4,-1)", r).get<int64_t>() == 3);    // 30
        REQUIRE(eng.evaluate("=MATCH(10,A1:A4,-1)", r).get<int64_t>() == 1);
        REQUIRE(eng.evaluate("=MATCH(50,A1:A4,-1)", r).type() == XLValueType::Error);
    }
}

TEST_CASE("XLOOKUP match modes and reverse search", "[XLFormulaLookup][XLOOKUP]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(std::string("alpha"))},
        {"A2", XLCellValue(std::string("beta"))},
        {"A3", XLCellValue(std::string("beta"))},
        {"A4", XLCellValue(std::string("gamma"))},
        {"B1", XLCellValue(1.0)},
        {"B2", XLCellValue(2.0)},
        {"B3", XLCellValue(3.0)},
        {"B4", XLCellValue(4.0)},
        {"C1", XLCellValue(10.0)},
        {"C2", XLCellValue(20.0)},
        {"C3", XLCellValue(30.0)},
        {"C4", XLCellValue(40.0)},
        {"D1", XLCellValue(std::string("t"))},
        {"D2", XLCellValue(std::string("u"))},
        {"D3", XLCellValue(std::string("v"))},
        {"D4", XLCellValue(std::string("w"))},
    });
    auto r = XLFormulaEngine::makeResolver(ctx);

    SECTION("wildcard match_mode 2")
    {
        REQUIRE(eng.evaluate("=XLOOKUP(\"b*\",A1:A4,B1:B4,,2)", r).get<double>() == Catch::Approx(2.0));
    }

    SECTION("search last-to-first")
    {
        // Two "beta" — reverse finds last (row3 → 3)
        REQUIRE(eng.evaluate("=XLOOKUP(\"beta\",A1:A4,B1:B4,,0,-1)", r).get<double>() == Catch::Approx(3.0));
        REQUIRE(eng.evaluate("=XLOOKUP(\"beta\",A1:A4,B1:B4,,0,1)", r).get<double>() == Catch::Approx(2.0));
    }

    SECTION("exact or next smaller / larger linear")
    {
        REQUIRE(eng.evaluate("=XLOOKUP(25,C1:C4,D1:D4,,-1,1)", r).get<std::string>() == "u");    // 20
        REQUIRE(eng.evaluate("=XLOOKUP(25,C1:C4,D1:D4,,1,1)", r).get<std::string>() == "v");     // 30
    }

    SECTION("if_not_found")
    {
        REQUIRE(eng.evaluate("=XLOOKUP(\"zzz\",A1:A4,B1:B4,\"none\")", r).get<std::string>() == "none");
    }
}

TEST_CASE("HLOOKUP approximate", "[XLFormulaLookup][HLOOKUP]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(1.0)},
        {"B1", XLCellValue(5.0)},
        {"C1", XLCellValue(10.0)},
        {"A2", XLCellValue(std::string("x"))},
        {"B2", XLCellValue(std::string("y"))},
        {"C2", XLCellValue(std::string("z"))},
    });
    auto r = XLFormulaEngine::makeResolver(ctx);

    REQUIRE(eng.evaluate("=HLOOKUP(5,A1:C2,2,FALSE)", r).get<std::string>() == "y");
    REQUIRE(eng.evaluate("=HLOOKUP(7,A1:C2,2,TRUE)", r).get<std::string>() == "y");     // 5
    REQUIRE(eng.evaluate("=HLOOKUP(0,A1:C2,2,TRUE)", r).type() == XLValueType::Error);
}

/**
 * @file testXLFormulaMathP0.cpp
 * @brief Gold-style tests for formula engine P0: criteria, empty arithmetic, matrix, 2D broadcast.
 */

#include <XLEvaluationContext.hpp>
#include <XLFormulaEngine.hpp>
#include <catch2/catch_all.hpp>
#include <cmath>

using namespace OpenXLSX;

namespace
{
    XLCellResolver mapResolver(XLMapEvaluationContext& ctx)
    {
        return XLFormulaEngine::makeResolver(ctx);
    }
}

TEST_CASE("P0 COUNTIF/SUMIF criteria blanks and wildcards", "[XLFormulaMathP0][Criteria]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(1.0)},
        {"A2", XLCellValue(2.0)},
        {"A3", XLCellValue(3.0)},
        {"A4", XLCellValue()},                   // blank
        {"A5", XLCellValue(std::string("apple"))},
        {"A6", XLCellValue(std::string("Apply"))},
        {"B1", XLCellValue(10.0)},
        {"B2", XLCellValue(20.0)},
        {"B3", XLCellValue(30.0)},
        {"B4", XLCellValue(40.0)},
        {"B5", XLCellValue(50.0)},
        {"B6", XLCellValue(60.0)},
    });
    auto r = mapResolver(ctx);

    SECTION("numeric comparison")
    {
        REQUIRE(eng.evaluate("=COUNTIF(A1:A3,\">1\")", r).get<int64_t>() == 2);
        REQUIRE(eng.evaluate("=SUMIF(A1:A3,\">=2\",B1:B3)", r).get<double>() == Catch::Approx(50.0));
    }

    SECTION("blank criteria")
    {
        REQUIRE(eng.evaluate("=COUNTIF(A1:A4,\"\")", r).get<int64_t>() == 1);
        REQUIRE(eng.evaluate("=COUNTIF(A1:A4,\"<>\")", r).get<int64_t>() == 3);
    }

    SECTION("wildcard text")
    {
        REQUIRE(eng.evaluate("=COUNTIF(A5:A6,\"a*\")", r).get<int64_t>() == 2);
        REQUIRE(eng.evaluate("=COUNTIF(A5:A6,\"apple\")", r).get<int64_t>() == 1);
    }

    SECTION("SUMIFS multi-criteria")
    {
        // A>1 and B<40 → A2 (B=20) only? A2=2,B2=20; A3=3,B3=30 → both B<40
        REQUIRE(eng.evaluate("=SUMIFS(B1:B3,A1:A3,\">1\",B1:B3,\"<40\")", r).get<double>() == Catch::Approx(50.0));
    }

    SECTION("AVERAGEIFS only numeric averages")
    {
        REQUIRE(eng.evaluate("=AVERAGEIFS(B1:B3,A1:A3,\">=2\")", r).get<double>() == Catch::Approx(25.0));
    }
}

TEST_CASE("P0 empty cell arithmetic coercion", "[XLFormulaMathP0][Empty]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue()},    // blank
        {"B1", XLCellValue(5.0)},
    });
    auto r = mapResolver(ctx);

    // Excel: blank + 5 = 5, blank * 5 = 0
    REQUIRE(eng.evaluate("=A1+B1", r).get<double>() == Catch::Approx(5.0));
    REQUIRE(eng.evaluate("=A1*B1", r).get<double>() == Catch::Approx(0.0));
    REQUIRE(eng.evaluate("=-A1", r).get<double>() == Catch::Approx(0.0));
    REQUIRE(eng.evaluate("=SUM(A1,B1)", r).get<double>() == Catch::Approx(5.0));
}

TEST_CASE("P0 2D broadcast row x column", "[XLFormulaMathP0][Broadcast]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(10.0)},
        {"B1", XLCellValue(20.0)},
        {"A2", XLCellValue(1.0)},
        {"A3", XLCellValue(2.0)},
    });
    auto r = mapResolver(ctx);

    // A1:B1 is 1×2, A2:A3 is 2×1 → result 2×2: [[10,20],[20,40]]
    auto arr = eng.evaluateArray("=A1:B1*A2:A3", r);
    REQUIRE(arr.rows() == 2);
    REQUIRE(arr.cols() == 2);
    REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(10.0));
    REQUIRE(arr.at(0, 1).get<double>() == Catch::Approx(20.0));
    REQUIRE(arr.at(1, 0).get<double>() == Catch::Approx(20.0));
    REQUIRE(arr.at(1, 1).get<double>() == Catch::Approx(40.0));

    // Shape mismatch → #N/A
    XLMapEvaluationContext ctx2({
        {"A1", XLCellValue(1.0)},
        {"B1", XLCellValue(2.0)},
        {"C1", XLCellValue(3.0)},
        {"A2", XLCellValue(1.0)},
        {"B2", XLCellValue(2.0)},
    });
    auto r2 = mapResolver(ctx2);
    // 1×3 vs 1×2 incompatible
    auto bad = eng.evaluate("=A1:C1+A2:B2", r2);
    REQUIRE(bad.type() == XLValueType::Error);
}

TEST_CASE("P0 MMULT MINVERSE MDETERM", "[XLFormulaMathP0][Matrix]")
{
    XLFormulaEngine eng;
    // Identity 2×2 in A1:B2
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(1.0)},
        {"B1", XLCellValue(0.0)},
        {"A2", XLCellValue(0.0)},
        {"B2", XLCellValue(1.0)},
        // [[1,2],[3,4]]
        {"C1", XLCellValue(1.0)},
        {"D1", XLCellValue(2.0)},
        {"C2", XLCellValue(3.0)},
        {"D2", XLCellValue(4.0)},
        // vector [[5],[6]]
        {"E1", XLCellValue(5.0)},
        {"E2", XLCellValue(6.0)},
    });
    auto r = mapResolver(ctx);

    SECTION("MMULT identity * matrix")
    {
        auto arr = eng.evaluateArray("=MMULT(A1:B2,C1:D2)", r);
        REQUIRE(arr.rows() == 2);
        REQUIRE(arr.cols() == 2);
        REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(1.0));
        REQUIRE(arr.at(0, 1).get<double>() == Catch::Approx(2.0));
        REQUIRE(arr.at(1, 0).get<double>() == Catch::Approx(3.0));
        REQUIRE(arr.at(1, 1).get<double>() == Catch::Approx(4.0));
    }

    SECTION("MMULT matrix * vector")
    {
        auto arr = eng.evaluateArray("=MMULT(C1:D2,E1:E2)", r);
        REQUIRE(arr.rows() == 2);
        REQUIRE(arr.cols() == 1);
        // [1,2;3,4]*[5;6] = [17;39]
        REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(17.0));
        REQUIRE(arr.at(1, 0).get<double>() == Catch::Approx(39.0));
    }

    SECTION("MDETERM")
    {
        // det([[1,2],[3,4]]) = -2
        REQUIRE(eng.evaluate("=MDETERM(C1:D2)", r).get<double>() == Catch::Approx(-2.0));
        REQUIRE(eng.evaluate("=MDETERM(A1:B2)", r).get<double>() == Catch::Approx(1.0));
    }

    SECTION("MINVERSE of identity")
    {
        auto inv = eng.evaluateArray("=MINVERSE(A1:B2)", r);
        REQUIRE(inv.rows() == 2);
        REQUIRE(inv.at(0, 0).get<double>() == Catch::Approx(1.0));
        REQUIRE(inv.at(1, 1).get<double>() == Catch::Approx(1.0));
        REQUIRE(std::abs(inv.at(0, 1).get<double>()) < 1e-9);
    }

    SECTION("MINVERSE of [[1,2],[3,4]]")
    {
        auto inv = eng.evaluateArray("=MINVERSE(C1:D2)", r);
        // inv = [[-2,1],[1.5,-0.5]]
        REQUIRE(inv.at(0, 0).get<double>() == Catch::Approx(-2.0));
        REQUIRE(inv.at(0, 1).get<double>() == Catch::Approx(1.0));
        REQUIRE(inv.at(1, 0).get<double>() == Catch::Approx(1.5));
        REQUIRE(inv.at(1, 1).get<double>() == Catch::Approx(-0.5));
    }

    SECTION("MMULT dimension error")
    {
        auto bad = eng.evaluate("=MMULT(A1:B2,E1:E2)", r);    // 2×2 * 2×1 is OK actually
        // 2×2 * 1×2 would fail - use C1:D1 (1×2) * A1:B2 (2×2) fails? 1×2 * 2×2 = 1×2 OK
        // 2×1 * 2×2 fails
        auto bad2 = eng.evaluateArray("=MMULT(E1:E2,C1:D2)", r);
        REQUIRE(bad2.asScalar().type() == XLValueType::Error);
    }
}

/**
 * @file testXLFormulaMathPhaseB.cpp
 * @brief Phase B: BASE/DECIMAL/MULTINOMIAL/SERIESSUM, extra trig, _xlfn., array constants.
 */

#include <XLEvaluationContext.hpp>
#include <XLFormulaEngine.hpp>
#include <catch2/catch_all.hpp>
#include <cmath>

using namespace OpenXLSX;

namespace
{
    XLCellResolver mapResolver(XLMapEvaluationContext& ctx) { return XLFormulaEngine::makeResolver(ctx); }

    bool isErr(XLCellValue v, const char* token)
    {
        return v.type() == XLValueType::Error && v.getString() == token;
    }
}    // namespace

// =============================================================================
// BASE / DECIMAL
// =============================================================================

TEST_CASE("PhaseB BASE and DECIMAL", "[XLFormulaMathPhaseB][Base]")
{
    XLFormulaEngine eng;

    REQUIRE(eng.evaluate("=BASE(10,2)").getString() == "1010");
    REQUIRE(eng.evaluate("=BASE(15,16)").getString() == "F");
    REQUIRE(eng.evaluate("=BASE(7,2,4)").getString() == "0111");
    REQUIRE(eng.evaluate("=DECIMAL(\"1010\",2)").get<double>() == Catch::Approx(10.0));
    REQUIRE(eng.evaluate("=DECIMAL(\"FF\",16)").get<double>() == Catch::Approx(255.0));
    REQUIRE(eng.evaluate("=DECIMAL(\"ff\",16)").get<double>() == Catch::Approx(255.0));
    REQUIRE(isErr(eng.evaluate("=BASE(-1,2)"), "#NUM!"));
    REQUIRE(isErr(eng.evaluate("=BASE(10,1)"), "#NUM!"));
    REQUIRE(isErr(eng.evaluate("=DECIMAL(\"G\",16)"), "#NUM!"));
}

// =============================================================================
// MULTINOMIAL / SERIESSUM
// =============================================================================

TEST_CASE("PhaseB MULTINOMIAL and SERIESSUM", "[XLFormulaMathPhaseB][Series]")
{
    XLFormulaEngine eng;

    // MULTINOMIAL(2,3,4) = 9! / (2!3!4!) = 1260
    REQUIRE(eng.evaluate("=MULTINOMIAL(2,3,4)").get<double>() == Catch::Approx(1260.0));
    REQUIRE(eng.evaluate("=MULTINOMIAL(5,0)").get<double>() == Catch::Approx(1.0));

    // SERIESSUM(2, 0, 1, {1,2,3}) = 1*2^0 + 2*2^1 + 3*2^2 = 1+4+12 = 17
    REQUIRE(eng.evaluate("=SERIESSUM(2,0,1,{1,2,3})").get<double>() == Catch::Approx(17.0));

    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(1.0)},
        {"A2", XLCellValue(2.0)},
        {"A3", XLCellValue(3.0)},
    });
    auto r = mapResolver(ctx);
    REQUIRE(eng.evaluate("=SERIESSUM(2,0,1,A1:A3)", r).get<double>() == Catch::Approx(17.0));
}

// =============================================================================
// Trig extensions
// =============================================================================

TEST_CASE("PhaseB CSC SEC COT ACOT hyperbolic", "[XLFormulaMathPhaseB][Trig]")
{
    XLFormulaEngine eng;
    const double    pi = 3.14159265358979323846;

    REQUIRE(eng.evaluate("=CSC(PI()/2)").get<double>() == Catch::Approx(1.0));
    REQUIRE(eng.evaluate("=SEC(0)").get<double>() == Catch::Approx(1.0));
    REQUIRE(eng.evaluate("=COT(PI()/4)").get<double>() == Catch::Approx(1.0));
    REQUIRE(eng.evaluate("=ACOT(0)").get<double>() == Catch::Approx(pi / 2.0));
    REQUIRE(eng.evaluate("=ACOT(1)").get<double>() == Catch::Approx(pi / 4.0));
    // ACOT(-1) ≈ 3π/4
    REQUIRE(eng.evaluate("=ACOT(-1)").get<double>() == Catch::Approx(3.0 * pi / 4.0));

    REQUIRE(eng.evaluate("=SECH(0)").get<double>() == Catch::Approx(1.0));
    REQUIRE(eng.evaluate("=CSCH(1)").get<double>() == Catch::Approx(1.0 / std::sinh(1.0)));
    REQUIRE(eng.evaluate("=COTH(1)").get<double>() == Catch::Approx(std::cosh(1.0) / std::sinh(1.0)));
    REQUIRE(eng.evaluate("=ACOTH(2)").get<double>() ==
            Catch::Approx(0.5 * std::log((2.0 + 1.0) / (2.0 - 1.0))));
    REQUIRE(isErr(eng.evaluate("=ACOTH(0.5)"), "#NUM!"));
    REQUIRE(isErr(eng.evaluate("=CSC(0)"), "#DIV/0!"));
}

// =============================================================================
// _xlfn. prefix
// =============================================================================

TEST_CASE("PhaseB _xlfn. function prefix stripping", "[XLFormulaMathPhaseB][Xlfn]")
{
    XLFormulaEngine eng;

    REQUIRE(eng.evaluate("=_xlfn.SUM(1,2,3)").get<double>() == Catch::Approx(6.0));
    REQUIRE(eng.evaluate("=_XLFN.BASE(10,2)").getString() == "1010");
    REQUIRE(eng.evaluate("=_xlfn.CEILING.PRECISE(2.5)").get<double>() == Catch::Approx(3.0));
    REQUIRE(eng.evaluate("=_xlfn.CSC(PI()/2)").get<double>() == Catch::Approx(1.0));
}

// =============================================================================
// Array constants
// =============================================================================

TEST_CASE("PhaseB array constant literals", "[XLFormulaMathPhaseB][ArrayLit]")
{
    XLFormulaEngine eng;

    SECTION("row vector")
    {
        auto arr = eng.evaluateArray("={1,2,3}");
        REQUIRE(arr.rows() == 1);
        REQUIRE(arr.cols() == 3);
        REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(1.0));
        REQUIRE(arr.at(0, 2).get<double>() == Catch::Approx(3.0));
    }

    SECTION("column vector")
    {
        auto arr = eng.evaluateArray("={1;2;3}");
        REQUIRE(arr.rows() == 3);
        REQUIRE(arr.cols() == 1);
        REQUIRE(arr.at(2, 0).get<double>() == Catch::Approx(3.0));
    }

    SECTION("matrix")
    {
        auto arr = eng.evaluateArray("={1,2;3,4}");
        REQUIRE(arr.rows() == 2);
        REQUIRE(arr.cols() == 2);
        REQUIRE(arr.at(0, 1).get<double>() == Catch::Approx(2.0));
        REQUIRE(arr.at(1, 0).get<double>() == Catch::Approx(3.0));
    }

    SECTION("implicit intersection scalar")
    {
        REQUIRE(eng.evaluate("={10,20;30,40}").get<double>() == Catch::Approx(10.0));
    }

    SECTION("SUM over array constant")
    {
        REQUIRE(eng.evaluate("=SUM({1,2,3})").get<double>() == Catch::Approx(6.0));
        REQUIRE(eng.evaluate("=SUM({1,2;3,4})").get<double>() == Catch::Approx(10.0));
    }

    SECTION("MMULT with array constants")
    {
        // I * A
        auto arr = eng.evaluateArray("=MMULT({1,0;0,1},{1,2;3,4})");
        REQUIRE(arr.rows() == 2);
        REQUIRE(arr.cols() == 2);
        REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(1.0));
        REQUIRE(arr.at(1, 1).get<double>() == Catch::Approx(4.0));
    }

    SECTION("unary minus in constant")
    {
        auto arr = eng.evaluateArray("={-1,2}");
        REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(-1.0));
    }
}

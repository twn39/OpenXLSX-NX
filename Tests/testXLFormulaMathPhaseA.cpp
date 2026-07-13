/**
 * @file testXLFormulaMathPhaseA.cpp
 * @brief Phase A Excel-compatibility hardening: MOD, CEILING/FLOOR family, SUMPRODUCT, error/coercion.
 */

#include <XLCellValue.hpp>
#include <XLEvaluationContext.hpp>
#include <XLFormulaEngine.hpp>
#include <XLFormulaUtils.hpp>
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
// Utils: coercion & error propagation helpers
// =============================================================================

TEST_CASE("PhaseA numeric string parse", "[XLFormulaMathPhaseA][Coerce]")
{
    double d = 0.0;
    REQUIRE(tryParseNumericString("12.5", d));
    REQUIRE(d == Catch::Approx(12.5));
    REQUIRE(tryParseNumericString("  -3  ", d));
    REQUIRE(d == Catch::Approx(-3.0));
    REQUIRE_FALSE(tryParseNumericString("12x", d));
    REQUIRE_FALSE(tryParseNumericString("TRUE", d));
    REQUIRE_FALSE(tryParseNumericString("", d));

    auto c1 = coerceToNumber(XLCellValue(std::string("42")), true);
    REQUIRE(c1.type() != XLValueType::Error);
    REQUIRE(toDouble(c1) == Catch::Approx(42.0));

    auto blank = coerceToNumber(XLCellValue{}, true);
    REQUIRE(toDouble(blank) == Catch::Approx(0.0));
    REQUIRE(isError(coerceToNumber(XLCellValue{}, false)));

    XLCellValue err;
    err.setError("#DIV/0!");
    auto pe = coerceToNumber(err, true);
    REQUIRE(isErr(pe, "#DIV/0!"));
}

TEST_CASE("PhaseA firstError and numericsOrError", "[XLFormulaMathPhaseA][Coerce]")
{
    XLCellValue e1;
    e1.setError("#N/A");
    XLFormulaArg arg({XLCellValue(1.0), e1, XLCellValue(3.0)}, 1, 3);
    auto         fe = firstError(arg);
    REQUIRE(isErr(fe, "#N/A"));

    XLCellValue outErr;
    auto        nums = numericsOrError(arg, outErr);
    REQUIRE(nums.empty());
    REQUIRE(isErr(outErr, "#N/A"));

    XLFormulaArg clean({XLCellValue(1.0), XLCellValue(2.0)}, 1, 2);
    nums = numericsOrError(clean, outErr);
    REQUIRE(nums.size() == 2);
    REQUIRE_FALSE(isError(outErr));
}

// =============================================================================
// MOD — Excel formula n - d * floor(n/d)
// =============================================================================

TEST_CASE("PhaseA MOD Excel semantics", "[XLFormulaMathPhaseA][MOD]")
{
    XLFormulaEngine eng;

    REQUIRE(eng.evaluate("=MOD(3,2)").get<double>() == Catch::Approx(1.0));
    REQUIRE(eng.evaluate("=MOD(-3,2)").get<double>() == Catch::Approx(1.0));     // not C fmod (-1)
    REQUIRE(eng.evaluate("=MOD(3,-2)").get<double>() == Catch::Approx(-1.0));
    REQUIRE(eng.evaluate("=MOD(-3,-2)").get<double>() == Catch::Approx(-1.0));
    REQUIRE(eng.evaluate("=MOD(5.5,2)").get<double>() == Catch::Approx(1.5));
    REQUIRE(isErr(eng.evaluate("=MOD(1,0)"), "#DIV/0!"));

    // Numeric string coercion
    REQUIRE(eng.evaluate("=MOD(\"5\",\"2\")").get<double>() == Catch::Approx(1.0));
}

// =============================================================================
// CEILING / FLOOR family
// =============================================================================

TEST_CASE("PhaseA CEILING FLOOR classic and precise", "[XLFormulaMathPhaseA][Ceiling]")
{
    XLFormulaEngine eng;

    SECTION("classic positive")
    {
        REQUIRE(eng.evaluate("=CEILING(2.5,1)").get<double>() == Catch::Approx(3.0));
        REQUIRE(eng.evaluate("=FLOOR(3.7,2)").get<double>() == Catch::Approx(2.0));
        REQUIRE(eng.evaluate("=CEILING(2.5,2)").get<double>() == Catch::Approx(4.0));
    }

    SECTION("classic same-sign negatives")
    {
        REQUIRE(eng.evaluate("=CEILING(-2.5,-1)").get<double>() == Catch::Approx(-3.0));
        REQUIRE(eng.evaluate("=FLOOR(-3.7,-2)").get<double>() == Catch::Approx(-2.0));
    }

    SECTION("classic different signs → #NUM!")
    {
        REQUIRE(isErr(eng.evaluate("=CEILING(-2.5,1)"), "#NUM!"));
        REQUIRE(isErr(eng.evaluate("=FLOOR(-2.5,1)"), "#NUM!"));
        REQUIRE(isErr(eng.evaluate("=CEILING(2.5,-1)"), "#NUM!"));
    }

    SECTION("CEILING.PRECISE / FLOOR.PRECISE / ISO.CEILING")
    {
        REQUIRE(eng.evaluate("=CEILING.PRECISE(2.5)").get<double>() == Catch::Approx(3.0));
        REQUIRE(eng.evaluate("=CEILING.PRECISE(-2.5,1)").get<double>() == Catch::Approx(-2.0));
        REQUIRE(eng.evaluate("=CEILING.PRECISE(-2.5,-1)").get<double>() == Catch::Approx(-2.0));    // sign ignored
        REQUIRE(eng.evaluate("=FLOOR.PRECISE(-2.5,1)").get<double>() == Catch::Approx(-3.0));
        REQUIRE(eng.evaluate("=FLOOR.PRECISE(2.5,-1)").get<double>() == Catch::Approx(2.0));
        REQUIRE(eng.evaluate("=ISO.CEILING(-2.5,1)").get<double>() == Catch::Approx(-2.0));
    }
}

// =============================================================================
// SUMPRODUCT shape + errors
// =============================================================================

TEST_CASE("PhaseA SUMPRODUCT 2D shape and errors", "[XLFormulaMathPhaseA][Sumproduct]")
{
    XLFormulaEngine eng;

    SECTION("matching 1D")
    {
        XLMapEvaluationContext ctx({
            {"A1", XLCellValue(1.0)},
            {"A2", XLCellValue(2.0)},
            {"A3", XLCellValue(3.0)},
            {"B1", XLCellValue(4.0)},
            {"B2", XLCellValue(5.0)},
            {"B3", XLCellValue(6.0)},
        });
        auto r = mapResolver(ctx);
        // 1*4 + 2*5 + 3*6 = 4+10+18 = 32
        REQUIRE(eng.evaluate("=SUMPRODUCT(A1:A3,B1:B3)", r).get<double>() == Catch::Approx(32.0));
    }

    SECTION("matching 2D")
    {
        XLMapEvaluationContext ctx({
            {"A1", XLCellValue(1.0)},
            {"B1", XLCellValue(2.0)},
            {"A2", XLCellValue(3.0)},
            {"B2", XLCellValue(4.0)},
            {"C1", XLCellValue(10.0)},
            {"D1", XLCellValue(20.0)},
            {"C2", XLCellValue(30.0)},
            {"D2", XLCellValue(40.0)},
        });
        auto r = mapResolver(ctx);
        // 1*10 + 2*20 + 3*30 + 4*40 = 10+40+90+160 = 300
        REQUIRE(eng.evaluate("=SUMPRODUCT(A1:B2,C1:D2)", r).get<double>() == Catch::Approx(300.0));
    }

    SECTION("shape mismatch → #VALUE! (no silent truncate)")
    {
        XLMapEvaluationContext ctx({
            {"A1", XLCellValue(1.0)},
            {"A2", XLCellValue(2.0)},
            {"A3", XLCellValue(3.0)},
            {"B1", XLCellValue(4.0)},
            {"B2", XLCellValue(5.0)},
        });
        auto r = mapResolver(ctx);
        REQUIRE(isErr(eng.evaluate("=SUMPRODUCT(A1:A3,B1:B2)", r), "#VALUE!"));
    }

    SECTION("non-numeric as zero; text number coerced")
    {
        XLMapEvaluationContext ctx({
            {"A1", XLCellValue(1.0)},
            {"A2", XLCellValue(std::string("x"))},
            {"B1", XLCellValue(2.0)},
            {"B2", XLCellValue(3.0)},
        });
        auto r = mapResolver(ctx);
        // 1*2 + 0*3 = 2
        REQUIRE(eng.evaluate("=SUMPRODUCT(A1:A2,B1:B2)", r).get<double>() == Catch::Approx(2.0));
    }

    SECTION("error propagates")
    {
        XLCellValue e;
        e.setError("#N/A");
        XLMapEvaluationContext ctx({
            {"A1", XLCellValue(1.0)},
            {"A2", e},
            {"B1", XLCellValue(2.0)},
            {"B2", XLCellValue(3.0)},
        });
        auto r = mapResolver(ctx);
        REQUIRE(isErr(eng.evaluate("=SUMPRODUCT(A1:A2,B1:B2)", r), "#N/A"));
    }
}

// =============================================================================
// Error-first on SUM; string coercion on scalar math
// =============================================================================

TEST_CASE("PhaseA SUM error propagation and ABS string coerce", "[XLFormulaMathPhaseA][SumAbs]")
{
    XLFormulaEngine eng;

    XLCellValue e;
    e.setError("#VALUE!");
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(1.0)},
        {"A2", e},
        {"A3", XLCellValue(3.0)},
    });
    auto r = mapResolver(ctx);
    REQUIRE(isErr(eng.evaluate("=SUM(A1:A3)", r), "#VALUE!"));

    REQUIRE(eng.evaluate("=ABS(\"-3.5\")").get<double>() == Catch::Approx(3.5));
    REQUIRE(eng.evaluate("=SQRT(\"9\")").get<double>() == Catch::Approx(3.0));
    REQUIRE(isErr(eng.evaluate("=SQRT(\"x\")"), "#VALUE!"));
}

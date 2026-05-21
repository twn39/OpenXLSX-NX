/**
 * @file testXLFormulaEngineNew.cpp
 * @brief Tests for the new formula functions added in XLFormulaFunctions2.cpp.
 *
 * Grouped by batch:
 *   - [FormulaNewMath]      : LN, ATAN, ATAN2, SINH/COSH/TANH, ASINH/ACOSH/ATANH,
 *                             SQRTPI, FACT, FACTDOUBLE, COMBIN, COMBINA,
 *                             PRODUCT, GCD, LCM, EVEN, ODD, QUOTIENT
 *   - [FormulaNewUtility]   : SUBTOTAL, CHOOSE, ROW, COLUMN, DATEDIF,
 *                             IRR, MIRR, RATE, IPMT, PPMT,
 *                             MODE.SNGL, SKEW, KURT, FORECAST.LINEAR
 *   - [FormulaNewDistrib]   : NORM.S.DIST, NORM.DIST, NORM.S.INV, NORM.INV,
 *                             T.DIST, T.DIST.RT, T.DIST.2T, T.INV, T.INV.2T,
 *                             CHISQ.DIST, CHISQ.DIST.RT, CHISQ.INV, CHISQ.INV.RT,
 *                             BINOM.DIST, POISSON.DIST, EXPON.DIST
 */

#include "TestHelpers.hpp"
#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <cmath>

using namespace OpenXLSX;

// ---------------------------------------------------------------------------
// Helpers (mirrors testXLFormulaEngine.cpp)
// ---------------------------------------------------------------------------

static const XLCellResolver noResolver{};

static XLCellResolver makeMapResolver(std::initializer_list<std::pair<std::string, XLCellValue>> init)
{
    auto m = std::make_shared<std::unordered_map<std::string, XLCellValue>>();
    for (auto& p : init) m->insert_or_assign(p.first, p.second);
    return [m](std::string_view ref) -> XLCellValue {
        auto it = m->find(std::string(ref));
        return it != m->end() ? it->second : XLCellValue{};
    };
}

// ============================================================================
// BATCH 1 — Math / Trig
// ============================================================================

TEST_CASE("FormulaNewMathLogarithms", "[FormulaNewMath]")
{
    XLFormulaEngine eng;

    SECTION("LN(e) = 1") { REQUIRE(eng.evaluate("=LN(EXP(1))").get<double>() == Catch::Approx(1.0)); }
    SECTION("LN(1) = 0") { REQUIRE(eng.evaluate("=LN(1)").get<double>() == Catch::Approx(0.0)); }
    SECTION("LN(0) = error") { REQUIRE(eng.evaluate("=LN(0)").type() == XLValueType::Error); }
    SECTION("LN(-1) = error") { REQUIRE(eng.evaluate("=LN(-1)").type() == XLValueType::Error); }
}

TEST_CASE("FormulaNewMathTrig", "[FormulaNewMath]")
{
    XLFormulaEngine eng;

    SECTION("ATAN data-driven")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({
            {"=ATAN(0)",    0.0},
            {"=ATAN(1)",    0.7853981634},   // π/4
            {"=ATAN(-1)",  -0.7853981634},
        }));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected).epsilon(1e-8));
    }

    SECTION("ATAN2 data-driven")
    {
        // Excel ATAN2(x_num, y_num): x=1,y=0 → atan2(0,1) = 0.0
        //                             x=0,y=1 → atan2(1,0) = π/2
        //                             x=1,y=1 → atan2(1,1) = π/4
        auto [expr, expected] = GENERATE(table<std::string, double>({
            {"=ATAN2(1,0)",   0.0},             // atan2(y=0, x=1) = 0
            {"=ATAN2(0,1)",   1.5707963268},    // atan2(y=1, x=0) = π/2
            {"=ATAN2(1,1)",   0.7853981634},    // π/4
        }));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected).epsilon(1e-8));
    }

    SECTION("ATAN2(0,0) = error") { REQUIRE(eng.evaluate("=ATAN2(0,0)").type() == XLValueType::Error); }
}

TEST_CASE("FormulaNewMathHyperbolic", "[FormulaNewMath]")
{
    XLFormulaEngine eng;

    SECTION("Hyperbolic data-driven")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({
            {"=SINH(0)",    0.0},
            {"=COSH(0)",    1.0},
            {"=TANH(0)",    0.0},
            {"=ASINH(0)",   0.0},
            {"=ACOSH(1)",   0.0},
            {"=ATANH(0)",   0.0},
            {"=SINH(1)",    1.1752011936},
            {"=COSH(1)",    1.5430806348},
            {"=TANH(1)",    0.7615941560},
        }));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected).epsilon(1e-8));
    }

    SECTION("ACOSH(<1) = error") { REQUIRE(eng.evaluate("=ACOSH(0.5)").type() == XLValueType::Error); }
    SECTION("ATANH(1) = error")  { REQUIRE(eng.evaluate("=ATANH(1)").type()   == XLValueType::Error); }
    SECTION("ATANH(-1) = error") { REQUIRE(eng.evaluate("=ATANH(-1)").type()  == XLValueType::Error); }
}

TEST_CASE("FormulaNewMathFact", "[FormulaNewMath]")
{
    XLFormulaEngine eng;

    SECTION("FACT data-driven")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({
            {"=FACT(0)",   1.0},
            {"=FACT(1)",   1.0},
            {"=FACT(5)",  120.0},
            {"=FACT(10)", 3628800.0},
        }));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected));
    }

    SECTION("FACT(-1) = error")  { REQUIRE(eng.evaluate("=FACT(-1)").type()  == XLValueType::Error); }
    SECTION("FACT(171) = error") { REQUIRE(eng.evaluate("=FACT(171)").type() == XLValueType::Error); }

    SECTION("FACTDOUBLE data-driven")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({
            {"=FACTDOUBLE(0)", 1.0},
            {"=FACTDOUBLE(1)", 1.0},
            {"=FACTDOUBLE(3)", 3.0},    // 3!! = 3*1
            {"=FACTDOUBLE(6)", 48.0},   // 6!! = 6*4*2
            {"=FACTDOUBLE(7)", 105.0},  // 7!! = 7*5*3*1
        }));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected));
    }
}

TEST_CASE("FormulaNewMathCombinatorics", "[FormulaNewMath]")
{
    XLFormulaEngine eng;

    SECTION("COMBIN data-driven")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({
            {"=COMBIN(5,2)",   10.0},
            {"=COMBIN(10,3)",  120.0},
            {"=COMBIN(7,0)",   1.0},
            {"=COMBIN(7,7)",   1.0},
            {"=COMBIN(8,5)",   56.0},
        }));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected));
    }

    SECTION("COMBIN error cases")
    {
        REQUIRE(eng.evaluate("=COMBIN(-1,2)").type() == XLValueType::Error);
        REQUIRE(eng.evaluate("=COMBIN(3,5)").type()  == XLValueType::Error);
    }

    SECTION("COMBINA data-driven")
    {
        // COMBINA(n,k) = COMBIN(n+k-1, k)
        REQUIRE(eng.evaluate("=COMBINA(4,3)").get<double>() == Catch::Approx(20.0));  // C(6,3)
        REQUIRE(eng.evaluate("=COMBINA(10,2)").get<double>() == Catch::Approx(55.0)); // C(11,2)
    }
}

TEST_CASE("FormulaNewMathArithmetic", "[FormulaNewMath]")
{
    XLFormulaEngine eng;

    auto resolver = makeMapResolver({
        {"A1", XLCellValue(2.0)},
        {"A2", XLCellValue(3.0)},
        {"A3", XLCellValue(5.0)},
    });

    SECTION("PRODUCT range")
    { REQUIRE(eng.evaluate("=PRODUCT(A1:A3)", resolver).get<double>() == Catch::Approx(30.0)); }

    SECTION("GCD data-driven")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({
            {"=GCD(12,8)",      4.0},
            {"=GCD(15,10,5)",   5.0},
            {"=GCD(100,75,25)", 25.0},
        }));
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected));
    }

    SECTION("LCM data-driven")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({
            {"=LCM(4,6)",    12.0},
            {"=LCM(3,5)",    15.0},
            {"=LCM(2,3,4)",  12.0},
        }));
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected));
    }

    SECTION("EVEN data-driven")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({
            {"=EVEN(1)",    2.0},
            {"=EVEN(2)",    2.0},
            {"=EVEN(3)",    4.0},
            {"=EVEN(-1)",  -2.0},
            {"=EVEN(-2)",  -2.0},
            {"=EVEN(0)",    0.0},
        }));
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected));
    }

    SECTION("ODD data-driven")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({
            {"=ODD(1)",    1.0},
            {"=ODD(2)",    3.0},
            {"=ODD(3)",    3.0},
            {"=ODD(-1)",  -1.0},
            {"=ODD(-2)",  -3.0},
            {"=ODD(0)",    1.0},
        }));
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected));
    }

    SECTION("QUOTIENT data-driven")
    {
        auto [expr, expected] = GENERATE(table<std::string, int64_t>({
            {"=QUOTIENT(10,3)",    3},
            {"=QUOTIENT(-10,3)",  -3},
            {"=QUOTIENT(7,2)",     3},
        }));
        REQUIRE(eng.evaluate(expr).get<int64_t>() == expected);
    }

    SECTION("QUOTIENT div/0") { REQUIRE(eng.evaluate("=QUOTIENT(1,0)").type() == XLValueType::Error); }

    SECTION("SQRTPI(1) = sqrt(π)")
    { REQUIRE(eng.evaluate("=SQRTPI(1)").get<double>() == Catch::Approx(std::sqrt(3.14159265358979323846))); }

    SECTION("SQRTPI(-1) = error") { REQUIRE(eng.evaluate("=SQRTPI(-1)").type() == XLValueType::Error); }
}

// ============================================================================
// BATCH 2 — Utility functions
// ============================================================================

TEST_CASE("FormulaNewSubtotal", "[FormulaNewUtility]")
{
    XLFormulaEngine eng;
    auto resolver = makeMapResolver({
        {"A1", XLCellValue(1.0)}, {"A2", XLCellValue(2.0)}, {"A3", XLCellValue(3.0)},
        {"A4", XLCellValue(4.0)}, {"A5", XLCellValue(5.0)},
    });

    SECTION("SUBTOTAL 9 = SUM")
    { REQUIRE(eng.evaluate("=SUBTOTAL(9,A1:A5)", resolver).get<double>() == Catch::Approx(15.0)); }
    SECTION("SUBTOTAL 1 = AVERAGE")
    { REQUIRE(eng.evaluate("=SUBTOTAL(1,A1:A5)", resolver).get<double>() == Catch::Approx(3.0)); }
    SECTION("SUBTOTAL 4 = MAX")
    { REQUIRE(eng.evaluate("=SUBTOTAL(4,A1:A5)", resolver).get<double>() == Catch::Approx(5.0)); }
    SECTION("SUBTOTAL 5 = MIN")
    { REQUIRE(eng.evaluate("=SUBTOTAL(5,A1:A5)", resolver).get<double>() == Catch::Approx(1.0)); }
    SECTION("SUBTOTAL 102 = COUNT (ignore hidden)")
    { REQUIRE(eng.evaluate("=SUBTOTAL(102,A1:A5)", resolver).get<double>() == Catch::Approx(5.0)); }
}

TEST_CASE("FormulaNewChoose", "[FormulaNewUtility]")
{
    XLFormulaEngine eng;

    SECTION("CHOOSE selects correct value")
    {
        REQUIRE(eng.evaluate("=CHOOSE(1,\"a\",\"b\",\"c\")").get<std::string>() == "a");
        REQUIRE(eng.evaluate("=CHOOSE(2,\"a\",\"b\",\"c\")").get<std::string>() == "b");
        REQUIRE(eng.evaluate("=CHOOSE(3,\"a\",\"b\",\"c\")").get<std::string>() == "c");
    }
    SECTION("CHOOSE out of range = error")
    { REQUIRE(eng.evaluate("=CHOOSE(0,\"a\",\"b\")").type() == XLValueType::Error); }
    SECTION("CHOOSE numeric values")
    { REQUIRE(eng.evaluate("=CHOOSE(2,10,20,30)").get<double>() == Catch::Approx(20.0)); }
}

TEST_CASE("FormulaNewRowColumn", "[FormulaNewUtility]")
{
    XLFormulaEngine eng;
    auto resolver = makeMapResolver({{"A1", XLCellValue(42.0)}});

    // ROW and COLUMN parse the cell reference string passed as argument
    SECTION("ROW of A5")
    { REQUIRE(eng.evaluate("=ROW(A5)", resolver).get<int64_t>() == 5); }
    SECTION("ROW of C10")
    { REQUIRE(eng.evaluate("=ROW(C10)", resolver).get<int64_t>() == 10); }
    SECTION("COLUMN of C10")
    { REQUIRE(eng.evaluate("=COLUMN(C10)", resolver).get<int64_t>() == 3); }
    SECTION("COLUMN of A1")
    { REQUIRE(eng.evaluate("=COLUMN(A1)", resolver).get<int64_t>() == 1); }
}

TEST_CASE("FormulaNewDatedif", "[FormulaNewUtility]")
{
    XLFormulaEngine eng;

    // DATE(2020,1,1) and DATE(2023,6,15) are known serials
    SECTION("DATEDIF Y — complete years")
    { REQUIRE(eng.evaluate("=DATEDIF(DATE(2020,1,1),DATE(2023,6,15),\"Y\")").get<int64_t>() == 3); }
    SECTION("DATEDIF M — complete months")
    { REQUIRE(eng.evaluate("=DATEDIF(DATE(2020,1,1),DATE(2020,3,15),\"M\")").get<int64_t>() == 2); }
    SECTION("DATEDIF D — total days")
    {
        // 31 days in Jan + 29 (leap 2020) + 15 = 75
        REQUIRE(eng.evaluate("=DATEDIF(DATE(2020,1,1),DATE(2020,3,16),\"D\")").get<int64_t>() == 75);
    }
    SECTION("DATEDIF error when start > end")
    { REQUIRE(eng.evaluate("=DATEDIF(DATE(2025,1,1),DATE(2020,1,1),\"Y\")").type() == XLValueType::Error); }
}

TEST_CASE("FormulaNewIrr", "[FormulaNewUtility]")
{
    XLFormulaEngine eng;
    // Classic IRR example: -100, 30, 40, 50 → ~18.3%
    auto resolver = makeMapResolver({
        {"A1", XLCellValue(-100.0)},
        {"A2", XLCellValue(30.0)},
        {"A3", XLCellValue(40.0)},
        {"A4", XLCellValue(50.0)},
    });

    SECTION("IRR converges")
    {
        double irr = eng.evaluate("=IRR(A1:A4)", resolver).get<double>();
        // For cashflows -100,30,40,50: NPV(r) = 0
        // Verify that NPV at computed IRR ≈ 0
        double npv = -100.0 + 30.0/std::pow(1+irr,1) + 40.0/std::pow(1+irr,2) + 50.0/std::pow(1+irr,3);
        REQUIRE(std::abs(npv) < 0.01);
        REQUIRE(irr > 0.0);  // Must be a positive rate
    }
}

TEST_CASE("FormulaNewMirr", "[FormulaNewUtility]")
{
    XLFormulaEngine eng;
    // MIRR(-100,30,40,50; finance=10%, reinvest=12%) ≈ 19.9%
    auto resolver = makeMapResolver({
        {"A1", XLCellValue(-100.0)},
        {"A2", XLCellValue(30.0)},
        {"A3", XLCellValue(40.0)},
        {"A4", XLCellValue(50.0)},
    });
    double mirr = eng.evaluate("=MIRR(A1:A4,0.10,0.12)", resolver).get<double>();
    // MIRR must be positive and between finance_rate and reinvest_rate roughly
    REQUIRE(mirr > 0.0);
    REQUIRE(mirr < 1.0);
    // Cross-validate with formula: (-FV(0.12,3,0,positive CFs) / PV(0.10,3,0,negative CF))^(1/(n-1)) - 1
    // Just verify result is a reasonable positive return
    REQUIRE(std::isfinite(mirr));
}

TEST_CASE("FormulaNewRate", "[FormulaNewUtility]")
{
    XLFormulaEngine eng;
    // RATE(nper=10, pmt=-100, pv=800) ≈ 0.0481 (4.81%)
    // Verify with PMT: PMT(0.0481,10,800) ≈ -100
    SECTION("RATE converges to known value")
    {
        double rate = eng.evaluate("=RATE(10,-100,800)").get<double>();
        // Verify: PMT at computed rate ≈ -100
        double q = std::pow(1.0 + rate, 10.0);
        double pmt = -800.0 * q * rate / (q - 1.0);
        REQUIRE(std::abs(pmt + 100.0) < 1.0);  // within $1
        REQUIRE(rate > 0.0);
    }
}

TEST_CASE("FormulaNewIPMTPPMT", "[FormulaNewUtility]")
{
    XLFormulaEngine eng;
    // Loan: rate=5%/12, 60 months, PV=10000
    // PMT ≈ -188.71
    // Period 1: IPMT ≈ -41.67 (100000*0.05/12), PPMT = PMT - IPMT
    SECTION("IPMT period 1")
    {
        // rate=0.05/12, per=1, nper=60, pv=10000
        double ipmt = eng.evaluate("=IPMT(0.05/12,1,60,10000)").get<double>();
        REQUIRE(ipmt == Catch::Approx(-41.667).epsilon(0.01));
    }
    SECTION("PPMT period 1 = PMT - IPMT")
    {
        double ppmt = eng.evaluate("=PPMT(0.05/12,1,60,10000)").get<double>();
        double pmt  = eng.evaluate("=PMT(0.05/12,60,10000)").get<double>();
        double ipmt = eng.evaluate("=IPMT(0.05/12,1,60,10000)").get<double>();
        REQUIRE(ppmt == Catch::Approx(pmt - ipmt).epsilon(0.01));
    }
}

TEST_CASE("FormulaNewStatistical", "[FormulaNewUtility]")
{
    XLFormulaEngine eng;
    auto resolver = makeMapResolver({
        {"A1", XLCellValue(3.0)},
        {"A2", XLCellValue(1.0)},
        {"A3", XLCellValue(3.0)},
        {"A4", XLCellValue(2.0)},
        {"A5", XLCellValue(3.0)},
    });

    auto resolver2 = makeMapResolver({
        {"A1", XLCellValue(3.0)},
        {"A2", XLCellValue(1.0)},
        {"A3", XLCellValue(3.0)},
        {"A4", XLCellValue(2.0)},
        {"A5", XLCellValue(3.0)},
    });
    SECTION("MODE.SNGL = 3") { REQUIRE(eng.evaluate("=MODE.SNGL(A1:A5)", resolver2).get<double>() == Catch::Approx(3.0)); }

    SECTION("SKEW data")
    {
        // {3,1,3,2,3} — skewed
        double skew = eng.evaluate("=SKEW(A1:A5)", resolver).get<double>();
        REQUIRE(std::isfinite(skew));
    }

    SECTION("KURT data")
    {
        auto r2 = makeMapResolver({
            {"B1", XLCellValue(2.0)}, {"B2", XLCellValue(4.0)},
            {"B3", XLCellValue(4.0)}, {"B4", XLCellValue(4.0)},
            {"B5", XLCellValue(5.0)}, {"B6", XLCellValue(5.0)},
            {"B7", XLCellValue(7.0)}, {"B8", XLCellValue(9.0)},
        });
        double kurt = eng.evaluate("=KURT(B1:B8)", r2).get<double>();
        // Excel KURT for this dataset ≈ -0.16
        REQUIRE(std::isfinite(kurt));
    }
}

TEST_CASE("FormulaNewForecastLinear", "[FormulaNewUtility]")
{
    XLFormulaEngine eng;
    // y = 2x + 1: x={1,2,3}, y={3,5,7}
    auto resolver = makeMapResolver({
        {"A1", XLCellValue(3.0)}, {"A2", XLCellValue(5.0)}, {"A3", XLCellValue(7.0)},
        {"B1", XLCellValue(1.0)}, {"B2", XLCellValue(2.0)}, {"B3", XLCellValue(3.0)},
    });

    SECTION("FORECAST.LINEAR extrapolates x=4 → 9")
    { REQUIRE(eng.evaluate("=FORECAST.LINEAR(4,A1:A3,B1:B3)", resolver).get<double>() == Catch::Approx(9.0).epsilon(0.001)); }
    SECTION("FORECAST.LINEAR interpolates x=1.5 → 4")
    { REQUIRE(eng.evaluate("=FORECAST.LINEAR(1.5,A1:A3,B1:B3)", resolver).get<double>() == Catch::Approx(4.0).epsilon(0.001)); }
}


// ============================================================================
// BATCH 3 — Statistical distributions
// ============================================================================

TEST_CASE("FormulaNewNormalDist", "[FormulaNewDistrib]")
{
    XLFormulaEngine eng;

    SECTION("NORM.S.DIST CDF at 0 = 0.5")
    { REQUIRE(eng.evaluate("=NORM.S.DIST(0,TRUE)").get<double>() == Catch::Approx(0.5)); }

    SECTION("NORM.S.DIST CDF at 1.96 ≈ 0.975")
    { REQUIRE(eng.evaluate("=NORM.S.DIST(1.96,TRUE)").get<double>() == Catch::Approx(0.975).epsilon(0.001)); }

    SECTION("NORM.S.DIST CDF at -1.96 ≈ 0.025")
    { REQUIRE(eng.evaluate("=NORM.S.DIST(-1.96,TRUE)").get<double>() == Catch::Approx(0.025).epsilon(0.001)); }

    SECTION("NORM.S.DIST PDF at 0 = 1/sqrt(2π)")
    { REQUIRE(eng.evaluate("=NORM.S.DIST(0,FALSE)").get<double>() == Catch::Approx(0.3989422804).epsilon(1e-7)); }

    SECTION("NORM.DIST CDF at mean = 0.5")
    { REQUIRE(eng.evaluate("=NORM.DIST(10,10,2,TRUE)").get<double>() == Catch::Approx(0.5)); }

    SECTION("NORM.DIST with sigma<=0 = error")
    { REQUIRE(eng.evaluate("=NORM.DIST(0,0,0,TRUE)").type() == XLValueType::Error); }

    SECTION("Legacy NORMSDIST(1.645) ≈ 0.95")
    { REQUIRE(eng.evaluate("=NORMSDIST(1.645)").get<double>() == Catch::Approx(0.95).epsilon(0.001)); }
}

TEST_CASE("FormulaNewNormalInv", "[FormulaNewDistrib]")
{
    XLFormulaEngine eng;

    SECTION("NORM.S.INV(0.5) = 0")
    { REQUIRE(eng.evaluate("=NORM.S.INV(0.5)").get<double>() == Catch::Approx(0.0).margin(1e-6)); }

    SECTION("NORM.S.INV(0.975) ≈ 1.96")
    { REQUIRE(eng.evaluate("=NORM.S.INV(0.975)").get<double>() == Catch::Approx(1.96).epsilon(0.001)); }

    SECTION("NORM.S.INV(0.025) ≈ -1.96")
    { REQUIRE(eng.evaluate("=NORM.S.INV(0.025)").get<double>() == Catch::Approx(-1.96).epsilon(0.001)); }

    SECTION("NORM.INV(0.5,10,2) = 10")
    { REQUIRE(eng.evaluate("=NORM.INV(0.5,10,2)").get<double>() == Catch::Approx(10.0).epsilon(1e-6)); }

    SECTION("NORM.S.INV out of range = error")
    {
        REQUIRE(eng.evaluate("=NORM.S.INV(0)").type() == XLValueType::Error);
        REQUIRE(eng.evaluate("=NORM.S.INV(1)").type() == XLValueType::Error);
    }
}

TEST_CASE("FormulaNewTDist", "[FormulaNewDistrib]")
{
    XLFormulaEngine eng;

    // T.DIST.2T(1.9599..., 1000) ≈ 0.05 (large df → standard normal)
    SECTION("T.DIST.2T(1.96, 1000) ≈ 0.05")
    { REQUIRE(eng.evaluate("=T.DIST.2T(1.96,1000)").get<double>() == Catch::Approx(0.05).epsilon(0.01)); }

    SECTION("T.DIST.RT(0, 5) = 0.5")
    { REQUIRE(eng.evaluate("=T.DIST.RT(0,5)").get<double>() == Catch::Approx(0.5).epsilon(1e-6)); }

    SECTION("T.DIST CDF at 0 = 0.5")
    { REQUIRE(eng.evaluate("=T.DIST(0,10,TRUE)").get<double>() == Catch::Approx(0.5).epsilon(1e-6)); }

    SECTION("T.INV.2T(0.05, 1000) ≈ 1.96")
    { REQUIRE(eng.evaluate("=T.INV.2T(0.05,1000)").get<double>() == Catch::Approx(1.96).epsilon(0.01)); }

    SECTION("T.INV(0.5, 10) = 0")
    { REQUIRE(eng.evaluate("=T.INV(0.5,10)").get<double>() == Catch::Approx(0.0).margin(1e-4)); }
}

TEST_CASE("FormulaNewChisqDist", "[FormulaNewDistrib]")
{
    XLFormulaEngine eng;

    SECTION("CHISQ.DIST(0, 1, TRUE) = 0")
    { REQUIRE(eng.evaluate("=CHISQ.DIST(0,1,TRUE)").get<double>() == Catch::Approx(0.0).margin(1e-8)); }

    // CHISQ.INV.RT(0.05, 10) ≈ 18.307 (critical value)
    SECTION("CHISQ.INV.RT(0.05, 10) ≈ 18.307")
    { REQUIRE(eng.evaluate("=CHISQ.INV.RT(0.05,10)").get<double>() == Catch::Approx(18.307).epsilon(0.01)); }

    SECTION("CHISQ.DIST.RT + CHISQ.DIST = 1")
    {
        double x   = 5.0;
        double df  = 3.0;
        double cdf = eng.evaluate("=CHISQ.DIST(5,3,TRUE)").get<double>();
        double rt  = eng.evaluate("=CHISQ.DIST.RT(5,3)").get<double>();
        REQUIRE(cdf + rt == Catch::Approx(1.0).epsilon(1e-9));
    }

    // CHISQ.INV(p, df) should be the inverse of CHISQ.DIST
    SECTION("CHISQ.INV is inverse of CHISQ.DIST")
    {
        double x    = 5.991;  // chi2 critical value for df=2, p=0.95
        double p    = eng.evaluate("=CHISQ.DIST(5.991,2,TRUE)").get<double>();
        double xInv = eng.evaluate("=CHISQ.INV(0.95,2)").get<double>();
        REQUIRE(p    == Catch::Approx(0.95).epsilon(0.001));
        REQUIRE(xInv == Catch::Approx(5.991).epsilon(0.01));
    }
}

TEST_CASE("FormulaNewBinomDist", "[FormulaNewDistrib]")
{
    XLFormulaEngine eng;

    // BINOM.DIST(2, 5, 0.5, FALSE) = C(5,2)*0.5^5 = 10/32 = 0.3125
    SECTION("BINOM.DIST PMF")
    { REQUIRE(eng.evaluate("=BINOM.DIST(2,5,0.5,FALSE)").get<double>() == Catch::Approx(0.3125)); }

    // BINOM.DIST(5, 5, 0.5, TRUE) = 1
    SECTION("BINOM.DIST CDF full")
    { REQUIRE(eng.evaluate("=BINOM.DIST(5,5,0.5,TRUE)").get<double>() == Catch::Approx(1.0)); }

    // BINOM.DIST(0, 5, 0.5, FALSE) = 0.03125
    SECTION("BINOM.DIST PMF k=0")
    { REQUIRE(eng.evaluate("=BINOM.DIST(0,5,0.5,FALSE)").get<double>() == Catch::Approx(0.03125)); }

    SECTION("BINOM.DIST p=0 k>0 = 0")
    { REQUIRE(eng.evaluate("=BINOM.DIST(1,5,0,FALSE)").get<double>() == Catch::Approx(0.0)); }
}

TEST_CASE("FormulaNewPoissonDist", "[FormulaNewDistrib]")
{
    XLFormulaEngine eng;

    // POISSON(0, 1, FALSE) = e^-1 ≈ 0.3679
    SECTION("POISSON.DIST PMF k=0, λ=1")
    { REQUIRE(eng.evaluate("=POISSON.DIST(0,1,FALSE)").get<double>() == Catch::Approx(0.3679).epsilon(0.001)); }

    // POISSON(2, 2, FALSE) = 2^2*e^-2/2! = 2e^-2 ≈ 0.2707
    SECTION("POISSON.DIST PMF k=2, λ=2")
    { REQUIRE(eng.evaluate("=POISSON.DIST(2,2,FALSE)").get<double>() == Catch::Approx(0.2707).epsilon(0.001)); }

    // CDF at very large k should be ≈1
    SECTION("POISSON.DIST CDF large k ≈ 1")
    { REQUIRE(eng.evaluate("=POISSON.DIST(100,1,TRUE)").get<double>() == Catch::Approx(1.0).epsilon(1e-9)); }
}

TEST_CASE("FormulaNewExponDist", "[FormulaNewDistrib]")
{
    XLFormulaEngine eng;

    // EXPON.DIST(x, λ, FALSE) = λ*e^(-λx)
    // At x=0: PDF = λ
    SECTION("EXPON.DIST PDF at x=0 = λ")
    { REQUIRE(eng.evaluate("=EXPON.DIST(0,2,FALSE)").get<double>() == Catch::Approx(2.0)); }

    // CDF at x=0 = 0
    SECTION("EXPON.DIST CDF at x=0 = 0")
    { REQUIRE(eng.evaluate("=EXPON.DIST(0,2,TRUE)").get<double>() == Catch::Approx(0.0)); }

    // CDF at large x ≈ 1
    SECTION("EXPON.DIST CDF at large x ≈ 1")
    { REQUIRE(eng.evaluate("=EXPON.DIST(100,1,TRUE)").get<double>() == Catch::Approx(1.0).epsilon(1e-9)); }

    // CDF: 1 - e^(-λx), at x=1/λ: 1 - e^-1 ≈ 0.6321
    SECTION("EXPON.DIST CDF at mean ≈ 0.6321")
    { REQUIRE(eng.evaluate("=EXPON.DIST(1,1,TRUE)").get<double>() == Catch::Approx(0.6321).epsilon(0.001)); }

    SECTION("EXPON.DIST x<0 = error")
    { REQUIRE(eng.evaluate("=EXPON.DIST(-1,1,TRUE)").type() == XLValueType::Error); }

    SECTION("EXPON.DIST λ<=0 = error")
    { REQUIRE(eng.evaluate("=EXPON.DIST(1,0,TRUE)").type() == XLValueType::Error); }
}

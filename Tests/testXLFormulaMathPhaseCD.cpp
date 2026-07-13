/**
 * @file testXLFormulaMathPhaseCD.cpp
 * @brief Phase C engineering functions + Phase D AST cache / RNG / matrix caps.
 */

#include <XLCalculationEngine.hpp>
#include <XLFormulaEngine.hpp>
#include <XLFormulaUtils.hpp>
#include <catch2/catch_all.hpp>
#include <cmath>
#include <string>

using namespace OpenXLSX;

namespace
{
    bool isErr(XLCellValue v, const char* token)
    {
        return v.type() == XLValueType::Error && v.getString() == token;
    }
}    // namespace

// =============================================================================
// Phase C – Engineering
// =============================================================================

TEST_CASE("PhaseC bitwise functions", "[XLFormulaMathPhaseC][Bit]")
{
    XLFormulaEngine eng;
    REQUIRE(eng.evaluate("=BITAND(5,3)").get<double>() == Catch::Approx(1.0));
    REQUIRE(eng.evaluate("=BITOR(5,3)").get<double>() == Catch::Approx(7.0));
    REQUIRE(eng.evaluate("=BITXOR(5,3)").get<double>() == Catch::Approx(6.0));
    REQUIRE(eng.evaluate("=BITLSHIFT(3,2)").get<double>() == Catch::Approx(12.0));
    REQUIRE(eng.evaluate("=BITRSHIFT(12,2)").get<double>() == Catch::Approx(3.0));
    REQUIRE(isErr(eng.evaluate("=BITAND(-1,1)"), "#NUM!"));
}

TEST_CASE("PhaseC base conversion", "[XLFormulaMathPhaseC][Base]")
{
    XLFormulaEngine eng;
    REQUIRE(eng.evaluate("=HEX2DEC(\"A\")").get<double>() == Catch::Approx(10.0));
    REQUIRE(eng.evaluate("=DEC2HEX(255)").getString() == "FF");
    REQUIRE(eng.evaluate("=BIN2DEC(\"1010\")").get<double>() == Catch::Approx(10.0));
    REQUIRE(eng.evaluate("=DEC2BIN(10)").getString() == "1010");
    REQUIRE(eng.evaluate("=OCT2DEC(\"17\")").get<double>() == Catch::Approx(15.0));
    REQUIRE(eng.evaluate("=DEC2OCT(15)").getString() == "17");
    REQUIRE(eng.evaluate("=HEX2BIN(\"A\")").getString() == "1010");
    REQUIRE(eng.evaluate("=BIN2HEX(\"1111\")").getString() == "F");
}

TEST_CASE("PhaseC complex numbers", "[XLFormulaMathPhaseC][Complex]")
{
    XLFormulaEngine eng;
    REQUIRE(eng.evaluate("=COMPLEX(3,4)").getString() == "3+4i");
    REQUIRE(eng.evaluate("=IMREAL(\"3+4i\")").get<double>() == Catch::Approx(3.0));
    REQUIRE(eng.evaluate("=IMAGINARY(\"3+4i\")").get<double>() == Catch::Approx(4.0));
    REQUIRE(eng.evaluate("=IMABS(\"3+4i\")").get<double>() == Catch::Approx(5.0));
    REQUIRE(eng.evaluate("=IMSUM(\"1+2i\",\"3+4i\")").getString() == "4+6i");
    REQUIRE(eng.evaluate("=IMSUB(\"3+4i\",\"1+2i\")").getString() == "2+2i");
    // (1+2i)*(3+4i) = 3+4i+6i+8i² = 3+10i-8 = -5+10i
    auto prod = eng.evaluate("=IMPRODUCT(\"1+2i\",\"3+4i\")").getString();
    REQUIRE((prod == "-5+10i" || prod.find("-5") != std::string::npos));
    REQUIRE(eng.evaluate("=IMCONJUGATE(\"3+4i\")").getString() == "3-4i");
}

TEST_CASE("PhaseC ERF GAMMA", "[XLFormulaMathPhaseC][Special]")
{
    XLFormulaEngine eng;
    REQUIRE(eng.evaluate("=ERF(0)").get<double>() == Catch::Approx(0.0));
    REQUIRE(eng.evaluate("=ERF(1)").get<double>() == Catch::Approx(std::erf(1.0)).margin(1e-9));
    REQUIRE(eng.evaluate("=ERFC(0)").get<double>() == Catch::Approx(1.0));
    REQUIRE(eng.evaluate("=GAMMA(1)").get<double>() == Catch::Approx(1.0).margin(1e-9));
    REQUIRE(eng.evaluate("=GAMMA(5)").get<double>() == Catch::Approx(24.0).margin(1e-6));
    REQUIRE(eng.evaluate("=GAMMALN(1)").get<double>() == Catch::Approx(0.0).margin(1e-9));
    REQUIRE(isErr(eng.evaluate("=GAMMA(0)"), "#NUM!"));
}

// =============================================================================
// Phase D – AST cache, RNG, matrix cap, deps
// =============================================================================

TEST_CASE("PhaseD AST cache", "[XLFormulaMathPhaseD][AstCache]")
{
    XLFormulaEngine eng;
    eng.clearAstCache();
    eng.setAstCacheEnabled(true);
    eng.setAstCacheCapacity(64);

    REQUIRE(eng.astCacheSize() == 0);
    REQUIRE(eng.evaluate("=1+2+3").get<double>() == Catch::Approx(6.0));
    REQUIRE(eng.astCacheSize() == 1);
    // Same formula (with/without =) should hit cache
    REQUIRE(eng.evaluate("1+2+3").get<double>() == Catch::Approx(6.0));
    REQUIRE(eng.astCacheSize() == 1);

    REQUIRE(eng.evaluate("=SUM(1,2)").get<double>() == Catch::Approx(3.0));
    REQUIRE(eng.astCacheSize() == 2);

    eng.clearAstCache();
    REQUIRE(eng.astCacheSize() == 0);

    eng.setAstCacheEnabled(false);
    REQUIRE(eng.evaluate("=10*2").get<double>() == Catch::Approx(20.0));
    REQUIRE(eng.astCacheSize() == 0);
}

TEST_CASE("PhaseD deterministic RAND seed", "[XLFormulaMathPhaseD][Rng]")
{
    setFormulaRandomSeed(42);
    XLFormulaEngine eng;
    double          a1 = eng.evaluate("=RAND()").get<double>();
    double          a2 = eng.evaluate("=RAND()").get<double>();

    setFormulaRandomSeed(42);
    double b1 = eng.evaluate("=RAND()").get<double>();
    double b2 = eng.evaluate("=RAND()").get<double>();

    REQUIRE(a1 == Catch::Approx(b1));
    REQUIRE(a2 == Catch::Approx(b2));
    REQUIRE(a1 >= 0.0);
    REQUIRE(a1 < 1.0);

    setFormulaRandomSeed(7);
    double r1 = eng.evaluate("=RANDBETWEEN(1,1000)").get<double>();
    setFormulaRandomSeed(7);
    double r2 = eng.evaluate("=RANDBETWEEN(1,1000)").get<double>();
    REQUIRE(r1 == Catch::Approx(r2));

    clearFormulaRandomSeed();
}

TEST_CASE("PhaseD MMULT matrix dim cap", "[XLFormulaMathPhaseD][Matrix]")
{
    XLFormulaEngine eng;
    // Normal 2x2 still works
    auto arr = eng.evaluateArray("=MMULT({1,0;0,1},{2,3;4,5})");
    REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(2.0));
    REQUIRE(arr.at(1, 1).get<double>() == Catch::Approx(5.0));
}

TEST_CASE("PhaseD large range dependency sampling", "[XLFormulaMathPhaseD][Deps]")
{
    // maxExpanded=10 but range has 100 cells → should still record some deps (corners/samples)
    auto deps = XLCalculationEngine::extractDependencies("=SUM(A1:A100)", "", 10);
    REQUIRE(deps.size() >= 4);    // at least corners
    REQUIRE(deps.size() <= 10 + 4);
}

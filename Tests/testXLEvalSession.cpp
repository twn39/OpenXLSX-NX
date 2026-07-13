/**
 * @file testXLEvalSession.cpp
 * @brief Phase A formula-engine tests: XLEvalSession, INDIRECT, OFFSET, named ranges, ROW/COLUMN.
 */

#include <OpenXLSX.hpp>
#include <XLEvaluationContext.hpp>
#include <XLFormulaEngine.hpp>
#include <catch2/catch_all.hpp>
#include <map>
#include <utility>

using namespace OpenXLSX;

TEST_CASE("XLEvalSession INDIRECT scalar and range", "[XLEvalSession][INDIRECT]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(10.0)},
        {"A2", XLCellValue(20.0)},
        {"A3", XLCellValue(30.0)},
        {"B1", XLCellValue(std::string("A2"))},
    });
    auto resolver = XLFormulaEngine::makeResolver(ctx);

    SECTION("INDIRECT single cell")
    {
        REQUIRE(eng.evaluate("=INDIRECT(\"A1\")", resolver).get<double>() == Catch::Approx(10.0));
        REQUIRE(eng.evaluate("=INDIRECT(B1)", resolver).get<double>() == Catch::Approx(20.0));
    }

    SECTION("INDIRECT as range argument to SUM")
    {
        REQUIRE(eng.evaluate("=SUM(INDIRECT(\"A1:A3\"))", resolver).get<double>() == Catch::Approx(60.0));
    }

    SECTION("INDIRECT empty / bad ref")
    {
        REQUIRE(eng.evaluate("=INDIRECT(\"\")", resolver).type() == XLValueType::Error);
    }
}

TEST_CASE("XLEvalSession OFFSET", "[XLEvalSession][OFFSET]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(1.0)},
        {"A2", XLCellValue(2.0)},
        {"A3", XLCellValue(3.0)},
        {"B1", XLCellValue(10.0)},
        {"B2", XLCellValue(20.0)},
        {"C2", XLCellValue(99.0)},
    });
    auto resolver = XLFormulaEngine::makeResolver(ctx);

    SECTION("OFFSET single cell")
    {
        // From A1, down 1, right 0 → A2
        REQUIRE(eng.evaluate("=OFFSET(A1,1,0)", resolver).get<double>() == Catch::Approx(2.0));
        // From A1, down 1, right 1 → B2
        REQUIRE(eng.evaluate("=OFFSET(A1,1,1)", resolver).get<double>() == Catch::Approx(20.0));
    }

    SECTION("OFFSET range height/width with SUM")
    {
        // A1 start, height 3, width 1 → A1:A3
        REQUIRE(eng.evaluate("=SUM(OFFSET(A1,0,0,3,1))", resolver).get<double>() == Catch::Approx(6.0));
    }

    SECTION("OFFSET out of bounds")
    {
        REQUIRE(eng.evaluate("=OFFSET(A1,-1,0)", resolver).type() == XLValueType::Error);
    }
}

TEST_CASE("XLEvalSession named range resolution", "[XLEvalSession][Name]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(5.0)},
        {"A2", XLCellValue(7.0)},
        {"B1", XLCellValue(100.0)},
    });
    auto resolver = XLFormulaEngine::makeResolver(ctx);

    XLEvalSession session(resolver);
    session.setNameResolver([](std::string_view name) -> std::optional<std::string> {
        if (name == "MyRange") return std::string("A1:A2");
        if (name == "Anchor") return std::string("B1");
        return std::nullopt;
    });

    SECTION("name as scalar")
    {
        REQUIRE(eng.evaluate("=Anchor", session).get<double>() == Catch::Approx(100.0));
    }

    SECTION("name via INDIRECT")
    {
        REQUIRE(eng.evaluate("=SUM(INDIRECT(\"MyRange\"))", session).get<double>() == Catch::Approx(12.0));
    }
}

TEST_CASE("XLEvalSession parameterless ROW/COLUMN", "[XLEvalSession][ROW][COLUMN]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx;
    auto                   resolver = XLFormulaEngine::makeResolver(ctx);

    XLEvalSession session(resolver);
    session.setCurrentCell(/*row=*/12, /*col=*/3);    // C12

    REQUIRE(eng.evaluate("=ROW()", session).get<int64_t>() == 12);
    REQUIRE(eng.evaluate("=COLUMN()", session).get<int64_t>() == 3);
    REQUIRE(eng.evaluate("=ROW(A5)", session).get<int64_t>() == 5);
    REQUIRE(eng.evaluate("=COLUMN(C10)", session).get<int64_t>() == 3);
}

TEST_CASE("XLEvalSession depth guard for nested INDIRECT", "[XLEvalSession][INDIRECT]")
{
    XLFormulaEngine eng;
    // A1 → "A1" would recurse forever through INDIRECT without a depth guard.
    XLMapEvaluationContext ctx({{"A1", XLCellValue(std::string("A1"))}});
    auto                   resolver = XLFormulaEngine::makeResolver(ctx);

    // Nested INDIRECT(INDIRECT(...)) still bottoms out; circular ref text via depth.
    // Build a chain that exceeds depth: not practical with formula text alone.
    // Directly verify enterCall limit.
    XLEvalSession session(resolver);
    for (int i = 0; i < XLEvalSession::kMaxCallDepth; ++i) {
        REQUIRE(session.enterCall());
    }
    REQUIRE_FALSE(session.enterCall());
    for (int i = 0; i < XLEvalSession::kMaxCallDepth; ++i) session.leaveCall();

    // Nested INDIRECT: each level coerces the inner ref to its cell *value* (text address).
    // Z1 → "Z2", Z2 → 7  ⇒  INDIRECT(INDIRECT("Z1")) ≡ INDIRECT("Z2") ≡ 7
    XLMapEvaluationContext ctx2({
        {"Z1", XLCellValue(std::string("Z2"))},
        {"Z2", XLCellValue(7.0)},
    });
    auto r2 = XLFormulaEngine::makeResolver(ctx2);
    REQUIRE(eng.evaluate("=INDIRECT(INDIRECT(\"Z1\"))", r2).get<double>() == Catch::Approx(7.0));
}

/**
 * Critical path: when INDIRECT is an argument to another function that needs a range
 * (SUM), expandArg must keep LazyRange shape — not collapse to a scalar first.
 */
TEST_CASE("INDIRECT range shape preserved for aggregations", "[XLEvalSession][INDIRECT]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"D1", XLCellValue(1.0)},
        {"D2", XLCellValue(2.0)},
        {"D3", XLCellValue(3.0)},
        {"E1", XLCellValue(std::string("D1:D3"))},
    });
    auto resolver = XLFormulaEngine::makeResolver(ctx);
    REQUIRE(eng.evaluate("=SUM(INDIRECT(E1))", resolver).get<double>() == Catch::Approx(6.0));
}

TEST_CASE("Backward compatible evaluate without session", "[XLEvalSession]")
{
    XLFormulaEngine eng;
    REQUIRE(eng.evaluate("=1+2").get<double>() == Catch::Approx(3.0));
    REQUIRE(eng.evaluate("=SUM(1,2,3)").get<double>() == Catch::Approx(6.0));
}

TEST_CASE("Phase B element-wise range arithmetic", "[XLEvalSession][ArrayBroadcast]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(1.0)},
        {"A2", XLCellValue(2.0)},
        {"A3", XLCellValue(3.0)},
        {"B1", XLCellValue(10.0)},
        {"B2", XLCellValue(20.0)},
        {"B3", XLCellValue(30.0)},
    });
    auto resolver = XLFormulaEngine::makeResolver(ctx);

    SECTION("scalar broadcast into SUM")
    {
        REQUIRE(eng.evaluate("=SUM(A1:A3*2)", resolver).get<double>() == Catch::Approx(12.0));
        REQUIRE(eng.evaluate("=SUM(A1:A3+1)", resolver).get<double>() == Catch::Approx(9.0));
    }

    SECTION("pairwise range + range")
    {
        REQUIRE(eng.evaluate("=SUM(A1:A3+B1:B3)", resolver).get<double>() == Catch::Approx(66.0));
    }

    SECTION("implicit intersection for bare range arithmetic")
    {
        // Scalar formula context returns the first element of the broadcast result.
        REQUIRE(eng.evaluate("=A1:A3*2", resolver).get<double>() == Catch::Approx(2.0));
    }

    SECTION("evaluateArray preserves full shape")
    {
        auto arr = eng.evaluateArray("=A1:A3*2", resolver);
        REQUIRE(arr.rows() == 3);
        REQUIRE(arr.cols() == 1);
        REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(2.0));
        REQUIRE(arr.at(1, 0).get<double>() == Catch::Approx(4.0));
        REQUIRE(arr.at(2, 0).get<double>() == Catch::Approx(6.0));
    }
}

TEST_CASE("Phase B FILTER UNIQUE SORT SEQUENCE TRANSPOSE", "[XLEvalSession][DynamicArray]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(1.0)},
        {"A2", XLCellValue(2.0)},
        {"A3", XLCellValue(3.0)},
        {"A4", XLCellValue(2.0)},
        {"B1", XLCellValue(true)},
        {"B2", XLCellValue(false)},
        {"B3", XLCellValue(true)},
        {"B4", XLCellValue(true)},
        {"C1", XLCellValue(30.0)},
        {"C2", XLCellValue(10.0)},
        {"C3", XLCellValue(20.0)},
        {"D1", XLCellValue(std::string("x"))},
        {"D2", XLCellValue(std::string("y"))},
        {"D3", XLCellValue(std::string("x"))},
        {"E1", XLCellValue(1.0)},
        {"E2", XLCellValue(2.0)},
        {"F1", XLCellValue(3.0)},
        {"F2", XLCellValue(4.0)},
    });
    auto resolver = XLFormulaEngine::makeResolver(ctx);

    SECTION("FILTER rows")
    {
        auto arr = eng.evaluateArray("=FILTER(A1:A4,B1:B4)", resolver);
        REQUIRE(arr.rows() == 3);
        REQUIRE(arr.cols() == 1);
        REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(1.0));
        REQUIRE(arr.at(1, 0).get<double>() == Catch::Approx(3.0));
        REQUIRE(arr.at(2, 0).get<double>() == Catch::Approx(2.0));
    }

    SECTION("FILTER with if_empty")
    {
        auto arr = eng.evaluateArray("=FILTER(A1:A1,B2:B2,\"none\")", resolver);
        REQUIRE(arr.asScalar().get<std::string>() == "none");
    }

    SECTION("UNIQUE")
    {
        auto arr = eng.evaluateArray("=UNIQUE(A1:A4)", resolver);
        REQUIRE(arr.rows() == 3);
        REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(1.0));
        REQUIRE(arr.at(1, 0).get<double>() == Catch::Approx(2.0));
        REQUIRE(arr.at(2, 0).get<double>() == Catch::Approx(3.0));
    }

    SECTION("SORT ascending by column 1")
    {
        auto arr = eng.evaluateArray("=SORT(C1:C3)", resolver);
        REQUIRE(arr.rows() == 3);
        REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(10.0));
        REQUIRE(arr.at(1, 0).get<double>() == Catch::Approx(20.0));
        REQUIRE(arr.at(2, 0).get<double>() == Catch::Approx(30.0));
    }

    SECTION("SORT descending")
    {
        auto arr = eng.evaluateArray("=SORT(C1:C3,1,-1)", resolver);
        REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(30.0));
        REQUIRE(arr.at(2, 0).get<double>() == Catch::Approx(10.0));
    }

    SECTION("SEQUENCE")
    {
        auto arr = eng.evaluateArray("=SEQUENCE(2,3,10,1)", resolver);
        REQUIRE(arr.rows() == 2);
        REQUIRE(arr.cols() == 3);
        REQUIRE(arr.at(0, 0).get<int64_t>() == 10);
        REQUIRE(arr.at(0, 1).get<int64_t>() == 11);
        REQUIRE(arr.at(1, 0).get<int64_t>() == 13);
        REQUIRE(arr.at(1, 2).get<int64_t>() == 15);
    }

    SECTION("TRANSPOSE")
    {
        auto arr = eng.evaluateArray("=TRANSPOSE(E1:F2)", resolver);
        REQUIRE(arr.rows() == 2);
        REQUIRE(arr.cols() == 2);
        REQUIRE(arr.at(0, 0).get<double>() == Catch::Approx(1.0));
        REQUIRE(arr.at(0, 1).get<double>() == Catch::Approx(2.0));
        REQUIRE(arr.at(1, 0).get<double>() == Catch::Approx(3.0));
        REQUIRE(arr.at(1, 1).get<double>() == Catch::Approx(4.0));
    }

    SECTION("spillArray writes rectangular block")
    {
        auto arr = eng.evaluateArray("=SEQUENCE(2,2,1,1)", resolver);
        std::map<std::pair<uint32_t, uint16_t>, XLCellValue> grid;
        size_t n = spillArray(arr, 5, 3, [&](uint32_t r, uint16_t c, const XLCellValue& v) {
            grid[{r, c}] = v;
        });
        REQUIRE(n == 4);
        REQUIRE(grid[{5, 3}].get<int64_t>() == 1);
        REQUIRE(grid[{5, 4}].get<int64_t>() == 2);
        REQUIRE(grid[{6, 3}].get<int64_t>() == 3);
        REQUIRE(grid[{6, 4}].get<int64_t>() == 4);
    }

    SECTION("scalar evaluate of array fn is top-left")
    {
        REQUIRE(eng.evaluate("=SEQUENCE(3,1,5,1)", resolver).get<int64_t>() == 5);
        REQUIRE(eng.evaluate("=SORT(C1:C3)", resolver).get<double>() == Catch::Approx(10.0));
    }
}

/**
 * @file testXLFormulaDynamicArray2.cpp
 * @brief Dynamic array batch 2: SORTBY, RANDARRAY, TOCOL/TOROW, HSTACK/VSTACK, #SPILL!.
 */

#include <XLEvaluationContext.hpp>
#include <XLFormulaEngine.hpp>
#include <catch2/catch_all.hpp>
#include <map>
#include <set>
#include <utility>

using namespace OpenXLSX;

TEST_CASE("SORTBY sorts rows by companion array", "[XLFormulaDynamicArray2][SORTBY]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(std::string("x"))},
        {"A2", XLCellValue(std::string("y"))},
        {"A3", XLCellValue(std::string("z"))},
        {"B1", XLCellValue(30.0)},
        {"B2", XLCellValue(10.0)},
        {"B3", XLCellValue(20.0)},
    });
    auto r = XLFormulaEngine::makeResolver(ctx);

    auto arr = eng.evaluateArray("=SORTBY(A1:A3,B1:B3)", r);
    REQUIRE(arr.rows() == 3);
    REQUIRE(arr.at(0, 0).get<std::string>() == "y");    // 10
    REQUIRE(arr.at(1, 0).get<std::string>() == "z");    // 20
    REQUIRE(arr.at(2, 0).get<std::string>() == "x");    // 30

    auto desc = eng.evaluateArray("=SORTBY(A1:A3,B1:B3,-1)", r);
    REQUIRE(desc.at(0, 0).get<std::string>() == "x");
}

TEST_CASE("RANDARRAY shape and integer mode", "[XLFormulaDynamicArray2][RANDARRAY]")
{
    XLFormulaEngine eng;
    auto            arr = eng.evaluateArray("=RANDARRAY(2,3,1,1,TRUE)");
    REQUIRE(arr.rows() == 2);
    REQUIRE(arr.cols() == 3);
    for (size_t i = 0; i < arr.size(); ++i) {
        REQUIRE(arr[i].get<int64_t>() == 1);
    }

    auto real = eng.evaluateArray("=RANDARRAY(1,5,0,1,FALSE)");
    REQUIRE(real.cols() == 5);
    for (size_t i = 0; i < real.size(); ++i) {
        double v = real[i].get<double>();
        REQUIRE(v >= 0.0);
        REQUIRE(v <= 1.0);
    }
}

TEST_CASE("TOCOL TOROW HSTACK VSTACK", "[XLFormulaDynamicArray2][Stack]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({
        {"A1", XLCellValue(1.0)},
        {"B1", XLCellValue(2.0)},
        {"A2", XLCellValue(3.0)},
        {"B2", XLCellValue(4.0)},
        {"C1", XLCellValue(9.0)},
        {"C2", XLCellValue(8.0)},
    });
    auto r = XLFormulaEngine::makeResolver(ctx);

    SECTION("TOCOL row-major")
    {
        auto col = eng.evaluateArray("=TOCOL(A1:B2)", r);
        REQUIRE(col.rows() == 4);
        REQUIRE(col.cols() == 1);
        REQUIRE(col.at(0, 0).get<double>() == Catch::Approx(1.0));
        REQUIRE(col.at(1, 0).get<double>() == Catch::Approx(2.0));
        REQUIRE(col.at(2, 0).get<double>() == Catch::Approx(3.0));
        REQUIRE(col.at(3, 0).get<double>() == Catch::Approx(4.0));
    }

    SECTION("TOROW")
    {
        auto row = eng.evaluateArray("=TOROW(A1:B2)", r);
        REQUIRE(row.rows() == 1);
        REQUIRE(row.cols() == 4);
        REQUIRE(row.at(0, 3).get<double>() == Catch::Approx(4.0));
    }

    SECTION("HSTACK")
    {
        auto h = eng.evaluateArray("=HSTACK(A1:B2,C1:C2)", r);
        REQUIRE(h.rows() == 2);
        REQUIRE(h.cols() == 3);
        REQUIRE(h.at(0, 2).get<double>() == Catch::Approx(9.0));
        REQUIRE(h.at(1, 2).get<double>() == Catch::Approx(8.0));
    }

    SECTION("VSTACK")
    {
        auto v = eng.evaluateArray("=VSTACK(A1:B1,A2:B2)", r);
        REQUIRE(v.rows() == 2);
        REQUIRE(v.cols() == 2);
        REQUIRE(v.at(1, 0).get<double>() == Catch::Approx(3.0));
    }
}

TEST_CASE("spillArrayChecked #SPILL!", "[XLFormulaDynamicArray2][Spill]")
{
    XLFormulaEngine eng;
    auto            arr = eng.evaluateArray("=SEQUENCE(2,2)");

    std::map<std::pair<uint32_t, uint16_t>, XLCellValue> grid;
    // Pre-occupy C6 (would be part of spill from B5: 2x2 covers B5,C5,B6,C6)
    grid[{6, 3}] = XLCellValue(99.0);    // C6

    auto isOcc = [&](uint32_t r, uint16_t c) {
        auto it = grid.find({r, c});
        return it != grid.end() && it->second.type() != XLValueType::Empty;
    };

    auto write = [&](uint32_t r, uint16_t c, const XLCellValue& v) { grid[{r, c}] = v; };

    // Anchor B5 (row 5, col 2) — spill 2x2 hits C6 which is occupied
    auto blocked = spillArrayChecked(arr, 5, 2, write, isOcc, 5, 2);
    REQUIRE_FALSE(blocked.ok);
    REQUIRE(blocked.error.type() == XLValueType::Error);
    REQUIRE(blocked.error.getString().find("SPILL") != std::string::npos);
    REQUIRE(blocked.cellsWritten == 0);

    // Clear blocker
    grid.erase({6, 3});
    auto ok = spillArrayChecked(arr, 5, 2, write, isOcc, 5, 2);
    REQUIRE(ok.ok);
    REQUIRE(ok.cellsWritten == 4);
    REQUIRE(grid[{5, 2}].get<int64_t>() == 1);
    REQUIRE(grid[{6, 3}].get<int64_t>() == 4);

    REQUIRE(spillRangeIsClear(arr, 10, 1, isOcc, 10, 1));
}

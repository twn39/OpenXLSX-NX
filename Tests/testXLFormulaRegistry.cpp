/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "XLFormulaRegistry.hpp"
#include "XLFormulaFunction.hpp"
#include "XLFormulaUtils.hpp"

using namespace OpenXLSX;

namespace {
    // Custom test function to register
    XLCellValue fnCustomAdd(const std::vector<XLFormulaArg>& args) {
        if (args.size() < 2) return errValue();
        double a = toDouble(args[0][0]);
        double b = toDouble(args[1][0]);
        return XLCellValue(a + b);
    }
}

TEST_CASE("XLFormulaRegistryTests", "[XLFormulaRegistry]")
{
    SECTION("Lookup builtin functions")
    {
        const auto& registry = XLFormulaRegistry::getInstance();
        
        // Verifying that some key math and text functions are registered
        const auto* sumFunc = registry.lookup("SUM");
        REQUIRE(sumFunc != nullptr);

        const auto* avgFunc = registry.lookup("AVERAGE");
        REQUIRE(avgFunc != nullptr);

        const auto* concatFunc = registry.lookup("CONCATENATE");
        REQUIRE(concatFunc != nullptr);

        // Lowercase input should be normalized/found correctly
        const auto* sumFuncLower = registry.lookup("sum");
        REQUIRE(sumFuncLower != nullptr);
        REQUIRE(sumFuncLower == sumFunc);
    }

    SECTION("Custom function registration and execution")
    {
        auto& registry = XLFormulaRegistry::getInstance();
        
        // Register custom function
        auto customFunc = std::make_shared<XLSimpleFormulaFunction>(fnCustomAdd);
        registry.registerFunction("CUSTOM.ADD", customFunc);

        // Lookup the custom function
        const auto* retrieved = registry.lookup("CUSTOM.ADD");
        REQUIRE(retrieved != nullptr);

        // Evaluate custom function
        std::vector<XLFormulaArg> args;
        args.emplace_back(XLCellValue(10.0));
        args.emplace_back(XLCellValue(5.5));

        XLEvalSession session;
        XLFormulaArg  result = retrieved->execute(args, session);
        REQUIRE(result.asScalar().type() == XLValueType::Float);
        REQUIRE(result.asScalar().get<double>() == Catch::Approx(15.5));
    }
}

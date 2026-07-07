#include "XLFormulaFunctionsInternal.hpp"
#include "XLFormulaRegistry.hpp"
#include "XLDateTime.hpp"
#include "XLNumberFormatter.hpp"
#include "XLFormulaUtils.hpp"
#include <algorithm>
#include <array>
#include <cmath>
#include <numeric>
#include <random>
#include <regex>
#include <unordered_set>
#include <cctype>

using namespace OpenXLSX;


XLCellValue Formula::fnIf(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    const auto& cond = args[0][0];
    bool        test = false;
    if (cond.type() == XLValueType::Boolean)
        test = cond.get<bool>();
    else if (isNumeric(cond))
        test = (toDouble(cond) != 0.0);
    else if (cond.type() == XLValueType::String)
        test = !cond.get<std::string>().empty();

    if (test)
        return (args.size() > 1 && !args[1].empty()) ? args[1][0] : XLCellValue(true);
    else
        return (args.size() > 2 && !args[2].empty()) ? args[2][0] : XLCellValue(false);
}

XLCellValue Formula::fnIfs(const std::vector<XLFormulaArg>& args)
{
    // IFS(condition1, value1, [condition2, value2], ...)
    if (args.size() % 2 != 0 || args.empty()) return errValue();    // Requires pairs

    for (std::size_t i = 0; i < args.size(); i += 2) {
        if (args[i].empty()) return errValue();
        const auto& cond = args[i][0];
        bool        test = false;
        if (cond.type() == XLValueType::Boolean)
            test = cond.get<bool>();
        else if (isNumeric(cond))
            test = (toDouble(cond) != 0.0);
        else if (cond.type() == XLValueType::String)
            test = !cond.get<std::string>().empty();

        if (test) { return args[i + 1].empty() ? XLCellValue() : args[i + 1][0]; }
    }
    return errNA();    // If no condition is met, Excel returns #N/A
}

XLCellValue Formula::fnSwitch(const std::vector<XLFormulaArg>& args)
{
    // SWITCH(expression, value1, result1, [default_or_value2, result2], ... [default])
    if (args.size() < 3) return errValue();

    if (args[0].empty()) return errValue();
    const auto& expr = args[0][0];

    std::size_t i = 1;
    while (i + 1 < args.size()) {
        if (args[i].empty()) return errValue();
        const auto& val = args[i][0];

        bool match = false;
        if (isNumeric(expr) && isNumeric(val))
            match = (toDouble(expr) == toDouble(val));
        else if (expr.type() == XLValueType::String && val.type() == XLValueType::String)
            match = (toString(expr) == toString(val));
        else if (expr.type() == XLValueType::Boolean && val.type() == XLValueType::Boolean)
            match = (expr.get<bool>() == val.get<bool>());

        if (match) { return args[i + 1].empty() ? XLCellValue() : args[i + 1][0]; }
        i += 2;
    }

    // Check if there is a default value at the end
    if (i < args.size()) { return args[i].empty() ? XLCellValue() : args[i][0]; }

    return errNA();
}

XLCellValue Formula::fnAnd(const std::vector<XLFormulaArg>& args)
{
    for (const auto& arg : args) {
        for (const auto& v : arg) {
            if (!isNumeric(v) && v.type() != XLValueType::Boolean) return errValue();
            if (!toDouble(v)) return XLCellValue(false);
        }
    }
    return XLCellValue(true);
}

XLCellValue Formula::fnOr(const std::vector<XLFormulaArg>& args)
{
    for (const auto& arg : args) {
        for (const auto& v : arg) {
            if (!isNumeric(v) && v.type() != XLValueType::Boolean) return errValue();
            if (toDouble(v)) return XLCellValue(true);
        }
    }
    return XLCellValue(false);
}

XLCellValue Formula::fnNot(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    const auto& v = args[0][0];
    if (!isNumeric(v) && v.type() != XLValueType::Boolean) return errValue();
    return XLCellValue(!static_cast<bool>(toDouble(v)));
}

XLCellValue Formula::fnIferror(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    return isError(args[0][0]) ? args[1][0] : args[0][0];
}

XLCellValue Formula::fnTrue(const std::vector<XLFormulaArg>& /*args*/) { return XLCellValue(true); }

XLCellValue Formula::fnFalse(const std::vector<XLFormulaArg>& /*args*/) { return XLCellValue(false); }

XLCellValue Formula::fnChoose(const std::vector<XLFormulaArg>& args)
{
    // CHOOSE(index_num, value1, value2, ...)
    if (args.size() < 2 || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    auto   idx = static_cast<std::size_t>(toDouble(args[0][0]));
    if (idx < 1 || idx >= args.size()) return errValue();
    return args[idx].empty() ? XLCellValue{} : args[idx][0];
}

void OpenXLSX::registerLogicalFunctions(XLFormulaRegistry& registry)
{
    registry.registerFunction("IF", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIf));
    registry.registerFunction("IFS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIfs));
    registry.registerFunction("SWITCH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSwitch));
    registry.registerFunction("AND", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAnd));
    registry.registerFunction("OR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnOr));
    registry.registerFunction("NOT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnNot));
    registry.registerFunction("IFERROR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIferror));
    registry.registerFunction("TRUE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTrue));
    registry.registerFunction("FALSE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnFalse));
    registry.registerFunction("CHOOSE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnChoose));
}

// ===== External Includes ===== //
#include <algorithm>
#include <cassert>
#include <cctype>
#include <cmath>
#include <ctime>
#include <fmt/format.h>
#include <functional>
#include <numeric>
#include <string>
#include <unordered_map>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLCellReference.hpp"
#include "XLDateTime.hpp"
#include "XLException.hpp"
#include "XLFormulaEngine.hpp"
#include "XLFormulaRegistry.hpp"
#include "XLWorksheet.hpp"

using namespace OpenXLSX;

// =============================================================================
#include "XLFormulaUtils.hpp"

// =============================================================================
// Lexer
// =============================================================================

XLFormulaArg XLFormulaEngine::expandArg(const XLASTNode& argNode, const XLCellResolver& resolver) const
{
    if (argNode.kind == XLNodeKind::Range) return expandRange(argNode.text, resolver);

    // Single cell reference: create a 1x1 LazyRange so position metadata (row/col) is preserved.
    // Functions like ROW() and COLUMN() need the address, not just the resolved value.
    // Other functions that iterate values will call operator[0] which resolves via the resolver.
    if (argNode.kind == XLNodeKind::CellRef) return expandRange(argNode.text, resolver);

    // Evaluate normally and wrap in a single-element scalar
    return XLFormulaArg(evalNode(argNode, resolver));
}

// =============================================================================
// Evaluator – evalNode
// =============================================================================

XLCellValue XLFormulaEngine::evalNode(const XLASTNode& node, const XLCellResolver& resolver) const
{
    switch (node.kind) {
        case XLNodeKind::Number:
            return XLCellValue(node.number);
        case XLNodeKind::StringLit:
            return XLCellValue(node.text);
        case XLNodeKind::BoolLit:
            return XLCellValue(node.boolean);
        case XLNodeKind::ErrorLit: {
            XLCellValue e;
            e.setError(node.text);
            return e;
        }

        case XLNodeKind::CellRef: {
            if (!resolver) return XLCellValue{};
            return resolver(node.text);
        }

        case XLNodeKind::Range: {
            // Range used as scalar = first cell value
            auto vals = expandRange(node.text, resolver);
            return vals.empty() ? XLCellValue{} : vals[0];
        }

        case XLNodeKind::UnaryOp: {
            Expects(node.children.size() == 1);
            auto val = evalNode(*node.children[0], resolver);
            if (node.op == XLTokenKind::Minus) {
                if (!isNumeric(val)) return errValue();
                double d = toDouble(val);
                if (val.type() == XLValueType::Integer) return XLCellValue(static_cast<int64_t>(-d));
                return XLCellValue(-d);
            }
            if (node.op == XLTokenKind::Percent) {
                if (!isNumeric(val)) return errValue();
                return XLCellValue(toDouble(val) / 100.0);
            }
            return val;
        }

        case XLNodeKind::BinOp: {
            Expects(node.children.size() == 2);

            // String concat – evaluate early, no numeric coercion
            if (node.op == XLTokenKind::Amp) {
                auto lv = evalNode(*node.children[0], resolver);
                auto rv = evalNode(*node.children[1], resolver);
                if (isError(lv)) return lv;
                if (isError(rv)) return rv;
                return XLCellValue(toString(lv) + toString(rv));
            }

            auto lv = evalNode(*node.children[0], resolver);
            auto rv = evalNode(*node.children[1], resolver);
            if (isError(lv)) return lv;
            if (isError(rv)) return rv;

            // Arithmetic operators
            if (node.op == XLTokenKind::Plus || node.op == XLTokenKind::Minus || node.op == XLTokenKind::Star ||
                node.op == XLTokenKind::Slash || node.op == XLTokenKind::Caret)
            {
                if (!isNumeric(lv) || !isNumeric(rv)) return errValue();
                double l = toDouble(lv), r = toDouble(rv);
                switch (node.op) {
                    case XLTokenKind::Plus:
                        return XLCellValue(l + r);
                    case XLTokenKind::Minus:
                        return XLCellValue(l - r);
                    case XLTokenKind::Star:
                        return XLCellValue(l * r);
                    case XLTokenKind::Slash:
                        if (r == 0.0) return errDiv0();
                        return XLCellValue(l / r);
                    case XLTokenKind::Caret:
                        return XLCellValue(std::pow(l, r));
                    default:
                        break;
                }
            }

            // Comparison operators
            {
                bool result = false;
                // Numeric comparison
                if (isNumeric(lv) && isNumeric(rv)) {
                    double l = toDouble(lv), r = toDouble(rv);
                    switch (node.op) {
                        case XLTokenKind::Eq:
                            result = (l == r);
                            break;
                        case XLTokenKind::NEq:
                            result = (l != r);
                            break;
                        case XLTokenKind::Lt:
                            result = (l < r);
                            break;
                        case XLTokenKind::Le:
                            result = (l <= r);
                            break;
                        case XLTokenKind::Gt:
                            result = (l > r);
                            break;
                        case XLTokenKind::Ge:
                            result = (l >= r);
                            break;
                        default:
                            return errValue();
                    }
                }
                else {
                    // String comparison (case-insensitive like Excel)
                    std::string ls = toString(lv), rs = toString(rv);
                    std::transform(ls.begin(), ls.end(), ls.begin(), ::tolower);
                    std::transform(rs.begin(), rs.end(), rs.begin(), ::tolower);
                    switch (node.op) {
                        case XLTokenKind::Eq:
                            result = (ls == rs);
                            break;
                        case XLTokenKind::NEq:
                            result = (ls != rs);
                            break;
                        case XLTokenKind::Lt:
                            result = (ls < rs);
                            break;
                        case XLTokenKind::Le:
                            result = (ls <= rs);
                            break;
                        case XLTokenKind::Gt:
                            result = (ls > rs);
                            break;
                        case XLTokenKind::Ge:
                            result = (ls >= rs);
                            break;
                        default:
                            return errValue();
                    }
                }
                return XLCellValue(result);
            }
        }

        case XLNodeKind::FuncCall: {
            const auto& funcs = getBuiltins();
            auto        it    = funcs.find(node.text);
            if (it == funcs.end()) return errName();

            // Build per-arg vectors (ranges are expanded, scalars wrapped)
            std::vector<XLFormulaArg> argVecs;
            argVecs.reserve(node.children.size());
            for (const auto& child : node.children) argVecs.push_back(expandArg(*child, resolver));

            try {
                return it->second(argVecs);
            }
            catch (const std::exception& ex) {
                XLCellValue e;
                e.setError(std::string("#ERROR: ") + ex.what());
                return e;
            }
        }

        default:
            return errValue();
    }
}

// =============================================================================
// Evaluator – public evaluate()
// =============================================================================

XLCellValue XLFormulaEngine::evaluate(std::string_view formula, const XLCellResolver& resolver, XLFormulaDiagnosticReporter* reporter) const
{
    if (formula.empty()) return XLCellValue{};
    try {
        if (reporter) {
            *reporter = XLFormulaDiagnosticReporter(std::string(formula));
        }
        auto tokens = XLFormulaLexer::tokenize(formula);
        auto ast    = XLFormulaParser::parse(gsl::span<const XLToken>(tokens), reporter);
        return evalNode(*ast, resolver);
    }
    catch (const XLException&) {
        throw;
    }
    catch (const std::exception& ex) {
        XLCellValue e;
        e.setError(std::string("#ERROR: ") + ex.what());
        return e;
    }
}

// =============================================================================
// makeResolver
// =============================================================================

XLCellResolver XLFormulaEngine::makeResolver(const XLWorksheet& wks)
{
    return [&wks](std::string_view ref) -> XLCellValue {
        try {
            // Strip sheet prefix (e.g. "Sheet1!A1" -> "A1")
            auto        bang     = ref.find('!');
            std::string cellAddr = std::string(bang != std::string_view::npos ? ref.substr(bang + 1) : ref);
            return wks.cell(cellAddr).value();
        }
        catch (...) {
            return XLCellValue{};
        }
    };
}

// =============================================================================
// Built-in function registrations
// =============================================================================

XLFormulaEngine::XLFormulaEngine() = default;

const std::unordered_map<std::string, XLFormulaEngine::FuncImpl>& XLFormulaEngine::getBuiltins()
{
    static const std::unordered_map<std::string, FuncImpl> builtins = []() {
        std::unordered_map<std::string, FuncImpl> map;
        const auto& registry = XLFormulaRegistry::getInstance();

        // 1. Register all functions from the registry
        for (const auto& pair : registry.getFunctions()) {
            const auto& name = pair.first;
            const auto& func = pair.second;
            map[name] = [func](const std::vector<XLFormulaArg>& args) {
                return func->execute(args);
            };
        }

        // Helper function to safely register an alias from another registered function
        auto addAlias = [&](const std::string& alias, const std::string& target) {
            auto it = map.find(target);
            if (it != map.end()) {
                map[alias] = it->second;
            }
        };

        // 2. Register all standard aliases
        addAlias("AVG", "AVERAGE");
        addAlias("CONCAT", "CONCATENATE");
        addAlias("CEIL", "CEILING");
        addAlias("RANK.EQ", "RANK");
        addAlias("STDEV.S", "STDEV");
        addAlias("VAR.S", "VAR");
        addAlias("VARP", "VAR.P");
        addAlias("STDEVP", "STDEV.P");
        addAlias("CORREL", "PEARSON");
        addAlias("COVAR", "COVARIANCE.P");
        addAlias("PERCENTILE", "PERCENTILE.INC");
        addAlias("QUARTILE", "QUARTILE.INC");
        addAlias("MODE", "MODE.SNGL");
        addAlias("FORECAST", "FORECAST.LINEAR");
        addAlias("NORMSDIST", "NORM.S.DIST");
        addAlias("NORMDIST", "NORM.DIST");
        addAlias("NORMSINV", "NORM.S.INV");
        addAlias("NORMINV", "NORM.INV");
        addAlias("TDIST", "T.DIST.2T");
        addAlias("TINV", "T.INV.2T");
        addAlias("CHIDIST", "CHISQ.DIST.RT");
        addAlias("CHIINV", "CHISQ.INV.RT");
        addAlias("BINOMDIST", "BINOM.DIST");
        addAlias("POISSON", "POISSON.DIST");
        addAlias("EXPONDIST", "EXPON.DIST");

        // 3. Register prefix variants (_xlfn.) for functions
        std::vector<std::string> prefixFunctions = {
            "DAYS", "TEXTJOIN", "MAXIFS", "MINIFS", "CEILING.MATH", "FLOOR.MATH",
            "VAR.P", "STDEV.P", "PERMUTATIONA", "ISOWEEKNUM", "COMBINA", "MODE.SNGL",
            "FORECAST.LINEAR", "NORM.S.DIST", "NORM.DIST", "NORM.S.INV", "NORM.INV",
            "T.DIST", "T.DIST.RT", "T.DIST.2T", "T.INV", "T.INV.2T", "CHISQ.DIST",
            "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "BINOM.DIST", "POISSON.DIST",
            "EXPON.DIST"
        };
        for (const auto& fn : prefixFunctions) {
            addAlias("_xlfn." + fn, fn);
        }

        return map;
    }();
    return builtins;
}

// =============================================================================
// Built-in: Math / Statistical
// =============================================================================

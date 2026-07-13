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
#include "XLEvaluationContext.hpp"
#include "XLWorksheet.hpp"

using namespace OpenXLSX;

// =============================================================================
#include "XLFormulaUtils.hpp"

// =============================================================================
// Lexer
// =============================================================================

// =============================================================================
// XLEvalSession
// =============================================================================

XLFormulaArg XLEvalSession::expandRange(std::string_view rangeRef) const
{
    if (!hasResolver()) return XLFormulaArg();
    return XLFormulaEngine::expandRange(rangeRef, *m_resolver);
}

// =============================================================================
// expandArg / ref-returning helpers
// =============================================================================

bool XLFormulaEngine::isRefReturningFunction(std::string_view name)
{
    return name == "INDIRECT" || name == "OFFSET";
}

XLFormulaArg XLFormulaEngine::evalRefFunction(const XLASTNode& node, XLEvalSession& session) const
{
    XLEvalCallGuard guard(session);
    if (!guard.ok()) {
        XLCellValue e;
        e.setError("#REF!");
        return XLFormulaArg(std::move(e));
    }

    std::vector<XLFormulaArg> argVecs;
    argVecs.reserve(node.children.size());
    for (const auto& child : node.children) argVecs.push_back(expandArg(*child, session));

    if (node.text == "INDIRECT") {
        // INDIRECT(ref_text, [a1=TRUE]) — a1 currently ignored (R1C1 not implemented).
        if (argVecs.empty() || argVecs[0].empty()) {
            XLCellValue e;
            e.setError("#REF!");
            return XLFormulaArg(std::move(e));
        }
        if (isError(argVecs[0][0])) return XLFormulaArg(argVecs[0][0]);

        std::string refText = toString(argVecs[0][0]);
        // Strip optional leading '='
        if (!refText.empty() && refText.front() == '=') refText.erase(refText.begin());
        // Trim whitespace
        while (!refText.empty() && std::isspace(static_cast<unsigned char>(refText.front()))) refText.erase(refText.begin());
        while (!refText.empty() && std::isspace(static_cast<unsigned char>(refText.back()))) refText.pop_back();
        if (refText.empty()) {
            XLCellValue e;
            e.setError("#REF!");
            return XLFormulaArg(std::move(e));
        }

        // Defined name → formula text, then expand.
        if (auto named = session.resolveName(refText)) {
            refText = *named;
            if (!refText.empty() && refText.front() == '=') refText.erase(refText.begin());
        }

        if (!session.hasResolver()) {
            XLCellValue e;
            e.setError("#REF!");
            return XLFormulaArg(std::move(e));
        }
        return session.expandRange(refText);
    }

    if (node.text == "OFFSET") {
        // OFFSET(reference, rows, cols, [height], [width])
        if (argVecs.size() < 3) {
            XLCellValue e;
            e.setError("#VALUE!");
            return XLFormulaArg(std::move(e));
        }

        const auto& base = argVecs[0];
        if (base.type() != XLFormulaArg::Type::LazyRange) {
            // Scalar that might be a ref string — not supported; Excel requires a reference.
            XLCellValue e;
            e.setError("#VALUE!");
            return XLFormulaArg(std::move(e));
        }
        if (argVecs[1].empty() || argVecs[2].empty()) {
            XLCellValue e;
            e.setError("#VALUE!");
            return XLFormulaArg(std::move(e));
        }
        if (!isNumeric(argVecs[1][0]) || !isNumeric(argVecs[2][0])) {
            XLCellValue e;
            e.setError("#VALUE!");
            return XLFormulaArg(std::move(e));
        }

        const int64_t rowOff = static_cast<int64_t>(std::trunc(toDouble(argVecs[1][0])));
        const int64_t colOff = static_cast<int64_t>(std::trunc(toDouble(argVecs[2][0])));

        int64_t height = static_cast<int64_t>(base.rows());
        int64_t width  = static_cast<int64_t>(base.cols());
        if (argVecs.size() >= 4 && !argVecs[3].empty()) {
            if (!isNumeric(argVecs[3][0])) {
                XLCellValue e;
                e.setError("#VALUE!");
                return XLFormulaArg(std::move(e));
            }
            height = static_cast<int64_t>(std::trunc(toDouble(argVecs[3][0])));
        }
        if (argVecs.size() >= 5 && !argVecs[4].empty()) {
            if (!isNumeric(argVecs[4][0])) {
                XLCellValue e;
                e.setError("#VALUE!");
                return XLFormulaArg(std::move(e));
            }
            width = static_cast<int64_t>(std::trunc(toDouble(argVecs[4][0])));
        }
        if (height == 0 || width == 0) {
            XLCellValue e;
            e.setError("#REF!");
            return XLFormulaArg(std::move(e));
        }
        // Negative height/width: Excel mirrors the range; we treat absolute size from TL.
        if (height < 0) height = -height;
        if (width < 0) width = -width;

        const int64_t r1 = static_cast<int64_t>(base.firstRow()) + rowOff;
        const int64_t c1 = static_cast<int64_t>(base.firstCol()) + colOff;
        const int64_t r2 = r1 + height - 1;
        const int64_t c2 = c1 + width - 1;

        if (r1 < 1 || c1 < 1 || r2 < 1 || c2 < 1 || r1 > 1048576 || r2 > 1048576 || c1 > 16384 || c2 > 16384) {
            XLCellValue e;
            e.setError("#REF!");
            return XLFormulaArg(std::move(e));
        }

        if (!session.hasResolver()) {
            XLCellValue e;
            e.setError("#REF!");
            return XLFormulaArg(std::move(e));
        }
        return XLFormulaArg(static_cast<uint32_t>(r1),
                            static_cast<uint32_t>(r2),
                            static_cast<uint16_t>(c1),
                            static_cast<uint16_t>(c2),
                            base.sheetName(),
                            session.resolverPtr());
    }

    XLCellValue e;
    e.setError("#NAME?");
    return XLFormulaArg(std::move(e));
}

XLFormulaArg XLFormulaEngine::expandArg(const XLASTNode& argNode, XLEvalSession& session) const
{
    if (argNode.kind == XLNodeKind::ArrayLit) {
        // Children are row-major constants; number = rows, text = cols as decimal string.
        const size_t rows = static_cast<size_t>(argNode.number > 0 ? argNode.number : 0);
        size_t       cols = 0;
        try {
            cols = static_cast<size_t>(std::stoul(argNode.text));
        }
        catch (...) {
            cols = argNode.children.empty() ? 0 : argNode.children.size();
        }
        std::vector<XLCellValue> vals;
        vals.reserve(argNode.children.size());
        for (const auto& ch : argNode.children) vals.push_back(evalNode(*ch, session));
        if (rows == 0 || cols == 0) {
            if (vals.empty()) return XLFormulaArg();
            return XLFormulaArg(std::move(vals), 1, vals.size());
        }
        return XLFormulaArg(std::move(vals), rows, cols);
    }

    if (argNode.kind == XLNodeKind::Range) {
        if (!session.hasResolver()) return XLFormulaArg();
        return session.expandRange(argNode.text);
    }

    // Single cell reference: create a 1x1 LazyRange so position metadata (row/col) is preserved.
    // Functions like ROW() and COLUMN() need the address, not just the resolved value.
    if (argNode.kind == XLNodeKind::CellRef) {
        if (!session.hasResolver()) return XLFormulaArg();
        // Named ranges masquerading as Ident/CellRef: try A1 expand first, then name resolver.
        auto expanded = session.expandRange(argNode.text);
        if (expanded.type() == XLFormulaArg::Type::LazyRange) return expanded;
        // expandRange fell back to resolver() scalar (empty for unknown names) — try defined names.
        if (auto named = session.resolveName(argNode.text)) {
            std::string ref = *named;
            if (!ref.empty() && ref.front() == '=') ref.erase(ref.begin());
            return session.expandRange(ref);
        }
        return expanded;
    }

    // All function calls preserve full shape when used as arguments (INDIRECT range,
    // FILTER result into SUM/INDEX, etc.).
    if (argNode.kind == XLNodeKind::FuncCall) {
        if (isRefReturningFunction(argNode.text)) return evalRefFunction(argNode, session);
        return evalFunctionAsArg(argNode, session);
    }

    // Phase B: arithmetic on ranges used as function args must keep full element vectors
    // e.g. SUM(A1:A3*2) / SUM(A1:A3+B1:B3)
    if (argNode.kind == XLNodeKind::BinOp && argNode.op != XLTokenKind::Amp) {
        Expects(argNode.children.size() == 2);
        auto left  = expandArg(*argNode.children[0], session);
        auto right = expandArg(*argNode.children[1], session);
        if (left.empty() && right.empty()) return XLFormulaArg();
        // Materialize LazyRanges so at(r,c) is efficient and shape is accurate.
        left  = left.materialize();
        right = right.materialize();

        // Excel arithmetic: blank cells coerce to 0; non-numeric text → #VALUE!.
        auto coerceArith = [](const XLCellValue& v) -> XLCellValue {
            if (isError(v)) return v;
            if (isEmpty(v)) return XLCellValue(0.0);
            if (isNumeric(v)) return v;
            if (v.type() == XLValueType::String) {
                try {
                    std::size_t idx = 0;
                    std::string s   = v.get<std::string>();
                    double      d   = std::stod(s, &idx);
                    if (idx == s.size() && !s.empty()) return XLCellValue(d);
                }
                catch (...) {
                }
            }
            return errValue();
        };

        auto applyOne = [&](const XLCellValue& lvIn, const XLCellValue& rvIn) -> XLCellValue {
            if (isError(lvIn)) return lvIn;
            if (isError(rvIn)) return rvIn;
            if (argNode.op == XLTokenKind::Plus || argNode.op == XLTokenKind::Minus || argNode.op == XLTokenKind::Star ||
                argNode.op == XLTokenKind::Slash || argNode.op == XLTokenKind::Caret)
            {
                auto lv = coerceArith(lvIn);
                auto rv = coerceArith(rvIn);
                if (isError(lv)) return lv;
                if (isError(rv)) return rv;
                double l = toDouble(lv), r = toDouble(rv);
                switch (argNode.op) {
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
            // Comparisons: blank treated as empty string / 0 via existing paths.
            auto lv = lvIn;
            auto rv = rvIn;
            if (isEmpty(lv) && isNumeric(rv)) lv = XLCellValue(0.0);
            if (isEmpty(rv) && isNumeric(lv)) rv = XLCellValue(0.0);
            bool result = false;
            if (isNumeric(lv) && isNumeric(rv)) {
                double l = toDouble(lv), r = toDouble(rv);
                switch (argNode.op) {
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
                std::string ls = toString(lv), rs = toString(rv);
                std::transform(ls.begin(), ls.end(), ls.begin(), ::tolower);
                std::transform(rs.begin(), rs.end(), rs.begin(), ::tolower);
                switch (argNode.op) {
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
        };

        // Excel dynamic-array style broadcast:
        // - equal dims: element-wise
        // - 1 in a dimension: expand that side
        // - row (1×n) × column (m×1) → m×n outer-style broadcast
        // - otherwise #N/A
        const size_t r1 = std::max<size_t>(left.rows(), 1);
        const size_t c1 = std::max<size_t>(left.cols(), 1);
        const size_t r2 = std::max<size_t>(right.rows(), 1);
        const size_t c2 = std::max<size_t>(right.cols(), 1);
        const bool   rowsOk = (r1 == r2) || r1 == 1 || r2 == 1;
        const bool   colsOk = (c1 == c2) || c1 == 1 || c2 == 1;
        if (!rowsOk || !colsOk) {
            XLCellValue e;
            e.setError("#N/A");
            return XLFormulaArg(std::move(e));
        }
        const size_t outRows = std::max(r1, r2);
        const size_t outCols = std::max(c1, c2);

        auto atClamped = [](const XLFormulaArg& a, size_t r, size_t c, size_t ar, size_t ac) -> XLCellValue {
            const size_t rr = (ar == 1) ? 0 : r;
            const size_t cc = (ac == 1) ? 0 : c;
            return a.at(rr, cc);
        };

        std::vector<XLCellValue> out;
        out.reserve(outRows * outCols);
        for (size_t r = 0; r < outRows; ++r) {
            for (size_t c = 0; c < outCols; ++c) {
                out.push_back(applyOne(atClamped(left, r, c, r1, c1), atClamped(right, r, c, r2, c2)));
            }
        }
        return XLFormulaArg(std::move(out), outRows, outCols);
    }

    if (argNode.kind == XLNodeKind::UnaryOp) {
        Expects(argNode.children.size() == 1);
        auto arg = expandArg(*argNode.children[0], session).materialize();
        if (arg.size() == 0) return XLFormulaArg();
        auto applyU = [](const XLCellValue& val, XLTokenKind op) -> XLCellValue {
            if (isError(val)) return val;
            // Blank → 0 for unary minus / percent (Excel).
            double d = 0.0;
            if (isEmpty(val))
                d = 0.0;
            else if (isNumeric(val))
                d = toDouble(val);
            else
                return errValue();
            if (op == XLTokenKind::Minus) return XLCellValue(-d);
            if (op == XLTokenKind::Percent) return XLCellValue(d / 100.0);
            return val;
        };
        if (arg.size() == 1) return XLFormulaArg(applyU(arg[0], argNode.op));
        std::vector<XLCellValue> out;
        out.reserve(arg.size());
        for (size_t i = 0; i < arg.size(); ++i) out.push_back(applyU(arg[i], argNode.op));
        return XLFormulaArg(std::move(out), arg.rows(), arg.cols());
    }

    // Evaluate normally and wrap in a single-element scalar
    return XLFormulaArg(evalNode(argNode, session));
}

// =============================================================================
// Evaluator – evalNode
// =============================================================================

XLCellValue XLFormulaEngine::evalNode(const XLASTNode& node, XLEvalSession& session) const
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
            // Prefer defined names when the token is not a plain A1 address that the
            // resolver can answer; still try cell lookup first for normal refs.
            auto val = session.cellValue(node.text);
            if (val.type() != XLValueType::Empty) return val;
            if (auto named = session.resolveName(node.text)) {
                std::string ref = *named;
                if (!ref.empty() && ref.front() == '=') ref.erase(ref.begin());
                auto arg = session.expandRange(ref);
                return arg.empty() ? XLCellValue{} : arg[0];
            }
            return val;
        }

        case XLNodeKind::Range: {
            // Range used as scalar = first cell value
            auto vals = session.hasResolver() ? session.expandRange(node.text) : XLFormulaArg();
            return vals.empty() ? XLCellValue{} : vals[0];
        }

        case XLNodeKind::ArrayLit: {
            // Implicit intersection: top-left element of the array constant
            return expandArg(node, session).asScalar();
        }

        case XLNodeKind::UnaryOp: {
            // Delegate to expandArg for consistent empty-coercion / shape rules; scalar = top-left.
            return expandArg(node, session).asScalar();
        }

        case XLNodeKind::BinOp: {
            Expects(node.children.size() == 2);
            // String concat – evaluate early, no numeric coercion (scalar path)
            if (node.op == XLTokenKind::Amp) {
                auto lv = evalNode(*node.children[0], session);
                auto rv = evalNode(*node.children[1], session);
                if (isError(lv)) return lv;
                if (isError(rv)) return rv;
                return XLCellValue(toString(lv) + toString(rv));
            }
            // Arithmetic / comparison: full 2-D broadcast in expandArg, then implicit intersection.
            return expandArg(node, session).asScalar();
        }

        case XLNodeKind::FuncCall: {
            auto arg = isRefReturningFunction(node.text) ? evalRefFunction(node, session) : evalFunctionAsArg(node, session);
            return arg.asScalar();
        }

        default:
            return errValue();
    }
}

XLFormulaArg XLFormulaEngine::evalFunctionAsArg(const XLASTNode& node, XLEvalSession& session) const
{
    const auto& funcs = getBuiltins();
    auto        it    = funcs.find(node.text);
    // Safety net: strip _XLFN. / _XLWS. if present (parser normally strips already)
    if (it == funcs.end() && node.text.size() > 6) {
        std::string_view n = node.text;
        if (n.compare(0, 6, "_XLFN.") == 0 || n.compare(0, 6, "_XLWS.") == 0) {
            it = funcs.find(std::string(n.substr(6)));
        }
    }
    if (it == funcs.end()) {
        XLCellValue e;
        e.setError("#NAME?");
        return XLFormulaArg(std::move(e));
    }

    std::vector<XLFormulaArg> argVecs;
    argVecs.reserve(node.children.size());
    for (const auto& child : node.children) argVecs.push_back(expandArg(*child, session));

    try {
        return it->second(argVecs, session);
    }
    catch (const std::exception& ex) {
        XLCellValue e;
        e.setError(std::string("#ERROR: ") + ex.what());
        return XLFormulaArg(std::move(e));
    }
}

XLFormulaArg XLFormulaEngine::evalAsArg(const XLASTNode& node, XLEvalSession& session) const
{
    // Preserve shape for roots that are ranges, arrays, or shape-producing ops/functions.
    if (node.kind == XLNodeKind::Range || node.kind == XLNodeKind::CellRef || node.kind == XLNodeKind::FuncCall ||
        node.kind == XLNodeKind::ArrayLit || node.kind == XLNodeKind::BinOp || node.kind == XLNodeKind::UnaryOp)
    {
        return expandArg(node, session);
    }
    return XLFormulaArg(evalNode(node, session));
}

// =============================================================================
// Evaluator – public evaluate()
// =============================================================================

XLCellValue XLFormulaEngine::evaluate(std::string_view formula, const XLCellResolver& resolver, XLFormulaDiagnosticReporter* reporter) const
{
    XLEvalSession session(resolver);
    return evaluate(formula, session, reporter);
}

XLCellValue XLFormulaEngine::evaluate(std::string_view formula, XLEvalSession& session, XLFormulaDiagnosticReporter* reporter) const
{
    return evaluateArray(formula, session, reporter).asScalar();
}

XLFormulaArg XLFormulaEngine::evaluateArray(std::string_view formula, const XLCellResolver& resolver, XLFormulaDiagnosticReporter* reporter) const
{
    XLEvalSession session(resolver);
    return evaluateArray(formula, session, reporter);
}

std::string XLFormulaEngine::cacheKey(std::string_view formula)
{
    // Strip leading '=' and surrounding whitespace for stable cache keys
    std::size_t b = 0, e = formula.size();
    while (b < e && std::isspace(static_cast<unsigned char>(formula[b]))) ++b;
    while (e > b && std::isspace(static_cast<unsigned char>(formula[e - 1]))) --e;
    if (b < e && formula[b] == '=') {
        ++b;
        while (b < e && std::isspace(static_cast<unsigned char>(formula[b]))) ++b;
    }
    return std::string(formula.substr(b, e - b));
}

void XLFormulaEngine::setAstCacheCapacity(std::size_t capacity) noexcept
{
    m_astCacheCapacity = capacity == 0 ? 1 : capacity;
    std::lock_guard<std::mutex> lock(m_astCacheMutex);
    while (m_astCache.size() > m_astCacheCapacity) m_astCache.erase(m_astCache.begin());
}

std::size_t XLFormulaEngine::astCacheSize() const
{
    std::lock_guard<std::mutex> lock(m_astCacheMutex);
    return m_astCache.size();
}

void XLFormulaEngine::clearAstCache()
{
    std::lock_guard<std::mutex> lock(m_astCacheMutex);
    m_astCache.clear();
}

std::shared_ptr<XLASTNode> XLFormulaEngine::getOrParseAst(std::string_view formula,
                                                          XLFormulaDiagnosticReporter* reporter) const
{
    if (reporter) {
        *reporter = XLFormulaDiagnosticReporter(std::string(formula));
        auto tokens = XLFormulaLexer::tokenize(formula);
        auto ast    = XLFormulaParser::parse(gsl::span<const XLToken>(tokens), reporter);
        return std::shared_ptr<XLASTNode>(ast.release());
    }

    if (!m_astCacheEnabled) {
        auto tokens = XLFormulaLexer::tokenize(formula);
        auto ast    = XLFormulaParser::parse(gsl::span<const XLToken>(tokens), nullptr);
        return std::shared_ptr<XLASTNode>(ast.release());
    }

    const std::string key = cacheKey(formula);
    {
        std::lock_guard<std::mutex> lock(m_astCacheMutex);
        auto it = m_astCache.find(key);
        if (it != m_astCache.end()) return it->second;
    }

    auto tokens = XLFormulaLexer::tokenize(formula);
    auto astUp  = XLFormulaParser::parse(gsl::span<const XLToken>(tokens), nullptr);
    std::shared_ptr<XLASTNode> ast(astUp.release());

    {
        std::lock_guard<std::mutex> lock(m_astCacheMutex);
        if (m_astCache.size() >= m_astCacheCapacity && m_astCache.find(key) == m_astCache.end()) {
            // Drop one arbitrary entry (unordered_map begin) when full
            m_astCache.erase(m_astCache.begin());
        }
        auto [it, inserted] = m_astCache.emplace(key, ast);
        return it->second;
    }
}

XLFormulaArg XLFormulaEngine::evaluateArray(std::string_view formula, XLEvalSession& session, XLFormulaDiagnosticReporter* reporter) const
{
    if (formula.empty()) return XLFormulaArg();
    try {
        auto ast = getOrParseAst(formula, reporter);
        if (!ast) return XLFormulaArg();
        return evalAsArg(*ast, session);
    }
    catch (const XLException&) {
        throw;
    }
    catch (const std::exception& ex) {
        XLCellValue e;
        e.setError(std::string("#ERROR: ") + ex.what());
        return XLFormulaArg(std::move(e));
    }
}

namespace
{
    void spillShape(const XLFormulaArg& result, size_t& rows, size_t& cols)
    {
        auto mat = result.materialize();
        if (mat.type() == XLFormulaArg::Type::Scalar || mat.size() <= 1) {
            rows = 1;
            cols = 1;
            return;
        }
        rows = std::max<size_t>(mat.rows(), 1);
        cols = std::max<size_t>(mat.cols(), 1);
        if (rows * cols == 0) {
            rows = 1;
            cols = 1;
        }
    }
}    // namespace

size_t OpenXLSX::spillArray(const XLFormulaArg& result, uint32_t topRow, uint16_t leftCol, const XLSpillWriter& write)
{
    if (!write || topRow == 0 || leftCol == 0) return 0;
    auto mat = result.materialize();
    if (mat.empty() && mat.type() != XLFormulaArg::Type::Array) {
        // empty scalar still "spills" nothing
        if (result.type() == XLFormulaArg::Type::Scalar) {
            write(topRow, leftCol, result.asScalar());
            return 1;
        }
        return 0;
    }
    if (mat.type() == XLFormulaArg::Type::Scalar || (mat.rows() == 0 && mat.cols() == 0 && mat.size() == 1)) {
        write(topRow, leftCol, mat.asScalar());
        return 1;
    }
    const size_t r = mat.rows();
    const size_t c = mat.cols();
    size_t       n = 0;
    for (size_t i = 0; i < r; ++i) {
        for (size_t j = 0; j < c; ++j) {
            write(topRow + static_cast<uint32_t>(i), static_cast<uint16_t>(leftCol + j), mat.at(i, j));
            ++n;
        }
    }
    return n;
}

bool OpenXLSX::spillRangeIsClear(const XLFormulaArg&     result,
                                 uint32_t                topRow,
                                 uint16_t                leftCol,
                                 const XLSpillOccupancy& isOccupied,
                                 uint32_t                anchorRow,
                                 uint16_t                anchorCol)
{
    if (!isOccupied || topRow == 0 || leftCol == 0) return true;
    size_t rows = 0, cols = 0;
    spillShape(result, rows, cols);
    for (size_t i = 0; i < rows; ++i) {
        for (size_t j = 0; j < cols; ++j) {
            const uint32_t r = topRow + static_cast<uint32_t>(i);
            const uint16_t c = static_cast<uint16_t>(leftCol + j);
            if (anchorRow != 0 && anchorCol != 0 && r == anchorRow && c == anchorCol) continue;
            if (isOccupied(r, c)) return false;
        }
    }
    return true;
}

XLSpillResult OpenXLSX::spillArrayChecked(const XLFormulaArg&     result,
                                           uint32_t                topRow,
                                           uint16_t                leftCol,
                                           const XLSpillWriter&    write,
                                           const XLSpillOccupancy& isOccupied,
                                           uint32_t                anchorRow,
                                           uint16_t                anchorCol)
{
    XLSpillResult out;
    spillShape(result, out.rows, out.cols);
    if (topRow == 0 || leftCol == 0) {
        out.error.setError("#SPILL!");
        return out;
    }
    if (!spillRangeIsClear(result, topRow, leftCol, isOccupied, anchorRow, anchorCol)) {
        out.error.setError("#SPILL!");
        return out;
    }
    out.cellsWritten = spillArray(result, topRow, leftCol, write);
    out.ok           = true;
    return out;
}

// =============================================================================
// makeResolver
// =============================================================================

XLCellResolver XLFormulaEngine::makeResolver(const XLEvaluationContext& context)
{
    return [&context](std::string_view ref) -> XLCellValue { return context.cellValue(ref); };
}

XLCellResolver XLFormulaEngine::makeResolver(const XLWorksheet& wks)
{
    // Capture worksheet pointer; each lookup builds a thin context adapter so cross-sheet
    // resolution stays in XLWorksheetEvaluationContext (not duplicated here).
    return [&wks](std::string_view ref) -> XLCellValue {
        return XLWorksheetEvaluationContext(wks).cellValue(ref);
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

        // 1. Register all functions from the registry (return full XLFormulaArg shape)
        for (const auto& pair : registry.getFunctions()) {
            const auto& name = pair.first;
            const auto& func = pair.second;
            map[name] = [func](const std::vector<XLFormulaArg>& args, XLEvalSession& session) -> XLFormulaArg {
                return func->execute(args, session);
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

        // 3. Register _xlfn. / _xlws. prefix aliases for every registered function (Excel future-fn encoding)
        // Snapshot names first — map grows as we insert aliases.
        std::vector<std::string> registeredNames;
        registeredNames.reserve(map.size());
        for (const auto& pair : map) registeredNames.push_back(pair.first);
        for (const auto& fn : registeredNames) {
            // Skip names that already carry a future-function prefix
            if (fn.size() >= 6 && (fn.compare(0, 6, "_XLFN.") == 0 || fn.compare(0, 6, "_XLWS.") == 0)) continue;
            addAlias("_XLFN." + fn, fn);
            addAlias("_XLWS." + fn, fn);
            // Lower-case prefix variants used in some serialized workbooks before uppercasing
            addAlias("_xlfn." + fn, fn);
            addAlias("_xlws." + fn, fn);
        }

        return map;
    }();
    return builtins;
}

// =============================================================================
// Built-in: Math / Statistical
// =============================================================================

/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

// ===== External Includes ===== //
#include <algorithm>
#include <cassert>
#include <cctype>
#include <cmath>
#include <ctime>
#include <fmt/format.h>
#include <functional>
#include <list>
#include <numeric>
#include <string>
#include <unordered_map>
#include <vector>
#include <ankerl/unordered_dense.h>

// ===== OpenXLSX Includes ===== //
#include "XLCellReference.hpp"
#include "XLDateTime.hpp"
#include "XLException.hpp"
#include "XLFormulaEngine.hpp"
#include "XLFormulaRegistry.hpp"
#include "XLEvaluationContext.hpp"
#include "IXLCellProvider.hpp"
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

XLEvalSession::PackedCellKey XLEvalSession::parsePackedKey(std::string_view ref) const
{
    if (ref.empty()) return PackedCellKey{0};

    // 1. Split sheetPart and cellPart if '!' is present outside single quotes
    std::string_view sheetPart;
    std::string_view cellPart = ref;

    bool inQuote = false;
    size_t bangPos = std::string_view::npos;
    for (size_t i = 0; i < ref.size(); ++i) {
        if (ref[i] == '\'') {
            inQuote = !inQuote;
        } else if (ref[i] == '!' && !inQuote) {
            bangPos = i;
            break;
        }
    }

    if (bangPos != std::string_view::npos) {
        sheetPart = ref.substr(0, bangPos);
        cellPart = ref.substr(bangPos + 1);
        if (!sheetPart.empty() && sheetPart.front() == '\'' && sheetPart.back() == '\'') {
            sheetPart = sheetPart.substr(1, sheetPart.size() - 2);
        }
    } else {
        sheetPart = m_currentSheet;
    }

    // 2. Compute 16-bit Sheet ID
    // If sheetPart is empty or matches m_currentSheet case-insensitively, treat as primary (0)
    uint16_t sheetId = 0;
    if (!sheetPart.empty()) {
        bool isCurrent = false;
        if (!m_currentSheet.empty() && sheetPart.size() == m_currentSheet.size()) {
            isCurrent = std::equal(sheetPart.begin(), sheetPart.end(), m_currentSheet.begin(),
                [](char a, char b) { return std::toupper(static_cast<unsigned char>(a)) == std::toupper(static_cast<unsigned char>(b)); });
        }
        if (!isCurrent) {
            uint32_t h = 2166136261u;
            for (char c : sheetPart) {
                h = (h ^ static_cast<unsigned char>(std::toupper(static_cast<unsigned char>(c)))) * 16777619u;
            }
            sheetId = static_cast<uint16_t>((h ^ (h >> 16)) & 0xFFFF);
            if (sheetId == 0) sheetId = 1;
        }
    }

    // 3. Parse cellPart (ignore $, parse alpha column and numeric row)
    uint32_t col = 0;
    size_t i = 0;
    while (i < cellPart.size() && std::isspace(static_cast<unsigned char>(cellPart[i]))) ++i;

    while (i < cellPart.size()) {
        char c = cellPart[i];
        if (c == '$') { ++i; continue; }
        if (std::isalpha(static_cast<unsigned char>(c))) {
            col = col * 26 + (std::toupper(static_cast<unsigned char>(c)) - 'A' + 1);
            ++i;
        } else {
            break;
        }
    }

    uint32_t row = 0;
    while (i < cellPart.size()) {
        char c = cellPart[i];
        if (c == '$') { ++i; continue; }
        if (std::isdigit(static_cast<unsigned char>(c))) {
            row = row * 10 + (c - '0');
            ++i;
        } else {
            break;
        }
    }

    while (i < cellPart.size() && std::isspace(static_cast<unsigned char>(cellPart[i]))) ++i;

    // Fallback for defined names or non-coordinate references: 64-bit FNV-1a hash
    if (col == 0 || row == 0 || i != cellPart.size() || row > 16777215 || col > 16777215) {
        uint64_t nameHash = 14695981039346656037ull;
        for (char c : ref) {
            nameHash = (nameHash ^ static_cast<unsigned char>(std::toupper(static_cast<unsigned char>(c)))) * 1099511628211ull;
        }
        return PackedCellKey(nameHash | (1ull << 63));
    }

    uint64_t packed = ((uint64_t)sheetId << 48) | ((uint64_t)row << 24) | (uint64_t)col;
    return PackedCellKey(packed);
}

bool XLEvalSession::pushEvaluatingCell(std::string_view ref)
{
    PackedCellKey key = parsePackedKey(ref);
    if (!key) return true;

    for (size_t i = 0; i < m_evaluatingCellsInlineCount; ++i) {
        if (m_evaluatingCellsInline[i] == key) return false;
    }
    for (const auto& existing : m_evaluatingCellsHeap) {
        if (existing == key) return false;
    }

    if (m_evaluatingCellsInlineCount < kInlineCellStackCapacity) {
        m_evaluatingCellsInline[m_evaluatingCellsInlineCount++] = key;
    } else {
        m_evaluatingCellsHeap.push_back(key);
    }
    return true;
}

void XLEvalSession::popEvaluatingCell() noexcept
{
    if (!m_evaluatingCellsHeap.empty()) {
        m_evaluatingCellsHeap.pop_back();
    } else if (m_evaluatingCellsInlineCount > 0) {
        --m_evaluatingCellsInlineCount;
    }
}

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
    XLEvalNodeGuard nodeGuard(session);
    if (!nodeGuard.ok()) {
        XLCellValue e;
        e.setError("#DEPTH!");
        return XLFormulaArg(std::move(e));
    }

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
        XLEvalCellGuard cellGuard(session, argNode.text);
        if (!cellGuard.ok()) {
            XLCellValue e;
            e.setError("#CIRC!");
            return XLFormulaArg(std::move(e));
        }
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
                double result = 0.0;
                switch (argNode.op) {
                    case XLTokenKind::Plus:
                        result = l + r;
                        break;
                    case XLTokenKind::Minus:
                        result = l - r;
                        break;
                    case XLTokenKind::Star:
                        result = l * r;
                        break;
                    case XLTokenKind::Slash:
                        if (r == 0.0) return errDiv0();
                        result = l / r;
                        break;
                    case XLTokenKind::Caret:
                        if (l < 0.0 && std::floor(r) != r) return errNum();
                        if (l == 0.0 && r < 0.0) return errDiv0();
                        result = std::pow(l, r);
                        break;
                    default:
                        break;
                }
                if (std::isnan(result)) return errNum();
                if (std::isinf(result)) {
                    if (argNode.op == XLTokenKind::Slash && r == 0.0) return errDiv0();
                    return errNum();
                }
                return XLCellValue(result);
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
    XLEvalNodeGuard nodeGuard(session);
    if (!nodeGuard.ok()) {
        XLCellValue e;
        e.setError("#DEPTH!");
        return e;
    }

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
            XLEvalCellGuard cellGuard(session, node.text);
            if (!cellGuard.ok()) {
                XLCellValue e;
                e.setError("#CIRC!");
                return e;
            }
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
    const auto* func = findBuiltin(node.text);
    if (!func) {
        XLCellValue e;
        e.setError("#NAME?");
        return XLFormulaArg(std::move(e));
    }

    std::vector<XLFormulaArg> argVecs;
    argVecs.reserve(node.children.size());
    for (const auto& child : node.children) argVecs.push_back(expandArg(*child, session));

    try {
        return (*func)(argVecs, session);
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
    XLEvalCallGuard guard(session);
    if (!guard.ok()) {
        XLCellValue e;
        e.setError("#CIRC!");
        return e;
    }
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

namespace OpenXLSX
{
    class ShardedAstCache
    {
    public:
        static constexpr size_t kShardCount = 8;

        struct Shard {
            mutable std::mutex mutex;
            std::size_t capacity{64};
            std::list<std::string> lruList;
            ankerl::unordered_dense::map<
                std::string,
                std::pair<std::shared_ptr<XLASTNode>, std::list<std::string>::iterator>>
                map;
        };

        explicit ShardedAstCache(std::size_t totalCapacity = 512)
        {
            setCapacity(totalCapacity);
        }

        void setCapacity(std::size_t totalCapacity) noexcept
        {
            const std::size_t perShard = (std::max<std::size_t>(totalCapacity, kShardCount) + kShardCount - 1) / kShardCount;
            for (auto& s : m_shards) {
                std::lock_guard<std::mutex> lock(s.mutex);
                s.capacity = perShard;
                while (s.map.size() > s.capacity && !s.lruList.empty()) {
                    const std::string oldest = s.lruList.back();
                    s.map.erase(oldest);
                    s.lruList.pop_back();
                }
            }
        }

        [[nodiscard]] std::size_t size() const noexcept
        {
            std::size_t total = 0;
            for (const auto& s : m_shards) {
                std::lock_guard<std::mutex> lock(s.mutex);
                total += s.map.size();
            }
            return total;
        }

        void clear() noexcept
        {
            for (auto& s : m_shards) {
                std::lock_guard<std::mutex> lock(s.mutex);
                s.map.clear();
                s.lruList.clear();
            }
        }

        [[nodiscard]] std::shared_ptr<XLASTNode> get(std::string_view key) const
        {
            const size_t shardIdx = hashKey(key) % kShardCount;
            auto& s = m_shards[shardIdx];
            std::lock_guard<std::mutex> lock(s.mutex);
            auto it = s.map.find(std::string(key));
            if (it == s.map.end()) return nullptr;
            s.lruList.splice(s.lruList.begin(), s.lruList, it->second.second);
            return it->second.first;
        }

        void put(std::string key, std::shared_ptr<XLASTNode> ast)
        {
            const size_t shardIdx = hashKey(key) % kShardCount;
            auto& s = m_shards[shardIdx];
            std::lock_guard<std::mutex> lock(s.mutex);
            auto it = s.map.find(key);
            if (it != s.map.end()) {
                it->second.first = std::move(ast);
                s.lruList.splice(s.lruList.begin(), s.lruList, it->second.second);
                return;
            }
            if (s.map.size() >= s.capacity && !s.lruList.empty()) {
                const std::string oldest = s.lruList.back();
                s.map.erase(oldest);
                s.lruList.pop_back();
            }
            s.lruList.push_front(key);
            s.map.emplace(std::move(key), std::make_pair(std::move(ast), s.lruList.begin()));
        }

    private:
        static size_t hashKey(std::string_view sv) noexcept
        {
            size_t h = 2166136261u;
            for (char c : sv) h = (h ^ static_cast<unsigned char>(c)) * 16777619u;
            return h;
        }

        mutable std::array<Shard, kShardCount> m_shards;
    };
}

void XLFormulaEngine::setAstCacheCapacity(std::size_t capacity) noexcept
{
    m_astCacheCapacity = capacity == 0 ? 1 : capacity;
    if (m_astCache) m_astCache->setCapacity(m_astCacheCapacity);
}

std::size_t XLFormulaEngine::astCacheSize() const
{
    return m_astCache ? m_astCache->size() : 0;
}

void XLFormulaEngine::clearAstCache()
{
    if (m_astCache) m_astCache->clear();
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

    if (!m_astCacheEnabled || !m_astCache) {
        auto tokens = XLFormulaLexer::tokenize(formula);
        auto ast    = XLFormulaParser::parse(gsl::span<const XLToken>(tokens), nullptr);
        return std::shared_ptr<XLASTNode>(ast.release());
    }

    const std::string key = cacheKey(formula);
    if (auto cached = m_astCache->get(key)) {
        return cached;
    }

    auto tokens = XLFormulaLexer::tokenize(formula);
    auto astUp  = XLFormulaParser::parse(gsl::span<const XLToken>(tokens), nullptr);
    std::shared_ptr<XLASTNode> ast(astUp.release());

    m_astCache->put(key, ast);
    return ast;
}

XLFormulaArg XLFormulaEngine::evaluateArray(std::string_view formula, XLEvalSession& session, XLFormulaDiagnosticReporter* reporter) const
{
    XLEvalCallGuard guard(session);
    if (!guard.ok()) {
        XLCellValue e;
        e.setError("#CIRC!");
        return XLFormulaArg(std::move(e));
    }

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

XLCellResolver XLFormulaEngine::makeResolver(const IXLCellProvider& provider)
{
    return [&provider](std::string_view ref) -> XLCellValue {
        std::string_view localRef = ref;
        auto bangPos = ref.find('!');
        if (bangPos != std::string_view::npos) {
            std::string_view sheetPart = ref.substr(0, bangPos);
            if (!sheetPart.empty() && (sheetPart.front() == '\'' || sheetPart.front() == '"')) {
                sheetPart.remove_prefix(1);
                if (!sheetPart.empty() && (sheetPart.back() == '\'' || sheetPart.back() == '"'))
                    sheetPart.remove_suffix(1);
            }
            if (sheetPart != provider.sheetName()) {
                return XLCellValue();
            }
            localRef = ref.substr(bangPos + 1);
        }
        try {
            XLCellReference cellRef(localRef);
            if (cellRef.row() == 0 || cellRef.column() == 0) return XLCellValue();
            return provider.getCellValue(cellRef.row(), cellRef.column());
        }
        catch (...) {
            return XLCellValue();
        }
    };
}

// =============================================================================
// Built-in function registrations
// =============================================================================

XLFormulaEngine::XLFormulaEngine()
    : m_astCache(std::make_unique<ShardedAstCache>(m_astCacheCapacity))
{
}

XLFormulaEngine::~XLFormulaEngine() = default;

namespace
{
    struct CaseInsensitiveStringViewHash
    {
        using is_transparent = void;
        size_t operator()(std::string_view sv) const noexcept
        {
            size_t h = 2166136261u;
            for (char c : sv) h = (h ^ static_cast<unsigned char>(std::toupper(static_cast<unsigned char>(c)))) * 16777619u;
            return h;
        }
    };

    struct CaseInsensitiveStringViewEqual
    {
        using is_transparent = void;
        bool operator()(std::string_view a, std::string_view b) const noexcept
        {
            if (a.size() != b.size()) return false;
            for (size_t i = 0; i < a.size(); ++i) {
                if (std::toupper(static_cast<unsigned char>(a[i])) != std::toupper(static_cast<unsigned char>(b[i]))) {
                    return false;
                }
            }
            return true;
        }
    };
}    // namespace

const XLFormulaEngine::FuncImpl* XLFormulaEngine::findBuiltin(std::string_view name)
{
    using BuiltinMap = ankerl::unordered_dense::map<
        std::string,
        XLFormulaEngine::FuncImpl,
        CaseInsensitiveStringViewHash,
        CaseInsensitiveStringViewEqual>;

    static const BuiltinMap builtins = []() {
        BuiltinMap map;
        const auto& registry = XLFormulaRegistry::getInstance();

        // 1. Register all functions from the registry (return full XLFormulaArg shape)
        for (const auto& pair : registry.getFunctions()) {
            const auto& fnName = pair.first;
            const auto& func   = pair.second;
            map[fnName]        = [func](const std::vector<XLFormulaArg>& args, XLEvalSession& session) -> XLFormulaArg {
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

        return map;
    }();

    std::string_view target = name;
    if (target.size() > 6) {
        std::string_view pfx = target.substr(0, 6);
        if (CaseInsensitiveStringViewEqual{}(pfx, "_XLFN.") || CaseInsensitiveStringViewEqual{}(pfx, "_XLWS.")) {
            target.remove_prefix(6);
        }
    }

    auto it = builtins.find(target);
    if (it != builtins.end()) return &it->second;
    return nullptr;
}

// =============================================================================
// Built-in: Math / Statistical
// =============================================================================

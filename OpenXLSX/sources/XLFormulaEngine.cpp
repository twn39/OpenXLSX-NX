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
#include "XLWorksheet.hpp"

using namespace OpenXLSX;

// =============================================================================
// Internal helpers (anonymous namespace)
// =============================================================================
namespace
{
    // Convert a value to double; return NaN for non-numeric/empty.
    double toDouble(const XLCellValue& v)
    {
        switch (v.type()) {
            case XLValueType::Integer:
                return static_cast<double>(v.get<int64_t>());
            case XLValueType::Float:
                return v.get<double>();
            case XLValueType::Boolean:
                return v.get<bool>() ? 1.0 : 0.0;
            default:
                return std::numeric_limits<double>::quiet_NaN();
        }
    }

    bool isNumeric(const XLCellValue& v)
    { return v.type() == XLValueType::Integer || v.type() == XLValueType::Float || v.type() == XLValueType::Boolean; }

    bool isEmpty(const XLCellValue& v) { return v.type() == XLValueType::Empty; }

    bool isError(const XLCellValue& v) { return v.type() == XLValueType::Error; }

    std::string toString(const XLCellValue& v)
    {
        switch (v.type()) {
            case XLValueType::String:
                return v.get<std::string>();
            case XLValueType::Integer:
                return fmt::format("{}", v.get<int64_t>());
            case XLValueType::Float:
                return fmt::format("{}", v.get<double>());
            case XLValueType::Boolean:
                return v.get<bool>() ? "TRUE" : "FALSE";
            default:
                return "";
        }
    }

    XLCellValue errValue()
    {
        XLCellValue r;
        r.setError("#VALUE!");
        return r;
    }
    XLCellValue errDiv0()
    {
        XLCellValue r;
        r.setError("#DIV/0!");
        return r;
    }
    XLCellValue errNA()
    {
        XLCellValue r;
        r.setError("#N/A");
        return r;
    }
    XLCellValue errNum()
    {
        XLCellValue r;
        r.setError("#NUM!");
        return r;
    }
    XLCellValue errRef()
    {
        XLCellValue r;
        r.setError("#REF!");
        return r;
    }
    XLCellValue errName()
    {
        XLCellValue r;
        r.setError("#NAME?");
        return r;
    }

    // Trim leading/trailing whitespace
    std::string strTrim(std::string s)
    {
        const char* ws = " \t\r\n";
        s.erase(0, s.find_first_not_of(ws));
        s.erase(s.find_last_not_of(ws) + 1);
        return s;
    }

    // Collect numeric values from a list of arguments directly
    std::vector<double> numerics(const std::vector<XLFormulaArg>& args)
    {
        std::vector<double> out;
        for (const auto& arg : args)
            for (const auto& v : arg)
                if (isNumeric(v)) out.push_back(toDouble(v));
        return out;
    }

    // Collect numeric values from a single argument
    std::vector<double> numerics(const XLFormulaArg& arg)
    {
        std::vector<double> out;
        for (const auto& v : arg)
            if (isNumeric(v)) out.push_back(toDouble(v));
        return out;
    }
}    // namespace

// =============================================================================
// Lexer
// =============================================================================

std::vector<XLToken> XLFormulaLexer::tokenize(std::string_view formula)
{
    std::vector<XLToken> tokens;

    // Strip leading '='
    std::size_t i = 0;
    if (!formula.empty() && formula[i] == '=') ++i;

    const auto len = formula.size();

    auto emit = [&](XLTokenKind k, std::string text = {}, double num = 0.0, bool b = false) {
        XLToken t;
        t.kind    = k;
        t.text    = std::move(text);
        t.number  = num;
        t.boolean = b;
        tokens.push_back(std::move(t));
    };

    while (i < len) {
        char c = formula[i];

        // --- Whitespace ---
        if (std::isspace(static_cast<unsigned char>(c))) {
            ++i;
            continue;
        }

        // --- String literal ---
        if (c == '"') {
            std::string s;
            ++i;
            while (i < len) {
                if (formula[i] == '"') {
                    ++i;
                    if (i < len && formula[i] == '"') {
                        s += '"';
                        ++i;
                    }    // escaped quote
                    else
                        break;
                }
                else {
                    s += formula[i++];
                }
            }
            emit(XLTokenKind::String, s);
            continue;
        }

        // --- Number ---
        if (std::isdigit(static_cast<unsigned char>(c)) ||
            (c == '.' && i + 1 < len && std::isdigit(static_cast<unsigned char>(formula[i + 1]))))
        {
            std::string numStr;
            while (i < len && (std::isdigit(static_cast<unsigned char>(formula[i])) || formula[i] == '.')) numStr += formula[i++];
            // Exponent
            if (i < len && (formula[i] == 'e' || formula[i] == 'E')) {
                numStr += formula[i++];
                if (i < len && (formula[i] == '+' || formula[i] == '-')) numStr += formula[i++];
                while (i < len && std::isdigit(static_cast<unsigned char>(formula[i]))) numStr += formula[i++];
            }
            double val = 0.0;
            try {
                val = std::stod(numStr);
            }
            catch (...) {
            }
            emit(XLTokenKind::Number, numStr, val);
            continue;
        }

        // --- Identifier or cell ref or bool ---
        if (std::isalpha(static_cast<unsigned char>(c)) || c == '$' || c == '_') {
            std::string ident;
            // Collect potentially qualified name: letters, digits, $, _, !
            while (i < len &&
                   (std::isalnum(static_cast<unsigned char>(formula[i])) || formula[i] == '$' || formula[i] == '_' || formula[i] == '!' || formula[i] == '.'))
            {
                ident += formula[i++];
            }
            // Check for range colon right after reference
            if (i < len && formula[i] == ':') {
                // Might be a range; collect the second ref
                std::string second;
                ++i;    // skip ':'
                while (i < len && (std::isalnum(static_cast<unsigned char>(formula[i])) || formula[i] == '$')) second += formula[i++];
                if (!second.empty()) { emit(XLTokenKind::CellRef, ident + ":" + second); }
                else {
                    emit(XLTokenKind::CellRef, ident);
                    emit(XLTokenKind::Colon);
                }
                continue;
            }

            // Boolean?
            std::string upper = ident;
            std::transform(upper.begin(), upper.end(), upper.begin(), ::toupper);
            if (upper == "TRUE") {
                emit(XLTokenKind::Bool, ident, 1.0, true);
                continue;
            }
            if (upper == "FALSE") {
                emit(XLTokenKind::Bool, ident, 0.0, false);
                continue;
            }

            // Cell reference? (letters then digits, optional $)
            // Pattern: optional$ + 1-3 letters + optional$ + 1-7 digits  (within ident already collected)
            {
                std::size_t j = 0;
                if (j < ident.size() && ident[j] == '$') ++j;
                std::size_t alphaStart = j;
                while (j < ident.size() && std::isalpha(static_cast<unsigned char>(ident[j]))) ++j;
                std::size_t alphaCount = j - alphaStart;
                if (j < ident.size() && ident[j] == '$') ++j;
                std::size_t digitStart = j;
                while (j < ident.size() && std::isdigit(static_cast<unsigned char>(ident[j]))) ++j;
                std::size_t digitCount = j - digitStart;
                
                // Check if next non-space is '(' to distinguish function from cell ref (like LOG10)
                std::size_t nextPos = i;
                while (nextPos < len && std::isspace(static_cast<unsigned char>(formula[nextPos]))) {
                    ++nextPos;
                }
                bool isFuncCall = (nextPos < len && formula[nextPos] == '(');

                if (alphaCount >= 1 && alphaCount <= 3 && digitCount >= 1 && j == ident.size() && !isFuncCall) {
                    emit(XLTokenKind::CellRef, ident);
                    continue;
                }
                // Sheet-qualified ref like "Sheet1!A1"
                auto bang = ident.find('!');
                if (bang != std::string::npos) {
                    emit(XLTokenKind::CellRef, ident);
                    continue;
                }
            }

            emit(XLTokenKind::Ident, ident);
            continue;
        }

        // --- Single or double char operators ---
        switch (c) {
            case '+':
                emit(XLTokenKind::Plus);
                ++i;
                break;
            case '-':
                emit(XLTokenKind::Minus);
                ++i;
                break;
            case '*':
                emit(XLTokenKind::Star);
                ++i;
                break;
            case '/':
                emit(XLTokenKind::Slash);
                ++i;
                break;
            case '^':
                emit(XLTokenKind::Caret);
                ++i;
                break;
            case '%':
                emit(XLTokenKind::Percent);
                ++i;
                break;
            case '&':
                emit(XLTokenKind::Amp);
                ++i;
                break;
            case '(':
                emit(XLTokenKind::LParen);
                ++i;
                break;
            case ')':
                emit(XLTokenKind::RParen);
                ++i;
                break;
            case ',':
                emit(XLTokenKind::Comma);
                ++i;
                break;
            case ';':
                emit(XLTokenKind::Semicolon);
                ++i;
                break;
            case ':':
                emit(XLTokenKind::Colon);
                ++i;
                break;
            case '=':
                emit(XLTokenKind::Eq);
                ++i;
                break;
            case '<':
                ++i;
                if (i < len && formula[i] == '=') {
                    emit(XLTokenKind::Le);
                    ++i;
                }
                else if (i < len && formula[i] == '>') {
                    emit(XLTokenKind::NEq);
                    ++i;
                }
                else
                    emit(XLTokenKind::Lt);
                break;
            case '>':
                ++i;
                if (i < len && formula[i] == '=') {
                    emit(XLTokenKind::Ge);
                    ++i;
                }
                else
                    emit(XLTokenKind::Gt);
                break;
            default:
                emit(XLTokenKind::Error, std::string(1, c));
                ++i;
                break;
        }
    }

    emit(XLTokenKind::End);
    return tokens;
}

// =============================================================================
// Parser helpers
// =============================================================================

const XLToken& XLFormulaParser::ParseContext::current() const
{
    Expects(pos < tokens.size());
    return tokens[pos];
}

const XLToken& XLFormulaParser::ParseContext::peek(std::size_t offset) const
{
    auto idx = pos + offset;
    return idx < tokens.size() ? tokens[idx] : tokens[tokens.size() - 1];    // return End if past
}

const XLToken& XLFormulaParser::ParseContext::consume()
{
    const XLToken& t = current();
    if (t.kind != XLTokenKind::End) ++pos;
    return t;
}

bool XLFormulaParser::ParseContext::matchKind(XLTokenKind k)
{
    if (current().kind == k) {
        consume();
        return true;
    }
    return false;
}

// Operator precedence (higher = binds tighter)
int XLFormulaParser::precedence(XLTokenKind k)
{
    switch (k) {
        case XLTokenKind::Amp:
            return 1;    // string concat
        case XLTokenKind::Eq:
        case XLTokenKind::NEq:
        case XLTokenKind::Lt:
        case XLTokenKind::Le:
        case XLTokenKind::Gt:
        case XLTokenKind::Ge:
            return 2;    // comparison
        case XLTokenKind::Plus:
        case XLTokenKind::Minus:
            return 3;    // add/sub
        case XLTokenKind::Star:
        case XLTokenKind::Slash:
            return 4;    // mul/div
        case XLTokenKind::Caret:
            return 5;    // power (right-assoc)
        default:
            return -1;    // not a binary operator
    }
}

bool XLFormulaParser::isRightAssoc(XLTokenKind k) { return k == XLTokenKind::Caret; }

std::unique_ptr<XLASTNode> XLFormulaParser::parse(gsl::span<const XLToken> tokens)
{
    ParseContext ctx{tokens, 0};
    auto         node = parseExpr(ctx, 0);
    return node;
}

// Pratt-style expression parser
std::unique_ptr<XLASTNode> XLFormulaParser::parseExpr(ParseContext& ctx, int minPrec)
{
    auto lhs = parseUnary(ctx);

    while (true) {
        const XLToken& op   = ctx.current();
        int            prec = precedence(op.kind);
        if (prec < minPrec) break;

        XLTokenKind opKind = op.kind;
        ctx.consume();

        int  nextPrec = isRightAssoc(opKind) ? prec : prec + 1;
        auto rhs      = parseExpr(ctx, nextPrec);

        auto binNode = std::make_unique<XLASTNode>(XLNodeKind::BinOp);
        binNode->op  = opKind;
        binNode->children.push_back(std::move(lhs));
        binNode->children.push_back(std::move(rhs));
        lhs = std::move(binNode);
    }

    return lhs;
}

std::unique_ptr<XLASTNode> XLFormulaParser::parseUnary(ParseContext& ctx)
{
    // Unary minus
    if (ctx.current().kind == XLTokenKind::Minus) {
        ctx.consume();
        auto operand = parseUnary(ctx);
        auto node    = std::make_unique<XLASTNode>(XLNodeKind::UnaryOp);
        node->op     = XLTokenKind::Minus;
        node->children.push_back(std::move(operand));
        return node;
    }
    // Unary plus (no-op)
    if (ctx.current().kind == XLTokenKind::Plus) {
        ctx.consume();
        return parseUnary(ctx);
    }
    auto expr = parsePrimary(ctx);
    // Postfix percent
    if (ctx.current().kind == XLTokenKind::Percent) {
        ctx.consume();
        auto node = std::make_unique<XLASTNode>(XLNodeKind::UnaryOp);
        node->op  = XLTokenKind::Percent;
        node->children.push_back(std::move(expr));
        return node;
    }
    return expr;
}

std::unique_ptr<XLASTNode> XLFormulaParser::parsePrimary(ParseContext& ctx)
{
    const XLToken& tok = ctx.current();

    if (tok.kind == XLTokenKind::Number) {
        ctx.consume();
        auto node    = std::make_unique<XLASTNode>(XLNodeKind::Number);
        node->number = tok.number;
        node->text   = tok.text;
        return node;
    }

    if (tok.kind == XLTokenKind::String) {
        ctx.consume();
        auto node  = std::make_unique<XLASTNode>(XLNodeKind::StringLit);
        node->text = tok.text;
        return node;
    }

    if (tok.kind == XLTokenKind::Bool) {
        ctx.consume();
        auto node     = std::make_unique<XLASTNode>(XLNodeKind::BoolLit);
        node->boolean = tok.boolean;
        if (ctx.current().kind == XLTokenKind::LParen) {
            ctx.consume(); // Eat '('
            if (ctx.current().kind == XLTokenKind::RParen) {
                ctx.consume(); // Eat ')'
            }
        }
        return node;
    }

    if (tok.kind == XLTokenKind::CellRef) {
        ctx.consume();
        bool isRange = (tok.text.find(':') != std::string::npos);
        auto node    = std::make_unique<XLASTNode>(isRange ? XLNodeKind::Range : XLNodeKind::CellRef);
        node->text   = tok.text;
        return node;
    }

    if (tok.kind == XLTokenKind::Ident) {
        std::string name = tok.text;
        ctx.consume();
        // Function call?
        if (ctx.current().kind == XLTokenKind::LParen) { return parseFuncCall(std::move(name), ctx); }
        // Named range / constant treated as cell ref (engine will try resolver)
        auto node  = std::make_unique<XLASTNode>(XLNodeKind::CellRef);
        node->text = name;
        return node;
    }

    if (tok.kind == XLTokenKind::LParen) {
        ctx.consume();
        auto inner = parseExpr(ctx, 0);
        if (ctx.current().kind == XLTokenKind::RParen) ctx.consume();
        return inner;
    }

    // Error literal (#DIV/0! etc.) - just capture the token as error
    auto node  = std::make_unique<XLASTNode>(XLNodeKind::ErrorLit);
    node->text = tok.text;
    ctx.consume();
    return node;
}

std::unique_ptr<XLASTNode> XLFormulaParser::parseFuncCall(std::string name, ParseContext& ctx)
{
    Expects(ctx.current().kind == XLTokenKind::LParen);
    ctx.consume();    // eat '('

    auto node = std::make_unique<XLASTNode>(XLNodeKind::FuncCall);
    // Store function name as uppercase
    std::transform(name.begin(), name.end(), name.begin(), ::toupper);
    node->text = std::move(name);

    // Parse argument list (comma or semicolon separated)
    while (ctx.current().kind != XLTokenKind::RParen && ctx.current().kind != XLTokenKind::End) {
        node->children.push_back(parseExpr(ctx, 0));
        if (ctx.current().kind == XLTokenKind::Comma || ctx.current().kind == XLTokenKind::Semicolon) { ctx.consume(); }
        else {
            break;
        }
    }
    if (ctx.current().kind == XLTokenKind::RParen) ctx.consume();
    return node;
}

// =============================================================================
// Range expansion helper
// =============================================================================

XLFormulaArg XLFormulaEngine::expandRange(std::string_view rangeRef, const XLCellResolver& resolver)
{
    if (!resolver) return XLFormulaArg();

    auto colonPos = rangeRef.find(':');
    if (colonPos == std::string_view::npos) {
        // Single cell
        return XLFormulaArg(resolver(rangeRef));
    }

    std::string startRef(rangeRef.substr(0, colonPos));
    std::string endRef(rangeRef.substr(colonPos + 1));

    std::string sheetName;
    auto exclPos = startRef.find('!');
    if (exclPos != std::string::npos) {
        sheetName = startRef.substr(0, exclPos);
        startRef = startRef.substr(exclPos + 1);
    }
    
    auto endExclPos = endRef.find('!');
    if (endExclPos != std::string::npos) {
        endRef = endRef.substr(endExclPos + 1);
    }

    // Strip optional absolute markers
    auto stripDollar = [](std::string& s) {
        s.erase(std::remove(s.begin(), s.end(), '$'), s.end());
    };
    stripDollar(startRef);
    stripDollar(endRef);

    try {
        auto     startCell = XLCellReference(startRef);
        auto     endCell   = XLCellReference(endRef);
        uint32_t r1 = startCell.row(), r2 = endCell.row();
        uint16_t c1 = startCell.column(), c2 = endCell.column();
        if (r1 > r2) std::swap(r1, r2);
        if (c1 > c2) std::swap(c1, c2);

        return XLFormulaArg(r1, r2, c1, c2, std::move(sheetName), &resolver);
    }
    catch (...) {
        XLCellValue err;
        err.setError("#REF!");
        return XLFormulaArg(std::move(err));
    }
}

// =============================================================================
// Evaluator – expandArg
// =============================================================================

XLFormulaArg XLFormulaEngine::expandArg(const XLASTNode& argNode, const XLCellResolver& resolver) const
{
    if (argNode.kind == XLNodeKind::Range) return expandRange(argNode.text, resolver);

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
            auto it = m_functions.find(node.text);
            if (it == m_functions.end()) return errName();

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

XLCellValue XLFormulaEngine::evaluate(std::string_view formula, const XLCellResolver& resolver) const
{
    if (formula.empty()) return XLCellValue{};
    try {
        auto tokens = XLFormulaLexer::tokenize(formula);
        auto ast    = XLFormulaParser::parse(gsl::span<const XLToken>(tokens));
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

XLFormulaEngine::XLFormulaEngine() { registerBuiltins(); }

void XLFormulaEngine::registerBuiltins()
{
    m_functions["SUM"]         = fnSum;
    m_functions["AVERAGE"]     = fnAverage;
    m_functions["AVG"]         = fnAverage;    // alias
    m_functions["MIN"]         = fnMin;
    m_functions["MAX"]         = fnMax;
    m_functions["COUNT"]       = fnCount;
    m_functions["COUNTA"]      = fnCounta;
    m_functions["IF"]          = fnIf;
    m_functions["IFS"]         = fnIfs;
    m_functions["SWITCH"]      = fnSwitch;
    m_functions["AND"]         = fnAnd;
    m_functions["OR"]          = fnOr;
    m_functions["NOT"]         = fnNot;
    m_functions["IFERROR"]     = fnIferror;
    m_functions["ABS"]         = fnAbs;
    m_functions["ROUND"]       = fnRound;
    m_functions["ROUNDUP"]     = fnRoundup;
    m_functions["ROUNDDOWN"]   = fnRounddown;
    m_functions["SQRT"]        = fnSqrt;
    m_functions["PI"]          = fnPi;
    m_functions["SIN"]         = fnSin;
    m_functions["COS"]         = fnCos;
    m_functions["TAN"]         = fnTan;
    m_functions["ASIN"]        = fnAsin;
    m_functions["ACOS"]        = fnAcos;
    m_functions["DEGREES"]     = fnDegrees;
    m_functions["RADIANS"]     = fnRadians;
    m_functions["RAND"]        = fnRand;
    m_functions["RANDBETWEEN"] = fnRandbetween;
    m_functions["INT"]         = fnInt;
    m_functions["MOD"]         = fnMod;
    m_functions["POWER"]       = fnPower;
    m_functions["VLOOKUP"]     = fnVlookup;
    m_functions["HLOOKUP"]     = fnHlookup;
    m_functions["XLOOKUP"]     = fnXlookup;
    m_functions["INDEX"]       = fnIndex;
    m_functions["MATCH"]       = fnMatch;
    m_functions["CONCATENATE"] = fnConcatenate;
    m_functions["CONCAT"]      = fnConcatenate;    // alias
    m_functions["LEN"]         = fnLen;
    m_functions["LEFT"]        = fnLeft;
    m_functions["RIGHT"]       = fnRight;
    m_functions["MID"]         = fnMid;
    m_functions["UPPER"]       = fnUpper;
    m_functions["LOWER"]       = fnLower;
    m_functions["TRIM"]        = fnTrim;
    m_functions["TEXT"]        = fnText;
    m_functions["ISNUMBER"]    = fnIsnumber;
    m_functions["ISBLANK"]     = fnIsblank;
    m_functions["ISERROR"]     = fnIserror;
    m_functions["ISTEXT"]      = fnIstext;

    // ---- Date / Time ----
    m_functions["TODAY"]       = fnToday;
    m_functions["NOW"]         = fnNow;
    m_functions["DATE"]        = fnDate;
    m_functions["TIME"]        = fnTime;
    m_functions["YEAR"]        = fnYear;
    m_functions["MONTH"]       = fnMonth;
    m_functions["DAY"]         = fnDay;
    m_functions["HOUR"]        = fnHour;
    m_functions["MINUTE"]      = fnMinute;
    m_functions["SECOND"]      = fnSecond;
    m_functions["DAYS"]         = fnDays;
    m_functions["_xlfn.DAYS"]   = fnDays;
    m_functions["WEEKDAY"]      = fnWeekday;
    m_functions["EDATE"]        = fnEdate;
    m_functions["EOMONTH"]      = fnEomonth;
    m_functions["WORKDAY"]      = fnWorkday;
    m_functions["NETWORKDAYS"]  = fnNetworkdays;

    // ---- Financial ----
    m_functions["PMT"] = fnPmt;
    m_functions["FV"]  = fnFv;
    m_functions["PV"]  = fnPv;
    m_functions["NPV"] = fnNpv;

    // ---- Math extended ----
    m_functions["SUMPRODUCT"] = fnSumproduct;
    m_functions["CEILING"]    = fnCeil;
    m_functions["CEIL"]       = fnCeil;
    m_functions["FLOOR"]      = fnFloor;
    m_functions["LOG"]        = fnLog;
    m_functions["LOG10"]      = fnLog10;
    m_functions["EXP"]        = fnExp;
    m_functions["SIGN"]       = fnSign;

    // ---- Text extended ----
    m_functions["FIND"]       = fnFind;
    m_functions["SEARCH"]     = fnSearch;
    m_functions["SUBSTITUTE"] = fnSubstitute;
    m_functions["REPLACE"]    = fnReplace;
    m_functions["REPT"]       = fnRept;
    m_functions["EXACT"]      = fnExact;
    m_functions["T"]          = fnT;
    m_functions["VALUE"]      = fnValue;
    m_functions["TEXTJOIN"]   = fnTextjoin;
    m_functions["_xlfn.TEXTJOIN"] = fnTextjoin;
    m_functions["CLEAN"]      = fnClean;
    m_functions["PROPER"]     = fnProper;

    // ---- Statistical / Conditional ----
    m_functions["SUMIF"]        = fnSumif;
    m_functions["COUNTIF"]      = fnCountif;
    m_functions["SUMIFS"]       = fnSumifs;
    m_functions["COUNTIFS"]     = fnCountifs;
    m_functions["MAXIFS"]       = fnMaxifs;
    m_functions["_xlfn.MAXIFS"] = fnMaxifs;
    m_functions["MINIFS"]       = fnMinifs;
    m_functions["_xlfn.MINIFS"] = fnMinifs;
    m_functions["AVERAGEIF"]    = fnAverageif;
    m_functions["RANK"]       = fnRank;
    m_functions["RANK.EQ"]    = fnRank;
    m_functions["LARGE"]      = fnLarge;
    m_functions["SMALL"]      = fnSmall;
    m_functions["STDEV"]      = fnStdev;
    m_functions["STDEV.S"]    = fnStdev;
    m_functions["VAR"]        = fnVar;
    m_functions["VAR.S"]      = fnVar;
    m_functions["MEDIAN"]     = fnMedian;
    m_functions["COUNTBLANK"] = fnCountblank;

    // ---- Info extended ----
    m_functions["ISNA"]      = fnIsna;
    m_functions["IFNA"]      = fnIfna;
    m_functions["ISLOGICAL"] = fnIslogical;
    m_functions["ISNONTEXT"] = fnIsnontext;

    // ---- Easy Additions ----
    m_functions["TRUE"]            = fnTrue;
    m_functions["FALSE"]           = fnFalse;
    m_functions["ISEVEN"]          = fnIseven;
    m_functions["ISODD"]           = fnIsodd;
    m_functions["MROUND"]          = fnMround;
    m_functions["CEILING.MATH"]    = fnCeilingMath;
    m_functions["_xlfn.CEILING.MATH"] = fnCeilingMath;
    m_functions["FLOOR.MATH"]      = fnFloorMath;
    m_functions["_xlfn.FLOOR.MATH"]= fnFloorMath;
    m_functions["VAR.P"]           = fnVarp;
    m_functions["_xlfn.VAR.P"]     = fnVarp;
    m_functions["VARP"]            = fnVarp;
    m_functions["STDEV.P"]         = fnStdevp;
    m_functions["_xlfn.STDEV.P"]   = fnStdevp;
    m_functions["STDEVP"]          = fnStdevp;
    m_functions["VARA"]            = fnVara;
    m_functions["VARPA"]           = fnVarpa;
    m_functions["STDEVA"]          = fnStdeva;
    m_functions["STDEVPA"]         = fnStdevpa;
    m_functions["PERMUT"]          = fnPermut;
    m_functions["PERMUTATIONA"]    = fnPermutationa;
    m_functions["_xlfn.PERMUTATIONA"] = fnPermutationa;
    m_functions["FISHER"]          = fnFisher;
    m_functions["FISHERINV"]       = fnFisherinv;
    m_functions["STANDARDIZE"]     = fnStandardize;
    m_functions["PEARSON"]         = fnPearson;
    m_functions["CORREL"]          = fnPearson;
    m_functions["COVAR"]           = fnCovarianceP;
    m_functions["COVARIANCE.P"]    = fnCovarianceP;
    m_functions["COVARIANCE.S"]    = fnCovarianceS;
    m_functions["PERCENTILE"]      = fnPercentileInc;
    m_functions["PERCENTILE.INC"]  = fnPercentileInc;
    m_functions["PERCENTILE.EXC"]  = fnPercentileExc;
    m_functions["QUARTILE"]        = fnQuartileInc;
    m_functions["QUARTILE.INC"]    = fnQuartileInc;
    m_functions["QUARTILE.EXC"]    = fnQuartileExc;
    m_functions["TRIMMEAN"]        = fnTrimmean;
    m_functions["SLOPE"]           = fnSlope;
    m_functions["INTERCEPT"]       = fnIntercept;
    m_functions["RSQ"]             = fnRsq;
    m_functions["AVERAGEIFS"]      = fnAverageifs;
    m_functions["ISOWEEKNUM"]      = fnIsoweeknum;
    m_functions["_xlfn.ISOWEEKNUM"]= fnIsoweeknum;
    m_functions["WEEKNUM"]         = fnWeeknum;
    m_functions["DAYS360"]         = fnDays360;
    m_functions["NPER"]            = fnNper;
    m_functions["DB"]              = fnDb;
    m_functions["DDB"]             = fnDdb;

    // Fix mapped functions from earlier additions
    m_functions["ISERR"]           = fnIserr;
    m_functions["TRUNC"]           = fnTrunc;
    m_functions["SUMSQ"]           = fnSumsq;
    m_functions["SUMX2MY2"]        = fnSumx2my2;
    m_functions["SUMX2PY2"]        = fnSumx2py2;
    m_functions["SUMXMY2"]         = fnSumxmy2;
    m_functions["AVEDEV"]          = fnAvedev;
    m_functions["DEVSQ"]           = fnDevsq;
    m_functions["AVERAGEA"]        = fnAveragea;

    // Missing checklist functions
    m_functions["SLN"]             = fnSln;
    m_functions["SYD"]             = fnSyd;
    m_functions["CHAR"]            = fnChar;
    m_functions["UNICHAR"]         = fnUnichar;
    m_functions["CODE"]            = fnCode;
    m_functions["UNICODE"]         = fnUnicode;
    m_functions["NOW"]             = fnNow;
    m_functions["TODAY"]           = fnToday;
    m_functions["TRUE"]            = fnTrue;
    m_functions["FALSE"]           = fnFalse;
}

// =============================================================================
// Built-in: Math / Statistical
// =============================================================================

XLCellValue XLFormulaEngine::fnSum(const std::vector<XLFormulaArg>& args)
{
    double total = 0.0;
    for (const auto& v : numerics(args)) total += v;
    return XLCellValue(total);
}

XLCellValue XLFormulaEngine::fnAverage(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errDiv0();
    double total = 0.0;
    for (double v : nums) total += v;
    return XLCellValue(total / static_cast<double>(nums.size()));
}

XLCellValue XLFormulaEngine::fnMin(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return XLCellValue(0.0);
    return XLCellValue(*std::min_element(nums.begin(), nums.end()));
}

XLCellValue XLFormulaEngine::fnMax(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return XLCellValue(0.0);
    return XLCellValue(*std::max_element(nums.begin(), nums.end()));
}

XLCellValue XLFormulaEngine::fnCount(const std::vector<XLFormulaArg>& args)
{
    int64_t cnt = 0;
    for (const auto& arg : args) {
        for (const auto& v : arg) {
            if (isNumeric(v)) ++cnt;
        }
    }
    return XLCellValue(cnt);
}

XLCellValue XLFormulaEngine::fnCounta(const std::vector<XLFormulaArg>& args)
{
    int64_t cnt = 0;
    for (const auto& arg : args) {
        for (const auto& v : arg) {
            if (!isEmpty(v)) {
                if (v.type() == XLValueType::String && v.get<std::string>().empty()) {
                    continue;
                }
                ++cnt;
            }
        }
    }
    return XLCellValue(cnt);
}

XLCellValue XLFormulaEngine::fnAbs(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    const auto& v = args[0][0];
    if (!isNumeric(v)) return errValue();
    return XLCellValue(std::abs(toDouble(v)));
}

XLCellValue XLFormulaEngine::fnSqrt(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double d = toDouble(args[0][0]);
    if (d < 0.0) return errNum();
    return XLCellValue(std::sqrt(d));
}

XLCellValue XLFormulaEngine::fnPi(const std::vector<XLFormulaArg>& args)
{
    (void)args;
    return XLCellValue(3.14159265358979323846);
}

XLCellValue XLFormulaEngine::fnSin(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::sin(toDouble(args[0][0])));
}

XLCellValue XLFormulaEngine::fnCos(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::cos(toDouble(args[0][0])));
}

XLCellValue XLFormulaEngine::fnTan(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::tan(toDouble(args[0][0])));
}

XLCellValue XLFormulaEngine::fnAsin(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double v = toDouble(args[0][0]);
    if (v < -1.0 || v > 1.0) return errNum();
    return XLCellValue(std::asin(v));
}

XLCellValue XLFormulaEngine::fnAcos(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double v = toDouble(args[0][0]);
    if (v < -1.0 || v > 1.0) return errNum();
    return XLCellValue(std::acos(v));
}

XLCellValue XLFormulaEngine::fnDegrees(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(toDouble(args[0][0]) * 180.0 / 3.14159265358979323846);
}

XLCellValue XLFormulaEngine::fnRadians(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(toDouble(args[0][0]) * 3.14159265358979323846 / 180.0);
}

XLCellValue XLFormulaEngine::fnRand(const std::vector<XLFormulaArg>& args)
{
    (void)args;
    return XLCellValue(static_cast<double>(std::rand()) / RAND_MAX);
}

XLCellValue XLFormulaEngine::fnRandbetween(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty() || !isNumeric(args[0][0]) || !isNumeric(args[1][0])) return errValue();
    int64_t low = static_cast<int64_t>(std::ceil(toDouble(args[0][0])));
    int64_t high = static_cast<int64_t>(std::floor(toDouble(args[1][0])));
    if (low > high) return errNum();
    return XLCellValue(low + (std::rand() % (high - low + 1)));
}

XLCellValue XLFormulaEngine::fnInt(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    if (!isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(std::floor(toDouble(args[0][0]))));
}

XLCellValue XLFormulaEngine::fnRound(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num    = toDouble(args[0][0]);
    int    digits = static_cast<int>(toDouble(args[1][0]));
    double factor = std::pow(10.0, digits);
    return XLCellValue(std::round(num * factor) / factor);
}

XLCellValue XLFormulaEngine::fnRoundup(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num    = toDouble(args[0][0]);
    int    digits = static_cast<int>(toDouble(args[1][0]));
    double factor = std::pow(10.0, digits);
    return XLCellValue(num >= 0 ? std::ceil(num * factor) / factor : std::floor(num * factor) / factor);
}

XLCellValue XLFormulaEngine::fnRounddown(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num    = toDouble(args[0][0]);
    int    digits = static_cast<int>(toDouble(args[1][0]));
    double factor = std::pow(10.0, digits);
    return XLCellValue(num >= 0 ? std::floor(num * factor) / factor : std::ceil(num * factor) / factor);
}

XLCellValue XLFormulaEngine::fnMod(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double n = toDouble(args[0][0]), d = toDouble(args[1][0]);
    if (d == 0.0) return errDiv0();
    return XLCellValue(std::fmod(n, d));
}

XLCellValue XLFormulaEngine::fnPower(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    return XLCellValue(std::pow(toDouble(args[0][0]), toDouble(args[1][0])));
}

// =============================================================================
// Built-in: Logical
// =============================================================================

XLCellValue XLFormulaEngine::fnIf(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnAnd(const std::vector<XLFormulaArg>& args)
{
    for (const auto& arg : args) {
        for (const auto& v : arg) {
            if (!isNumeric(v) && v.type() != XLValueType::Boolean) return errValue();
            if (!toDouble(v)) return XLCellValue(false);
        }
    }
    return XLCellValue(true);
}

XLCellValue XLFormulaEngine::fnOr(const std::vector<XLFormulaArg>& args)
{
    for (const auto& arg : args) {
        for (const auto& v : arg) {
            if (!isNumeric(v) && v.type() != XLValueType::Boolean) return errValue();
            if (toDouble(v)) return XLCellValue(true);
        }
    }
    return XLCellValue(false);
}

XLCellValue XLFormulaEngine::fnNot(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    const auto& v = args[0][0];
    if (!isNumeric(v) && v.type() != XLValueType::Boolean) return errValue();
    return XLCellValue(!static_cast<bool>(toDouble(v)));
}

XLCellValue XLFormulaEngine::fnIferror(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    return isError(args[0][0]) ? args[1][0] : args[0][0];
}

XLCellValue XLFormulaEngine::fnIfs(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnSwitch(const std::vector<XLFormulaArg>& args)
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

// =============================================================================
// Built-in: Lookup
// =============================================================================

XLCellValue XLFormulaEngine::fnVlookup(const std::vector<XLFormulaArg>& args)
{
    // VLOOKUP(lookup_value, table_array, col_index, [exact_match])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();

    const XLCellValue& lookupVal = args[0][0];
    const auto&        table     = args[1];
    int                colIdx    = static_cast<int>(toDouble(args[2][0]));    // 1-based
    bool               exact     = (args.size() < 4 || args[3].empty()) ? true : (toDouble(args[3][0]) == 0.0);

    if (colIdx < 1) return errValue();

    int nCols = static_cast<int>(table.cols());
    if (colIdx > nCols) return errRef();
    
    int nRows = static_cast<int>(table.rows());

    // Linear scan column 0 of each row
    for (int r = 0; r < nRows; ++r) {
        int cellIdx = r * nCols;
        if (cellIdx >= static_cast<int>(table.size())) break;
        const XLCellValue& key = table[static_cast<std::size_t>(cellIdx)];

        bool match = false;
        if (exact) {
            if (isNumeric(key) && isNumeric(lookupVal))
                match = (toDouble(key) == toDouble(lookupVal));
            else
                match = (toString(key) == toString(lookupVal));
        }
        else {
            // Approximate: find largest key <= lookupVal
            if (isNumeric(key) && isNumeric(lookupVal)) match = (toDouble(key) <= toDouble(lookupVal));
        }

        if (match) {
            int resultIdx = r * nCols + (colIdx - 1);
            if (resultIdx < static_cast<int>(table.size())) return table[static_cast<std::size_t>(resultIdx)];
            return errRef();
        }
    }
    return errNA();
}

XLCellValue XLFormulaEngine::fnHlookup(const std::vector<XLFormulaArg>& args)
{
    // HLOOKUP(lookup_value, table_array, row_index, [exact_match])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();

    const XLCellValue& lookupVal = args[0][0];
    const auto&        table     = args[1];
    int                rowIdx    = static_cast<int>(toDouble(args[2][0]));    // 1-based row to return
    bool               exact     = (args.size() < 4 || args[3].empty()) ? true : (toDouble(args[3][0]) == 0.0);

    if (rowIdx < 1) return errValue();

    int nRows = static_cast<int>(table.rows());
    if (rowIdx > nRows) return errRef();
    
    int nCols = static_cast<int>(table.cols());

    // HLOOKUP searches the first row.
    for (int c = 0; c < nCols; ++c) {
        if (c >= static_cast<int>(table.size())) break;
        const XLCellValue& key   = table[static_cast<std::size_t>(c)];
        bool               match = false;
        if (exact) {
            if (isNumeric(key) && isNumeric(lookupVal))
                match = (toDouble(key) == toDouble(lookupVal));
            else
                match = (toString(key) == toString(lookupVal));
        }
        else {
            if (isNumeric(key) && isNumeric(lookupVal)) match = (toDouble(key) <= toDouble(lookupVal));
        }
        if (match) {
            std::size_t resultIdx = static_cast<std::size_t>(c + (rowIdx - 1) * nCols);
            if (resultIdx < table.size()) return table[resultIdx];
            return errRef();
        }
    }
    return errNA();
}

XLCellValue XLFormulaEngine::fnXlookup(const std::vector<XLFormulaArg>& args)
{
    // XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();

    const auto& lookupVal = args[0][0];
    const auto& lookupArr = args[1];
    const auto& returnArr = args[2];

    int matchMode  = (args.size() > 4 && !args[4].empty()) ? static_cast<int>(toDouble(args[4][0])) : 0;
    int searchMode = (args.size() > 5 && !args[5].empty()) ? static_cast<int>(toDouble(args[5][0])) : 1;

    int bestMatchIdx = -1;

    auto compareValues = [matchMode](const XLCellValue& key, const XLCellValue& val) -> double {
        if (isNumeric(key) && isNumeric(val)) { return toDouble(key) - toDouble(val); }
        else {
            // For exact matches, if types mismatch entirely, don't loosely convert them unless they're both strings
            if (matchMode == 0 && (isNumeric(key) != isNumeric(val))) {
                return std::numeric_limits<double>::infinity();    // Force failure
            }

            std::string sKey = toString(key);
            std::string sVal = toString(val);
            std::transform(sKey.begin(), sKey.end(), sKey.begin(), ::tolower);
            std::transform(sVal.begin(), sVal.end(), sVal.begin(), ::tolower);
            int cmp = sKey.compare(sVal);
            return cmp < 0 ? -1.0 : (cmp > 0 ? 1.0 : 0.0);
        }
    };

    if (searchMode == 2 || searchMode == -2) {
        // Binary search (2 = ascending, -2 = descending)
        int low  = 0;
        int high = static_cast<int>(lookupArr.size()) - 1;

        while (low <= high) {
            int         mid  = low + (high - low) / 2;
            const auto& key  = lookupArr[static_cast<std::size_t>(mid)];
            double      diff = compareValues(key, lookupVal);

            if (diff == 0.0) {
                bestMatchIdx = mid;
                break;    // Exact match found
            }

            if (searchMode == 2) {                              // Ascending array
                if (diff < 0) {                                 // key < lookupVal
                    if (matchMode == -1) bestMatchIdx = mid;    // Candidate for exact or next smaller
                    low = mid + 1;
                }
                else {                                         // key > lookupVal
                    if (matchMode == 1) bestMatchIdx = mid;    // Candidate for exact or next larger
                    high = mid - 1;
                }
            }
            else {                                              // Descending array (searchMode == -2)
                if (diff < 0) {                                 // key < lookupVal
                    if (matchMode == -1) bestMatchIdx = mid;    // Candidate for exact or next smaller
                    high = mid - 1;                             // smaller values are to the left in descending
                }
                else {                                         // key > lookupVal
                    if (matchMode == 1) bestMatchIdx = mid;    // Candidate for exact or next larger
                    low = mid + 1;                             // larger values are to the left, so smaller-than-current are right
                }
            }
        }
    }
    else {
        // Linear search (1 = first-to-last, -1 = last-to-first)
        int startIdx = (searchMode == -1) ? static_cast<int>(lookupArr.size()) - 1 : 0;
        int endIdx   = (searchMode == -1) ? -1 : static_cast<int>(lookupArr.size());
        int step     = (searchMode == -1) ? -1 : 1;

        double bestDiff = std::numeric_limits<double>::infinity();

        for (int i = startIdx; i != endIdx; i += step) {
            const auto& key = lookupArr[static_cast<std::size_t>(i)];

            if (matchMode == 0) {    // Exact match
                if (compareValues(key, lookupVal) == 0.0) {
                    bestMatchIdx = i;
                    break;
                }
            }
            else if (matchMode == -1) {                         // Exact match or next smaller item
                double diff = compareValues(lookupVal, key);    // lookupVal - key
                if (diff == 0.0) {
                    bestMatchIdx = i;
                    break;
                }
                else if (diff > 0 && diff < bestDiff) {
                    bestDiff     = diff;
                    bestMatchIdx = i;
                }
            }
            else if (matchMode == 1) {                          // Exact match or next larger item
                double diff = compareValues(key, lookupVal);    // key - lookupVal
                if (diff == 0.0) {
                    bestMatchIdx = i;
                    break;
                }
                else if (diff > 0 && diff < bestDiff) {
                    bestDiff     = diff;
                    bestMatchIdx = i;
                }
            }
            // wildcard match_mode (2) is not fully implemented here yet
        }
    }

    if (bestMatchIdx != -1) {
        if (static_cast<std::size_t>(bestMatchIdx) < returnArr.size()) { return returnArr[static_cast<std::size_t>(bestMatchIdx)]; }
        return errRef();    // Return array smaller than lookup array
    }

    // Not found
    if (args.size() > 3 && !args[3].empty()) { return args[3][0]; }
    return errNA();
}

XLCellValue XLFormulaEngine::fnIndex(const std::vector<XLFormulaArg>& args)
{
    // INDEX(array, row_num, [col_num])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const auto& arr = args[0];
    int         r   = static_cast<int>(toDouble(args[1][0]));
    int         c   = (args.size() > 2 && !args[2].empty()) ? static_cast<int>(toDouble(args[2][0])) : 1;
    if (r < 1 || c < 1) return errValue();
    // We don't know nCols; assume 1D array or use c=1
    std::size_t idx = static_cast<std::size_t>(r - 1);
    if (c > 1) {
        // Can't determine nCols without metadata; best effort
        idx = static_cast<std::size_t>((r - 1) * c + (c - 1));
    }
    if (idx >= arr.size()) return errRef();
    return arr[idx];
}

XLCellValue XLFormulaEngine::fnMatch(const std::vector<XLFormulaArg>& args)
{
    // MATCH(lookup_value, lookup_array, [match_type])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const XLCellValue& lookupVal = args[0][0];
    const auto&        arr       = args[1];
    int                matchType = (args.size() > 2 && !args[2].empty()) ? static_cast<int>(toDouble(args[2][0])) : 1;

    if (matchType == 0) {
        // Exact match
        for (std::size_t i = 0; i < arr.size(); ++i) {
            bool match = false;
            if (isNumeric(arr[i]) && isNumeric(lookupVal))
                match = (toDouble(arr[i]) == toDouble(lookupVal));
            else
                match = (toString(arr[i]) == toString(lookupVal));
            if (match) return XLCellValue(static_cast<int64_t>(i + 1));
        }
        return errNA();
    }
    // Approximate match (matchType == 1: largest <= lookup)
    int64_t bestIdx = -1;
    double  bestVal = std::numeric_limits<double>::lowest();
    for (std::size_t i = 0; i < arr.size(); ++i) {
        if (isNumeric(arr[i]) && isNumeric(lookupVal)) {
            double d = toDouble(arr[i]);
            if (d <= toDouble(lookupVal) && d >= bestVal) {
                bestVal = d;
                bestIdx = static_cast<int64_t>(i + 1);
            }
        }
    }
    if (bestIdx < 0) return errNA();
    return XLCellValue(bestIdx);
}

// =============================================================================
// Built-in: Text
// =============================================================================

XLCellValue XLFormulaEngine::fnConcatenate(const std::vector<XLFormulaArg>& args)
{
    std::string result;
    for (const auto& arg : args) {
        for (const auto& v : arg) { result += toString(v); }
    }
    return XLCellValue(result);
}

XLCellValue XLFormulaEngine::fnLen(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    return XLCellValue(static_cast<int64_t>(toString(args[0][0]).size()));
}

XLCellValue XLFormulaEngine::fnLeft(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    int64_t     n = (args.size() > 1 && !args[1].empty()) ? static_cast<int64_t>(toDouble(args[1][0])) : 1;
    if (n < 0) return errValue();
    return XLCellValue(s.substr(0, static_cast<std::size_t>(n)));
}

XLCellValue XLFormulaEngine::fnRight(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    int64_t     n = (args.size() > 1 && !args[1].empty()) ? static_cast<int64_t>(toDouble(args[1][0])) : 1;
    if (n < 0) return errValue();
    auto len   = static_cast<int64_t>(s.size());
    auto start = len - n;
    if (start < 0) start = 0;
    return XLCellValue(s.substr(static_cast<std::size_t>(start)));
}

XLCellValue XLFormulaEngine::fnMid(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    std::string s     = toString(args[0][0]);
    int64_t     start = static_cast<int64_t>(toDouble(args[1][0])) - 1;    // 1-based
    int64_t     count = static_cast<int64_t>(toDouble(args[2][0]));
    if (start < 0 || count < 0) return errValue();
    return XLCellValue(s.substr(static_cast<std::size_t>(start), static_cast<std::size_t>(count)));
}

XLCellValue XLFormulaEngine::fnUpper(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    std::transform(s.begin(), s.end(), s.begin(), ::toupper);
    return XLCellValue(s);
}

XLCellValue XLFormulaEngine::fnLower(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    std::transform(s.begin(), s.end(), s.begin(), ::tolower);
    return XLCellValue(s);
}

XLCellValue XLFormulaEngine::fnTrim(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    return XLCellValue(strTrim(toString(args[0][0])));
}

XLCellValue XLFormulaEngine::fnText(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::string format = toString(args[1][0]);
    
    if (isNumeric(args[0][0])) {
        // Very basic date format support
        std::string lowerFmt = format;
        std::transform(lowerFmt.begin(), lowerFmt.end(), lowerFmt.begin(), ::tolower);
        if (lowerFmt == "yyyy-mm-dd" || lowerFmt == "yyyy/mm/dd") {
            try {
                std::tm t = XLDateTime(toDouble(args[0][0])).tm();
                char buf[64];
                snprintf(buf, sizeof(buf), "%04d-%02d-%02d", t.tm_year + 1900, t.tm_mon + 1, t.tm_mday);
                if (lowerFmt == "yyyy/mm/dd") buf[4] = buf[7] = '/';
                return XLCellValue(std::string(buf));
            } catch (...) {
                return errValue();
            }
        }
    }
    
    // Simplified fallback: just convert to string
    return XLCellValue(toString(args[0][0]));
}

// =============================================================================
// Built-in: Info
// =============================================================================

XLCellValue XLFormulaEngine::fnIsnumber(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    return XLCellValue(isNumeric(args[0][0]));
}

XLCellValue XLFormulaEngine::fnIsblank(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(true);
    return XLCellValue(isEmpty(args[0][0]));
}

XLCellValue XLFormulaEngine::fnIserror(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    return XLCellValue(isError(args[0][0]));
}

XLCellValue XLFormulaEngine::fnIstext(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    return XLCellValue(args[0][0].type() == XLValueType::String);
}

// =============================================================================
// Built-in: Date / Time
// =============================================================================

namespace
{
    /// Convert serialNumber → std::tm (no-throw; returns epoch on failure)
    std::tm serialToTm(double serial)
    {
        try {
            return XLDateTime(serial).tm();
        }
        catch (...) {
            return {};
        }
    }

    /// Convert std::tm → Excel serial number
    double tmToSerial(std::tm t)
    {
        try {
            return XLDateTime(t).serial();
        }
        catch (...) {
            return 0.0;
        }
    }

    /// Number of days in a given month/year
    int daysInMonth(int year, int month)
    {
        // Use std::array to satisfy clang-tidy's constant-expression subscript requirement
        constexpr std::array<int, 12> days = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
        bool                          leap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
        int                           d    = days.at(static_cast<std::size_t>(month - 1));
        if (month == 2 && leap) ++d;
        return d;
    }

    /// True if the Excel serial number falls on a Saturday or Sunday (weekday mode 1)
    bool isWeekend(double serial)
    {
        std::tm t = serialToTm(serial);
        return t.tm_wday == 0 || t.tm_wday == 6;    // 0=Sun, 6=Sat
    }

    // -------------------------------------------------------------------------
    // Shared criteria matcher for SUMIF / COUNTIF family
    // Supported operators: =, <>, <, <=, >, >=, and wildcard * / ?
    // -------------------------------------------------------------------------
    bool matchesCriteria(const XLCellValue& cell, const std::string& criteria)
    {
        if (criteria.empty()) return false;

        // Detect relational operator prefix
        std::string_view op;
        std::string_view rhs;
        auto             s = std::string_view(criteria);
        if (s.size() >= 2 && s[0] == '<' && s[1] == '>') {
            op  = "<>";
            rhs = s.substr(2);
        }
        else if (s.size() >= 2 && s[0] == '<' && s[1] == '=') {
            op  = "<=";
            rhs = s.substr(2);
        }
        else if (s.size() >= 2 && s[0] == '>' && s[1] == '=') {
            op  = ">=";
            rhs = s.substr(2);
        }
        else if (s[0] == '<') {
            op  = "<";
            rhs = s.substr(1);
        }
        else if (s[0] == '>') {
            op  = ">";
            rhs = s.substr(1);
        }
        else if (s[0] == '=') {
            op  = "=";
            rhs = s.substr(1);
        }
        else {
            op  = "=";
            rhs = s;
        }    // no prefix → exact / wildcard

        // Try numeric comparison first
        double rhsNum   = 0.0;
        bool   rhsIsNum = false;
        try {
            std::size_t idx = 0;
            rhsNum          = std::stod(std::string(rhs), &idx);
            rhsIsNum        = (idx == rhs.size());
        }
        catch (...) {
        }

        if (rhsIsNum && isNumeric(cell)) {
            double lhsNum = toDouble(cell);
            if (op == "=") return lhsNum == rhsNum;
            if (op == "<>") return lhsNum != rhsNum;
            if (op == "<") return lhsNum < rhsNum;
            if (op == "<=") return lhsNum <= rhsNum;
            if (op == ">") return lhsNum > rhsNum;
            if (op == ">=") return lhsNum >= rhsNum;
        }

        // String comparison with optional wildcard (* and ?)
        std::string cellStr = toString(cell);
        std::string rhsStr  = std::string(rhs);
        // Case-insensitive
        std::string cellLo = cellStr, rhsLo = rhsStr;
        std::transform(cellLo.begin(), cellLo.end(), cellLo.begin(), ::tolower);
        std::transform(rhsLo.begin(), rhsLo.end(), rhsLo.begin(), ::tolower);

        // Wildcard match (only for equality operators)
        bool hasWildcard   = rhsLo.find('*') != std::string::npos || rhsLo.find('?') != std::string::npos;
        auto wildcardMatch = [](const std::string& text, const std::string& pattern) -> bool {
            // Simple recursive wildcard: * matches any sequence, ? matches one character
            std::size_t t = 0, p = 0;
            std::size_t starT = std::string::npos, starP = std::string::npos;
            while (t < text.size()) {
                if (p < pattern.size() && (pattern[p] == '?' || pattern[p] == text[t])) {
                    ++t;
                    ++p;
                }
                else if (p < pattern.size() && pattern[p] == '*') {
                    starP = p++;
                    starT = t;
                }
                else if (starP != std::string::npos) {
                    p = starP + 1;
                    t = ++starT;
                }
                else
                    return false;
            }
            while (p < pattern.size() && pattern[p] == '*') ++p;
            return p == pattern.size();
        };

        if (op == "=" || (op == "=" && !hasWildcard)) return hasWildcard ? wildcardMatch(cellLo, rhsLo) : cellLo == rhsLo;
        if (op == "<>") return hasWildcard ? !wildcardMatch(cellLo, rhsLo) : cellLo != rhsLo;
        if (op == "<") return cellLo < rhsLo;
        if (op == "<=") return cellLo <= rhsLo;
        if (op == ">") return cellLo > rhsLo;
        if (op == ">=") return cellLo >= rhsLo;
        return false;
    }
}    // namespace

XLCellValue XLFormulaEngine::fnToday(const std::vector<XLFormulaArg>&)
{ return XLCellValue(XLDateTime::now().serial() - std::fmod(XLDateTime::now().serial(), 1.0)); }

XLCellValue XLFormulaEngine::fnNow(const std::vector<XLFormulaArg>&) { return XLCellValue(XLDateTime::now().serial()); }

XLCellValue XLFormulaEngine::fnDate(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    std::tm t{};
    t.tm_year = static_cast<int>(toDouble(args[0][0])) - 1900;
    t.tm_mon  = static_cast<int>(toDouble(args[1][0])) - 1;
    t.tm_mday = static_cast<int>(toDouble(args[2][0]));
    return XLCellValue(tmToSerial(t));
}

XLCellValue XLFormulaEngine::fnYear(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(serialToTm(toDouble(args[0][0])).tm_year + 1900));
}

XLCellValue XLFormulaEngine::fnMonth(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(serialToTm(toDouble(args[0][0])).tm_mon + 1));
}

XLCellValue XLFormulaEngine::fnDay(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(serialToTm(toDouble(args[0][0])).tm_mday));
}

XLCellValue XLFormulaEngine::fnHour(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(serialToTm(toDouble(args[0][0])).tm_hour));
}

XLCellValue XLFormulaEngine::fnMinute(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(serialToTm(toDouble(args[0][0])).tm_min));
}

XLCellValue XLFormulaEngine::fnSecond(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(serialToTm(toDouble(args[0][0])).tm_sec));
}

XLCellValue XLFormulaEngine::fnTime(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    if (!isNumeric(args[0][0]) || !isNumeric(args[1][0]) || !isNumeric(args[2][0])) return errValue();
    
    int64_t h = static_cast<int64_t>(toDouble(args[0][0]));
    int64_t m = static_cast<int64_t>(toDouble(args[1][0]));
    int64_t s = static_cast<int64_t>(toDouble(args[2][0]));
    
    double fraction = (h * 3600.0 + m * 60.0 + s) / 86400.0;
    
    // Normalize to 0-1
    fraction = fraction - std::floor(fraction);
    if (fraction < 0) fraction += 1.0;
    
    return XLCellValue(fraction);
}

XLCellValue XLFormulaEngine::fnDays(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    if (!isNumeric(args[0][0]) || !isNumeric(args[1][0])) return errValue();
    
    double end_date = toDouble(args[0][0]);
    double start_date = toDouble(args[1][0]);
    
    return XLCellValue(end_date - start_date);
}

XLCellValue XLFormulaEngine::fnWeekday(const std::vector<XLFormulaArg>& args)
{
    // WEEKDAY(serial, [return_type])
    // return_type 1 (default): 1=Sunday … 7=Saturday
    // return_type 2: 1=Monday … 7=Sunday
    // return_type 3: 0=Monday … 6=Sunday
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    int mode = (args.size() > 1 && !args[1].empty()) ? static_cast<int>(toDouble(args[1][0])) : 1;
    int wday = serialToTm(toDouble(args[0][0])).tm_wday;    // 0=Sun … 6=Sat
    switch (mode) {
        case 2:
            return XLCellValue(static_cast<int64_t>(wday == 0 ? 7 : wday));
        case 3:
            return XLCellValue(static_cast<int64_t>(wday == 0 ? 6 : wday - 1));
        default:
            return XLCellValue(static_cast<int64_t>(wday + 1));
    }
}

XLCellValue XLFormulaEngine::fnEdate(const std::vector<XLFormulaArg>& args)
{
    // EDATE(start_date, months) – same day, N months later
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::tm t      = serialToTm(toDouble(args[0][0]));
    int     months = static_cast<int>(toDouble(args[1][0]));
    t.tm_mon += months;
    // Normalise overflow
    t.tm_year += t.tm_mon / 12;
    t.tm_mon = t.tm_mon % 12;
    if (t.tm_mon < 0) {
        t.tm_year--;
        t.tm_mon += 12;
    }
    // Clamp day to month end
    int maxDay = daysInMonth(t.tm_year + 1900, t.tm_mon + 1);
    if (t.tm_mday > maxDay) t.tm_mday = maxDay;
    return XLCellValue(tmToSerial(t));
}

XLCellValue XLFormulaEngine::fnEomonth(const std::vector<XLFormulaArg>& args)
{
    // EOMONTH(start_date, months) – last day of month, N months later
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::tm t      = serialToTm(toDouble(args[0][0]));
    int     months = static_cast<int>(toDouble(args[1][0]));
    t.tm_mon += months;
    t.tm_year += t.tm_mon / 12;
    t.tm_mon = t.tm_mon % 12;
    if (t.tm_mon < 0) {
        t.tm_year--;
        t.tm_mon += 12;
    }
    t.tm_mday = daysInMonth(t.tm_year + 1900, t.tm_mon + 1);
    return XLCellValue(tmToSerial(t));
}

XLCellValue XLFormulaEngine::fnWorkday(const std::vector<XLFormulaArg>& args)
{
    // WORKDAY(start_date, days, [holidays])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double current = std::floor(toDouble(args[0][0]));
    int    days    = static_cast<int>(toDouble(args[1][0]));
    int    step    = (days > 0) ? 1 : -1;

    // Collect holidays (arg[2] if present)
    std::vector<double> holidays;
    if (args.size() > 2) {
        for (const auto& v : args[2])
            if (isNumeric(v)) holidays.push_back(std::floor(toDouble(v)));
    }

    int remaining = std::abs(days);
    while (remaining > 0) {
        current += step;
        if (isWeekend(current)) continue;
        bool isHoliday = std::find(holidays.begin(), holidays.end(), current) != holidays.end();
        if (!isHoliday) --remaining;
    }
    return XLCellValue(current);
}

XLCellValue XLFormulaEngine::fnNetworkdays(const std::vector<XLFormulaArg>& args)
{
    // NETWORKDAYS(start_date, end_date, [holidays])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double start = std::floor(toDouble(args[0][0]));
    double end   = std::floor(toDouble(args[1][0]));

    std::vector<double> holidays;
    if (args.size() > 2)
        for (const auto& v : args[2])
            if (isNumeric(v)) holidays.push_back(std::floor(toDouble(v)));

    int    count = 0;
    int    sign  = (end >= start) ? 1 : -1;
    double d     = start;
    while ((sign > 0 && d <= end) || (sign < 0 && d >= end)) {
        if (!isWeekend(d)) {
            bool isHoliday = std::find(holidays.begin(), holidays.end(), d) != holidays.end();
            if (!isHoliday) ++count;
        }
        d += sign;
    }
    return XLCellValue(static_cast<int64_t>(sign * count));
}

// =============================================================================
// Built-in: Financial
// =============================================================================

XLCellValue XLFormulaEngine::fnPmt(const std::vector<XLFormulaArg>& args)
{
    // PMT(rate, nper, pv, [fv=0], [type=0])
    // Payment = (pv*(1+r)^n*r + fv*r) / ((1+r*type)*((1+r)^n - 1))   when r≠0
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double r    = toDouble(args[0][0]);
    double n    = toDouble(args[1][0]);
    double pv   = toDouble(args[2][0]);
    double fv   = (args.size() > 3 && !args[3].empty()) ? toDouble(args[3][0]) : 0.0;
    int    type = (args.size() > 4 && !args[4].empty()) ? static_cast<int>(toDouble(args[4][0])) : 0;

    if (r == 0.0) return XLCellValue(-(pv + fv) / n);
    double q   = std::pow(1.0 + r, n);
    double pmt = -(pv * q + fv) * r / ((1.0 + r * type) * (q - 1.0));
    return XLCellValue(pmt);
}

XLCellValue XLFormulaEngine::fnFv(const std::vector<XLFormulaArg>& args)
{
    // FV(rate, nper, pmt, [pv=0], [type=0])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double r    = toDouble(args[0][0]);
    double n    = toDouble(args[1][0]);
    double pmt  = toDouble(args[2][0]);
    double pv   = (args.size() > 3 && !args[3].empty()) ? toDouble(args[3][0]) : 0.0;
    int    type = (args.size() > 4 && !args[4].empty()) ? static_cast<int>(toDouble(args[4][0])) : 0;

    if (r == 0.0) return XLCellValue(-(pv + pmt * n));
    double q  = std::pow(1.0 + r, n);
    double fv = -(pv * q + pmt * (1.0 + r * type) * (q - 1.0) / r);
    return XLCellValue(fv);
}

XLCellValue XLFormulaEngine::fnPv(const std::vector<XLFormulaArg>& args)
{
    // PV(rate, nper, pmt, [fv=0], [type=0])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double r    = toDouble(args[0][0]);
    double n    = toDouble(args[1][0]);
    double pmt  = toDouble(args[2][0]);
    double fv   = (args.size() > 3 && !args[3].empty()) ? toDouble(args[3][0]) : 0.0;
    int    type = (args.size() > 4 && !args[4].empty()) ? static_cast<int>(toDouble(args[4][0])) : 0;

    if (r == 0.0) return XLCellValue(-(fv + pmt * n));
    double q  = std::pow(1.0 + r, n);
    double pv = -(fv + pmt * (1.0 + r * type) * (q - 1.0) / r) / q;
    return XLCellValue(pv);
}

XLCellValue XLFormulaEngine::fnNpv(const std::vector<XLFormulaArg>& args)
{
    // NPV(rate, value1, value2, ...) – all values flattened, 1-based periods
    if (args.size() < 2 || args[0].empty()) return errValue();
    double r      = toDouble(args[0][0]);
    double npv    = 0.0;
    int    period = 1;
    for (std::size_t i = 1; i < args.size(); ++i)
        for (const auto& v : args[i])
            if (isNumeric(v)) { npv += toDouble(v) / std::pow(1.0 + r, period++); }
    return XLCellValue(npv);
}

// =============================================================================
// Built-in: Math extended
// =============================================================================

XLCellValue XLFormulaEngine::fnSumproduct(const std::vector<XLFormulaArg>& args)
{
    // SUMPRODUCT(array1, array2, …) – element-wise multiplication then sum
    if (args.empty()) return XLCellValue(0.0);
    std::size_t sz = args[0].size();
    for (const auto& a : args) sz = std::min(sz, a.size());    // use shortest
    double total = 0.0;
    for (std::size_t i = 0; i < sz; ++i) {
        double prod = 1.0;
        for (const auto& a : args) prod *= isNumeric(a[i]) ? toDouble(a[i]) : 0.0;
        total += prod;
    }
    return XLCellValue(total);
}

XLCellValue XLFormulaEngine::fnCeil(const std::vector<XLFormulaArg>& args)
{
    // CEILING(number, significance)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num = toDouble(args[0][0]);
    double sig = toDouble(args[1][0]);
    if (sig == 0.0) return XLCellValue(0.0);
    return XLCellValue(std::ceil(num / sig) * sig);
}

XLCellValue XLFormulaEngine::fnFloor(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num = toDouble(args[0][0]);
    double sig = toDouble(args[1][0]);
    if (sig == 0.0) return XLCellValue(0.0);
    return XLCellValue(std::floor(num / sig) * sig);
}

XLCellValue XLFormulaEngine::fnLog(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double num  = toDouble(args[0][0]);
    double base = (args.size() > 1 && !args[1].empty()) ? toDouble(args[1][0]) : 10.0;
    if (num <= 0.0 || base <= 0.0 || base == 1.0) return errNum();
    return XLCellValue(std::log(num) / std::log(base));
}

XLCellValue XLFormulaEngine::fnLog10(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double num = toDouble(args[0][0]);
    if (num <= 0.0) return errNum();
    return XLCellValue(std::log10(num));
}

XLCellValue XLFormulaEngine::fnExp(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    return XLCellValue(std::exp(toDouble(args[0][0])));
}

XLCellValue XLFormulaEngine::fnSign(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double  v = toDouble(args[0][0]);
    int64_t s = 0;
    if (v > 0.0)
        s = 1;
    else if (v < 0.0)
        s = -1;
    return XLCellValue(s);
}

// =============================================================================
// Built-in: Text extended
// =============================================================================

XLCellValue XLFormulaEngine::fnFind(const std::vector<XLFormulaArg>& args)
{
    // FIND(find_text, within_text, [start_num=1]) – case-sensitive, no wildcards
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::string needle   = toString(args[0][0]);
    std::string haystack = toString(args[1][0]);
    int         startPos = (args.size() > 2 && !args[2].empty()) ? static_cast<int>(toDouble(args[2][0])) - 1 : 0;
    if (startPos < 0 || startPos >= static_cast<int>(haystack.size())) return errValue();
    auto pos = haystack.find(needle, static_cast<std::size_t>(startPos));
    if (pos == std::string::npos) return errValue();
    return XLCellValue(static_cast<int64_t>(pos + 1));
}

XLCellValue XLFormulaEngine::fnSearch(const std::vector<XLFormulaArg>& args)
{
    // SEARCH(find_text, within_text, [start_num=1]) – case-insensitive
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::string needle   = toString(args[0][0]);
    std::string haystack = toString(args[1][0]);
    std::transform(needle.begin(), needle.end(), needle.begin(), ::tolower);
    std::transform(haystack.begin(), haystack.end(), haystack.begin(), ::tolower);
    int startPos = (args.size() > 2 && !args[2].empty()) ? static_cast<int>(toDouble(args[2][0])) - 1 : 0;
    if (startPos < 0 || startPos >= static_cast<int>(haystack.size())) return errValue();
    auto pos = haystack.find(needle, static_cast<std::size_t>(startPos));
    if (pos == std::string::npos) return errValue();
    return XLCellValue(static_cast<int64_t>(pos + 1));
}

XLCellValue XLFormulaEngine::fnSubstitute(const std::vector<XLFormulaArg>& args)
{
    // SUBSTITUTE(text, old_text, new_text, [instance_num])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    std::string s       = toString(args[0][0]);
    std::string oldStr  = toString(args[1][0]);
    std::string newStr  = toString(args[2][0]);
    int         instNum = (args.size() > 3 && !args[3].empty()) ? static_cast<int>(toDouble(args[3][0])) : 0;

    if (oldStr.empty()) return XLCellValue(s);
    std::string result;
    result.reserve(s.size());
    std::size_t pos        = 0;
    int         occurrence = 0;
    while (true) {
        auto found = s.find(oldStr, pos);
        if (found == std::string::npos) {
            result.append(s, pos, std::string::npos);
            break;
        }
        ++occurrence;
        result.append(s, pos, found - pos);
        if (instNum == 0 || occurrence == instNum)
            result += newStr;
        else
            result.append(oldStr);
        pos = found + oldStr.size();
    }
    return XLCellValue(result);
}

XLCellValue XLFormulaEngine::fnReplace(const std::vector<XLFormulaArg>& args)
{
    // REPLACE(old_text, start_num, num_chars, new_text)
    if (args.size() < 4 || args[0].empty() || args[1].empty() || args[2].empty() || args[3].empty()) return errValue();
    std::string s      = toString(args[0][0]);
    int         start  = static_cast<int>(toDouble(args[1][0])) - 1;    // convert to 0-based
    int         nchars = static_cast<int>(toDouble(args[2][0]));
    std::string newStr = toString(args[3][0]);
    if (start < 0) start = 0;
    auto startSz  = static_cast<std::size_t>(start);
    auto ncharsSz = static_cast<std::size_t>(std::max(0, nchars));
    if (startSz > s.size()) startSz = s.size();
    s.replace(startSz, ncharsSz, newStr);
    return XLCellValue(s);
}

XLCellValue XLFormulaEngine::fnRept(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::string s = toString(args[0][0]);
    int         n = static_cast<int>(toDouble(args[1][0]));
    if (n < 0) return errValue();
    std::string result;
    result.reserve(s.size() * static_cast<std::size_t>(n));
    for (int i = 0; i < n; ++i) result += s;
    return XLCellValue(result);
}

XLCellValue XLFormulaEngine::fnExact(const std::vector<XLFormulaArg>& args)
{
    // EXACT(text1, text2) – case-sensitive comparison
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    return XLCellValue(toString(args[0][0]) == toString(args[1][0]));
}

XLCellValue XLFormulaEngine::fnT(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(std::string(""));
    const auto& v = args[0][0];
    return v.type() == XLValueType::String ? v : XLCellValue(std::string(""));
}

XLCellValue XLFormulaEngine::fnValue(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    if (isNumeric(args[0][0])) return XLCellValue(toDouble(args[0][0]));
    std::string s = toString(args[0][0]);
    try {
        std::size_t idx = 0;
        double      v   = std::stod(s, &idx);
        if (idx == s.size()) return XLCellValue(v);
    }
    catch (...) {
    }
    return errValue();
}

XLCellValue XLFormulaEngine::fnTextjoin(const std::vector<XLFormulaArg>& args)
{
    // TEXTJOIN(delimiter, ignore_empty, text1, text2, ...)
    if (args.size() < 3 || args[0].empty() || args[1].empty()) return errValue();
    std::string delim       = toString(args[0][0]);
    bool        ignoreEmpty = static_cast<bool>(toDouble(args[1][0]));
    std::string result;
    bool        first = true;
    for (std::size_t i = 2; i < args.size(); ++i) {
        for (const auto& v : args[i]) {
            std::string s = toString(v);
            if (ignoreEmpty && s.empty()) continue;
            if (!first) result += delim;
            result += s;
            first = false;
        }
    }
    return XLCellValue(result);
}

XLCellValue XLFormulaEngine::fnClean(const std::vector<XLFormulaArg>& args)
{
    // CLEAN – remove non-printable characters (< 0x20)
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    s.erase(std::remove_if(s.begin(), s.end(), [](unsigned char c) { return c < 0x20; }), s.end());
    return XLCellValue(s);
}

XLCellValue XLFormulaEngine::fnProper(const std::vector<XLFormulaArg>& args)
{
    // PROPER – Title Case
    if (args.empty() || args[0].empty()) return errValue();
    std::string s       = toString(args[0][0]);
    bool        capNext = true;
    for (char& c : s) {
        if (std::isspace(static_cast<unsigned char>(c))) { capNext = true; }
        else if (capNext) {
            c       = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
            capNext = false;
        }
        else {
            c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
        }
    }
    return XLCellValue(s);
}

// =============================================================================
// Built-in: Statistical / Conditional
// =============================================================================

XLCellValue XLFormulaEngine::fnSumif(const std::vector<XLFormulaArg>& args)
{
    // SUMIF(range, criteria, [sum_range])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const auto& range    = args[0];
    std::string criteria = toString(args[1][0]);
    const auto& sumRange = (args.size() > 2) ? args[2] : args[0];
    double      total    = 0.0;
    for (std::size_t i = 0; i < range.size(); ++i) {
        if (matchesCriteria(range[i], criteria)) {
            if (i < sumRange.size() && isNumeric(sumRange[i])) total += toDouble(sumRange[i]);
        }
    }
    return XLCellValue(total);
}

XLCellValue XLFormulaEngine::fnCountif(const std::vector<XLFormulaArg>& args)
{
    // COUNTIF(range, criteria)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::string criteria = toString(args[1][0]);
    int64_t     cnt      = 0;
    for (const auto& v : args[0])
        if (matchesCriteria(v, criteria)) ++cnt;
    return XLCellValue(cnt);
}

XLCellValue XLFormulaEngine::fnSumifs(const std::vector<XLFormulaArg>& args)
{
    // SUMIFS(sum_range, crit_range1, crit1, crit_range2, crit2, ...)
    if (args.size() < 3 || args[0].empty()) return errValue();
    const auto& sumRange = args[0];
    double      total    = 0.0;
    for (std::size_t i = 0; i < sumRange.size(); ++i) {
        bool match = true;
        for (std::size_t p = 1; p + 1 < args.size(); p += 2) {
            if (args[p].empty() || args[p + 1].empty()) {
                match = false;
                break;
            }
            if (i >= args[p].size()) {
                match = false;
                break;
            }
            if (!matchesCriteria(args[p][i], toString(args[p + 1][0]))) {
                match = false;
                break;
            }
        }
        if (match && isNumeric(sumRange[i])) total += toDouble(sumRange[i]);
    }
    return XLCellValue(total);
}

XLCellValue XLFormulaEngine::fnCountifs(const std::vector<XLFormulaArg>& args)
{
    // COUNTIFS(crit_range1, crit1, crit_range2, crit2, ...)
    if (args.size() < 2) return errValue();
    std::size_t sz  = args[0].size();
    int64_t     cnt = 0;
    for (std::size_t i = 0; i < sz; ++i) {
        bool match = true;
        for (std::size_t p = 0; p + 1 < args.size(); p += 2) {
            if (args[p + 1].empty()) {
                match = false;
                break;
            }
            if (i >= args[p].size()) {
                match = false;
                break;
            }
            if (!matchesCriteria(args[p][i], toString(args[p + 1][0]))) {
                match = false;
                break;
            }
        }
        if (match) ++cnt;
    }
    return XLCellValue(cnt);
}

XLCellValue XLFormulaEngine::fnMaxifs(const std::vector<XLFormulaArg>& args)
{
    // MAXIFS(max_range, crit_range1, crit1, crit_range2, crit2, ...)
    if (args.size() < 3 || args[0].empty()) return errValue();
    const auto& maxRange = args[0];
    double      maxVal   = -std::numeric_limits<double>::infinity();
    bool        found    = false;
    for (std::size_t i = 0; i < maxRange.size(); ++i) {
        bool match = true;
        for (std::size_t p = 1; p + 1 < args.size(); p += 2) {
            if (args[p].empty() || args[p + 1].empty()) {
                match = false;
                break;
            }
            if (i >= args[p].size()) {
                match = false;
                break;
            }
            if (!matchesCriteria(args[p][i], toString(args[p + 1][0]))) {
                match = false;
                break;
            }
        }
        if (match && isNumeric(maxRange[i])) {
            double v = toDouble(maxRange[i]);
            if (!found || v > maxVal) {
                maxVal = v;
                found  = true;
            }
        }
    }
    if (!found) return XLCellValue(0.0);
    return XLCellValue(maxVal);
}

XLCellValue XLFormulaEngine::fnMinifs(const std::vector<XLFormulaArg>& args)
{
    // MINIFS(min_range, crit_range1, crit1, crit_range2, crit2, ...)
    if (args.size() < 3 || args[0].empty()) return errValue();
    const auto& minRange = args[0];
    double      minVal   = std::numeric_limits<double>::infinity();
    bool        found    = false;
    for (std::size_t i = 0; i < minRange.size(); ++i) {
        bool match = true;
        for (std::size_t p = 1; p + 1 < args.size(); p += 2) {
            if (args[p].empty() || args[p + 1].empty()) {
                match = false;
                break;
            }
            if (i >= args[p].size()) {
                match = false;
                break;
            }
            if (!matchesCriteria(args[p][i], toString(args[p + 1][0]))) {
                match = false;
                break;
            }
        }
        if (match && isNumeric(minRange[i])) {
            double v = toDouble(minRange[i]);
            if (!found || v < minVal) {
                minVal = v;
                found  = true;
            }
        }
    }
    if (!found) return XLCellValue(0.0);
    return XLCellValue(minVal);
}

XLCellValue XLFormulaEngine::fnAverageif(const std::vector<XLFormulaArg>& args)
{
    // AVERAGEIF(range, criteria, [average_range])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const auto& range    = args[0];
    std::string criteria = toString(args[1][0]);
    const auto& avgRange = (args.size() > 2 && !args[2].empty()) ? args[2] : args[0];
    double      total    = 0.0;
    int64_t     cnt      = 0;
    for (std::size_t i = 0; i < range.size(); ++i) {
        if (matchesCriteria(range[i], criteria)) {
            if (i < avgRange.size() && isNumeric(avgRange[i])) {
                total += toDouble(avgRange[i]);
                ++cnt;
            }
        }
    }
    if (cnt == 0) return errDiv0();
    return XLCellValue(total / static_cast<double>(cnt));
}

XLCellValue XLFormulaEngine::fnRank(const std::vector<XLFormulaArg>& args)
{
    // RANK(number, ref, [order=0])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double  num   = toDouble(args[0][0]);
    int     order = (args.size() > 2 && !args[2].empty()) ? static_cast<int>(toDouble(args[2][0])) : 0;
    auto    nums  = numerics(args[1]);
    int64_t rank  = 1;
    for (double v : nums) {
        if ((order == 0 && v > num) || (order != 0 && v < num)) ++rank;
    }
    if (std::find(nums.begin(), nums.end(), num) == nums.end()) return errNA();
    return XLCellValue(rank);
}

XLCellValue XLFormulaEngine::fnLarge(const std::vector<XLFormulaArg>& args)
{
    // LARGE(array, k)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    auto nums = numerics(args[0]);
    auto k    = static_cast<std::ptrdiff_t>(toDouble(args[1][0]));
    if (k < 1 || k > static_cast<std::ptrdiff_t>(nums.size())) return errNum();
    std::nth_element(nums.begin(), nums.begin() + (static_cast<std::ptrdiff_t>(nums.size()) - k), nums.end());
    return XLCellValue(nums[static_cast<std::size_t>(static_cast<std::ptrdiff_t>(nums.size()) - k)]);
}

XLCellValue XLFormulaEngine::fnSmall(const std::vector<XLFormulaArg>& args)
{
    // SMALL(array, k)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    auto nums = numerics(args[0]);
    auto k    = static_cast<std::ptrdiff_t>(toDouble(args[1][0]));
    if (k < 1 || k > static_cast<std::ptrdiff_t>(nums.size())) return errNum();
    std::nth_element(nums.begin(), nums.begin() + (k - 1), nums.end());
    return XLCellValue(nums[static_cast<std::size_t>(k - 1)]);
}

XLCellValue XLFormulaEngine::fnStdev(const std::vector<XLFormulaArg>& args)
{
    // STDEV (sample) – two-pass for numerical stability
    auto nums = numerics(args);
    if (nums.size() < 2) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq   = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(std::sqrt(sq / static_cast<double>(nums.size() - 1)));
}

XLCellValue XLFormulaEngine::fnVar(const std::vector<XLFormulaArg>& args)
{
    // VAR (sample variance)
    auto nums = numerics(args);
    if (nums.size() < 2) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq   = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(sq / static_cast<double>(nums.size() - 1));
}

XLCellValue XLFormulaEngine::fnMedian(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errNum();
    std::sort(nums.begin(), nums.end());
    std::size_t n = nums.size();
    if (n % 2 == 1) return XLCellValue(nums[n / 2]);
    return XLCellValue((nums[n / 2 - 1] + nums[n / 2]) / 2.0);
}

XLCellValue XLFormulaEngine::fnCountblank(const std::vector<XLFormulaArg>& args)
{
    int64_t cnt = 0;
    for (const auto& arg : args) {
        for (const auto& v : arg) {
            if (isEmpty(v) || (v.type() == XLValueType::String && v.get<std::string>().empty())) ++cnt;
        }
    }
    return XLCellValue(cnt);
}

// =============================================================================
// Built-in: Info extended
// =============================================================================

XLCellValue XLFormulaEngine::fnIsna(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    const auto& v = args[0][0];
    if (v.type() != XLValueType::Error) return XLCellValue(false);
    return XLCellValue(v.get<std::string>() == "#N/A");
}

XLCellValue XLFormulaEngine::fnIfna(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const auto& v = args[0][0];
    if (v.type() == XLValueType::Error && v.get<std::string>() == "#N/A") return args[1][0];
    return v;
}

XLCellValue XLFormulaEngine::fnIslogical(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    return XLCellValue(args[0][0].type() == XLValueType::Boolean);
}

// =============================================================================
// Easy Additions (Logical, Math, Stat, Text, Financial)
// =============================================================================

XLCellValue XLFormulaEngine::fnTrue(const std::vector<XLFormulaArg>& /*args*/)
{
    return XLCellValue(true);
}

XLCellValue XLFormulaEngine::fnFalse(const std::vector<XLFormulaArg>& /*args*/)
{
    return XLCellValue(false);
}

XLCellValue XLFormulaEngine::fnIseven(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double val = toDouble(args[0][0]);
    if (std::isnan(val)) return errValue();
    int64_t i = static_cast<int64_t>(std::trunc(val));
    return XLCellValue(i % 2 == 0);
}

XLCellValue XLFormulaEngine::fnIsodd(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double val = toDouble(args[0][0]);
    if (std::isnan(val)) return errValue();
    int64_t i = static_cast<int64_t>(std::trunc(val));
    return XLCellValue(i % 2 != 0);
}

XLCellValue XLFormulaEngine::fnIserr(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    const auto& v = args[0][0];
    if (v.type() != XLValueType::Error) return XLCellValue(false);
    return XLCellValue(v.get<std::string>() != "#N/A");
}

XLCellValue XLFormulaEngine::fnTrunc(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double num = toDouble(args[0][0]);
    if (std::isnan(num)) return errValue();
    int digits = 0;
    if (args.size() > 1 && !args[1].empty()) {
        digits = static_cast<int>(toDouble(args[1][0]));
    }
    double factor = std::pow(10.0, digits);
    return XLCellValue(std::trunc(num * factor) / factor);
}

XLCellValue XLFormulaEngine::fnSumsq(const std::vector<XLFormulaArg>& args)
{
    double total = 0.0;
    for (const auto& v : numerics(args)) {
        total += v * v;
    }
    return XLCellValue(total);
}

static std::vector<double> extractArrayStringFallback(const XLFormulaArg& arg) {
    std::vector<double> res;
    if (arg.empty()) return res;
    const auto& v = arg[0];
    if (v.type() == XLValueType::String) {
        std::string s = v.get<std::string>();
        if (!s.empty() && s.front() == '{' && s.back() == '}') {
            s = s.substr(1, s.length() - 2);
            std::size_t start = 0;
            while (start < s.length()) {
                std::size_t pos = s.find_first_of(",;", start);
                if (pos == std::string::npos) pos = s.length();
                std::string token = s.substr(start, pos - start);
                try { res.push_back(std::stod(token)); } catch (...) {}
                start = pos + 1;
            }
        }
    } else if (isNumeric(v)) {
        res.push_back(toDouble(v));
    }
    return res;
}

XLCellValue XLFormulaEngine::fnSumx2my2(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2) return errValue();
    auto numsX = numerics(args[0]);
    auto numsY = numerics(args[1]);
    if (numsX.empty()) numsX = extractArrayStringFallback(args[0]);
    if (numsY.empty()) numsY = extractArrayStringFallback(args[1]);

    if (numsX.size() != numsY.size()) return errNA();
    std::size_t sz = numsX.size();
    if (sz == 0) return XLCellValue(0.0);
    double total = 0.0;
    for (std::size_t i = 0; i < sz; ++i) {
        double x = numsX[i];
        double y = numsY[i];
        total += (x * x) - (y * y);
    }
    return XLCellValue(total);
}

XLCellValue XLFormulaEngine::fnSumx2py2(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2) return errValue();
    auto numsX = numerics(args[0]);
    auto numsY = numerics(args[1]);
    if (numsX.empty()) numsX = extractArrayStringFallback(args[0]);
    if (numsY.empty()) numsY = extractArrayStringFallback(args[1]);
    if (numsX.size() != numsY.size()) return errNA();
    std::size_t sz = numsX.size();
    if (sz == 0) return XLCellValue(0.0);
    double total = 0.0;
    for (std::size_t i = 0; i < sz; ++i) {
        double x = numsX[i];
        double y = numsY[i];
        total += (x * x) + (y * y);
    }
    return XLCellValue(total);
}

XLCellValue XLFormulaEngine::fnSumxmy2(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2) return errValue();
    auto numsX = numerics(args[0]);
    auto numsY = numerics(args[1]);
    if (numsX.empty()) numsX = extractArrayStringFallback(args[0]);
    if (numsY.empty()) numsY = extractArrayStringFallback(args[1]);
    if (numsX.size() != numsY.size()) return errNA();
    std::size_t sz = numsX.size();
    if (sz == 0) return XLCellValue(0.0);
    double total = 0.0;
    for (std::size_t i = 0; i < sz; ++i) {
        double x = numsX[i];
        double y = numsY[i];
        double diff = x - y;
        total += diff * diff;
    }
    return XLCellValue(total);
}

XLCellValue XLFormulaEngine::fnAvedev(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errDiv0();
    double total = 0.0;
    for (double v : nums) total += v;
    double mean = total / static_cast<double>(nums.size());
    
    double devSum = 0.0;
    for (double v : nums) devSum += std::abs(v - mean);
    return XLCellValue(devSum / static_cast<double>(nums.size()));
}

XLCellValue XLFormulaEngine::fnDevsq(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errDiv0();
    double total = 0.0;
    for (double v : nums) total += v;
    double mean = total / static_cast<double>(nums.size());
    
    double sqSum = 0.0;
    for (double v : nums) {
        double diff = v - mean;
        sqSum += diff * diff;
    }
    return XLCellValue(sqSum);
}

XLCellValue XLFormulaEngine::fnAveragea(const std::vector<XLFormulaArg>& args)
{
    double total = 0.0;
    int count = 0;
    for (const auto& arg : args) {
        bool isScalar = (arg.type() == XLFormulaArg::Type::Scalar);
        for (const auto& v : arg) {
            if (v.type() == XLValueType::Empty) continue;
            count++;
            if (v.type() == XLValueType::Integer || v.type() == XLValueType::Float) {
                total += toDouble(v);
            } else if (v.type() == XLValueType::Boolean) {
                total += (v.get<bool>() ? 1.0 : 0.0);
            } else if (v.type() == XLValueType::String) {
                if (isScalar) {
                    try {
                        total += std::stod(v.get<std::string>());
                    } catch (...) {
                        // keep as 0
                    }
                }
            }
            // Texts in ranges are counted but added as 0.0
        }
    }
    if (count == 0) return errDiv0();
    return XLCellValue(total / static_cast<double>(count));
}

XLCellValue XLFormulaEngine::fnSln(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double cost = toDouble(args[0][0]);
    double salvage = toDouble(args[1][0]);
    double life = toDouble(args[2][0]);
    if (std::isnan(cost) || std::isnan(salvage) || std::isnan(life) || life == 0.0) return errDiv0();
    return XLCellValue((cost - salvage) / life);
}

XLCellValue XLFormulaEngine::fnSyd(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 4 || args[0].empty() || args[1].empty() || args[2].empty() || args[3].empty()) return errValue();
    double cost = toDouble(args[0][0]);
    double salvage = toDouble(args[1][0]);
    double life = toDouble(args[2][0]);
    double per = toDouble(args[3][0]);
    if (std::isnan(cost) || std::isnan(salvage) || std::isnan(life) || std::isnan(per) || life == 0.0) return errDiv0();
    if (per < 1.0 || per > life) return errNum();
    return XLCellValue(((cost - salvage) * (life - per + 1.0) * 2.0) / (life * (life + 1.0)));
}

XLCellValue XLFormulaEngine::fnChar(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double dval = toDouble(args[0][0]);
    if (std::isnan(dval)) return errValue();
    int code = static_cast<int>(dval);
    if (code < 1 || code > 255) return errValue();
    std::string s(1, static_cast<char>(code));
    return XLCellValue(s);
}

XLCellValue XLFormulaEngine::fnUnichar(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double dval = toDouble(args[0][0]);
    if (std::isnan(dval)) return errValue();
    int code = static_cast<int>(dval);
    if (code <= 0 || code > 0x10FFFF) return errValue(); // Valid Unicode range
    
    std::string s;
    if (code <= 0x7F) {
        s.push_back(static_cast<char>(code));
    } else if (code <= 0x7FF) {
        s.push_back(static_cast<char>(0xC0 | (code >> 6)));
        s.push_back(static_cast<char>(0x80 | (code & 0x3F)));
    } else if (code <= 0xFFFF) {
        s.push_back(static_cast<char>(0xE0 | (code >> 12)));
        s.push_back(static_cast<char>(0x80 | ((code >> 6) & 0x3F)));
        s.push_back(static_cast<char>(0x80 | (code & 0x3F)));
    } else {
        s.push_back(static_cast<char>(0xF0 | (code >> 18)));
        s.push_back(static_cast<char>(0x80 | ((code >> 12) & 0x3F)));
        s.push_back(static_cast<char>(0x80 | ((code >> 6) & 0x3F)));
        s.push_back(static_cast<char>(0x80 | (code & 0x3F)));
    }
    return XLCellValue(s);
}

XLCellValue XLFormulaEngine::fnCode(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    if (s.empty()) return errValue();
    // Excel's CODE function typically just looks at the first single byte or system charset
    // Here we maintain basic behavior for first byte
    return XLCellValue(static_cast<int64_t>(static_cast<unsigned char>(s[0])));
}

XLCellValue XLFormulaEngine::fnUnicode(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    if (s.empty()) return errValue();
    
    const unsigned char* str = reinterpret_cast<const unsigned char*>(s.c_str());
    int cp = 0;
    if (str[0] < 0x80) {
        cp = str[0];
    } else if ((str[0] & 0xE0) == 0xC0) {
        if (s.length() < 2) return errValue();
        cp = ((str[0] & 0x1F) << 6) | (str[1] & 0x3F);
    } else if ((str[0] & 0xF0) == 0xE0) {
        if (s.length() < 3) return errValue();
        cp = ((str[0] & 0x0F) << 12) | ((str[1] & 0x3F) << 6) | (str[2] & 0x3F);
    } else if ((str[0] & 0xF8) == 0xF0) {
        if (s.length() < 4) return errValue();
        cp = ((str[0] & 0x07) << 18) | ((str[1] & 0x3F) << 12) | ((str[2] & 0x3F) << 6) | (str[3] & 0x3F);
    } else {
        return errValue(); // Invalid UTF-8
    }
    
    return XLCellValue(static_cast<int64_t>(cp));
}


// =============================================================================
// Easy Additions (Math & Trig)
// =============================================================================

XLCellValue XLFormulaEngine::fnMround(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 2) return errValue();
    if (nums[1] == 0) return XLCellValue(0.0);
    if ((nums[0] * nums[1]) < 0) return errNum();
    double res = std::round(nums[0] / nums[1]) * nums[1];
    return XLCellValue(res);
}

XLCellValue XLFormulaEngine::fnCeilingMath(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty() || nums.size() > 3) return errValue();
    double number = nums[0];
    double significance = (nums.size() > 1) ? nums[1] : (number >= 0 ? 1.0 : -1.0);
    if (significance == 0) return XLCellValue(0.0);
    double mode = (nums.size() > 2) ? nums[2] : 0.0;
    
    double res;
    if (number >= 0) {
        res = std::ceil(number / std::abs(significance)) * std::abs(significance);
    } else {
        if (mode == 0) {
            res = std::ceil(number / std::abs(significance)) * std::abs(significance);
        } else {
            res = std::floor(number / std::abs(significance)) * std::abs(significance);
        }
    }
    return XLCellValue(res);
}

XLCellValue XLFormulaEngine::fnFloorMath(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty() || nums.size() > 3) return errValue();
    double number = nums[0];
    double significance = (nums.size() > 1) ? nums[1] : (number >= 0 ? 1.0 : -1.0);
    if (significance == 0) return XLCellValue(0.0);
    double mode = (nums.size() > 2) ? nums[2] : 0.0;
    
    double res;
    if (number >= 0) {
        res = std::floor(number / std::abs(significance)) * std::abs(significance);
    } else {
        if (mode == 0) {
            res = std::floor(number / std::abs(significance)) * std::abs(significance);
        } else {
            res = std::ceil(number / std::abs(significance)) * std::abs(significance);
        }
    }
    return XLCellValue(res);
}

// =============================================================================
// Easy Additions (Statistical)
// =============================================================================

XLCellValue XLFormulaEngine::fnVarp(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(sq / static_cast<double>(nums.size()));
}

XLCellValue XLFormulaEngine::fnStdevp(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(std::sqrt(sq / static_cast<double>(nums.size())));
}

static std::vector<double> extractArrayValuesA(const std::vector<XLFormulaArg>& args) {
    std::vector<double> out;
    for (const auto& arg : args) {
        for (const auto& v : arg) {
            if (v.type() == XLValueType::Integer) {
                out.push_back(static_cast<double>(v.get<int64_t>()));
            } else if (v.type() == XLValueType::Float) {
                out.push_back(v.get<double>());
            } else if (v.type() == XLValueType::Boolean) {
                out.push_back(v.get<bool>() ? 1.0 : 0.0);
            } else if (v.type() == XLValueType::String) {
                out.push_back(0.0);
            }
        }
    }
    return out;
}

XLCellValue XLFormulaEngine::fnVara(const std::vector<XLFormulaArg>& args)
{
    auto nums = extractArrayValuesA(args);
    if (nums.size() < 2) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(sq / static_cast<double>(nums.size() - 1));
}

XLCellValue XLFormulaEngine::fnVarpa(const std::vector<XLFormulaArg>& args)
{
    auto nums = extractArrayValuesA(args);
    if (nums.empty()) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(sq / static_cast<double>(nums.size()));
}

XLCellValue XLFormulaEngine::fnStdeva(const std::vector<XLFormulaArg>& args)
{
    auto nums = extractArrayValuesA(args);
    if (nums.size() < 2) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(std::sqrt(sq / static_cast<double>(nums.size() - 1)));
}

XLCellValue XLFormulaEngine::fnStdevpa(const std::vector<XLFormulaArg>& args)
{
    auto nums = extractArrayValuesA(args);
    if (nums.empty()) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(std::sqrt(sq / static_cast<double>(nums.size())));
}

XLCellValue XLFormulaEngine::fnPermut(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 2) return errValue();
    double n = std::trunc(nums[0]);
    double k = std::trunc(nums[1]);
    if (n < 0 || k < 0 || n < k) return errNum();
    double res = 1.0;
    for (double i = n; i > n - k; --i) {
        res *= i;
    }
    return XLCellValue(res);
}

XLCellValue XLFormulaEngine::fnPermutationa(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 2) return errValue();
    double n = std::trunc(nums[0]);
    double k = std::trunc(nums[1]);
    if (n < 0 || k < 0) return errNum();
    return XLCellValue(std::pow(n, k));
}

XLCellValue XLFormulaEngine::fnFisher(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 1) return errValue();
    double x = nums[0];
    if (x <= -1 || x >= 1) return errNum();
    return XLCellValue(0.5 * std::log((1 + x) / (1 - x)));
}

XLCellValue XLFormulaEngine::fnFisherinv(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 1) return errValue();
    double y = nums[0];
    return XLCellValue((std::exp(2 * y) - 1) / (std::exp(2 * y) + 1));
}

XLCellValue XLFormulaEngine::fnStandardize(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 3) return errValue();
    double x = nums[0];
    double mean = nums[1];
    double stdev = nums[2];
    if (stdev <= 0) return errNum();
    return XLCellValue((x - mean) / stdev);
}

XLCellValue XLFormulaEngine::fnPearson(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums1 = numerics(args[0]);
    auto nums2 = numerics(args[1]);
    if (nums1.size() != nums2.size() || nums1.empty()) return errNA();
    
    double mean1 = std::accumulate(nums1.begin(), nums1.end(), 0.0) / nums1.size();
    double mean2 = std::accumulate(nums2.begin(), nums2.end(), 0.0) / nums2.size();
    
    double num = 0.0, den1 = 0.0, den2 = 0.0;
    for (size_t i = 0; i < nums1.size(); ++i) {
        double d1 = nums1[i] - mean1;
        double d2 = nums2[i] - mean2;
        num += d1 * d2;
        den1 += d1 * d1;
        den2 += d2 * d2;
    }
    if (den1 == 0 || den2 == 0) return errDiv0();
    return XLCellValue(num / std::sqrt(den1 * den2));
}

XLCellValue XLFormulaEngine::fnRsq(const std::vector<XLFormulaArg>& args)
{
    XLCellValue pearson = fnPearson(args);
    if (pearson.type() == XLValueType::Error) return pearson;
    double p = pearson.get<double>();
    return XLCellValue(p * p);
}

XLCellValue XLFormulaEngine::fnAverageifs(const std::vector<XLFormulaArg>& args)
{
    // Needs proper SUMIFS / COUNTIFS logic.
    // Excel AVERAGEIFS: AVERAGEIFS(average_range, criteria_range1, criteria1, ...)
    // If we have fnSumifs and fnCountifs, we can reuse them.
    // Wait, fnSumifs is expecting sum_range first. So is fnAverageifs.
    // Let's just call them.
    XLCellValue sum = fnSumifs(args);
    if (sum.type() == XLValueType::Error) return sum;
    
    // For COUNTIFS, we skip the average_range (args[0]) and pass the rest.
    std::vector<XLFormulaArg> countArgs;
    for (size_t i = 1; i < args.size(); ++i) {
        countArgs.push_back(args[i]);
    }
    XLCellValue count = fnCountifs(countArgs);
    if (count.type() == XLValueType::Error) return count;
    
    double dCount = count.get<double>();
    if (dCount == 0) return errDiv0();
    
    return XLCellValue(sum.get<double>() / dCount);
}

// =============================================================================
// Easy Additions (Date & Time)
// =============================================================================

XLCellValue XLFormulaEngine::fnIsoweeknum(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errValue();
    double serial = nums[0];
    if (serial < 0) return errNum();
    
    auto dayOfWeek = [](double s) {
        const auto day = static_cast<int64_t>(std::floor(s)) % 7;
        return (day == 0 ? 6 : static_cast<int>(day - 1));
    };

    int wday_mon = (dayOfWeek(serial) + 6) % 7; // 0=Mon .. 6=Sun
    double nearest_thu = std::floor(serial) + 3 - wday_mon;
    
    auto tm_thu = static_cast<std::tm>(XLDateTime(nearest_thu));
    
    std::tm tm_jan4{};
    tm_jan4.tm_year = tm_thu.tm_year;
    tm_jan4.tm_mon = 0;
    tm_jan4.tm_mday = 4;
    tm_jan4.tm_isdst = -1;
    double jan4_serial = std::floor(XLDateTime(tm_jan4).serial());
    
    int jan4_wday_mon = (dayOfWeek(jan4_serial) + 6) % 7;
    double first_thu = jan4_serial + 3 - jan4_wday_mon;
    
    int iso_week = 1 + static_cast<int>(std::round((nearest_thu - first_thu) / 7.0));
    return XLCellValue(static_cast<double>(iso_week));
}

XLCellValue XLFormulaEngine::fnWeeknum(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errValue();
    double serial = nums[0];
    if (serial < 0) return errNum();
    int return_type = (nums.size() > 1) ? static_cast<int>(std::trunc(nums[1])) : 1;
    
    if (return_type == 21) {
        return fnIsoweeknum(args); // System 2 (ISO)
    }
    
    auto dayOfWeek = [](double s) {
        const auto day = static_cast<int64_t>(std::floor(s)) % 7;
        return (day == 0 ? 6 : static_cast<int>(day - 1));
    };

    auto tm = static_cast<std::tm>(XLDateTime(serial));
    std::tm tm_jan1{};
    tm_jan1.tm_year = tm.tm_year;
    tm_jan1.tm_mon = 0;
    tm_jan1.tm_mday = 1;
    tm_jan1.tm_isdst = -1;
    double jan1_serial = std::floor(XLDateTime(tm_jan1).serial());
    
    int days_elapsed = static_cast<int>(std::floor(serial) - jan1_serial);
    
    if (return_type == 1 || return_type == 17) { // Week begins on Sunday
        int jan1_wday = dayOfWeek(jan1_serial); // 0=Sun..6=Sat
        int days_in_wk1 = 7 - jan1_wday;
        if (days_elapsed < days_in_wk1) return XLCellValue(1.0);
        return XLCellValue(2.0 + std::floor((days_elapsed - days_in_wk1) / 7.0));
    } else { // Week begins on Monday
        int jan1_wday_mon = (dayOfWeek(jan1_serial) + 6) % 7; // 0=Mon..6=Sun
        int days_in_wk1 = 7 - jan1_wday_mon;
        if (days_elapsed < days_in_wk1) return XLCellValue(1.0);
        return XLCellValue(2.0 + std::floor((days_elapsed - days_in_wk1) / 7.0));
    }
}

XLCellValue XLFormulaEngine::fnDays360(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() < 2) return errValue();
    double d1 = nums[0];
    double d2 = nums[1];
    bool method = (nums.size() > 2) && nums[2] != 0.0;
    
    auto tm1 = static_cast<std::tm>(XLDateTime(d1));
    auto tm2 = static_cast<std::tm>(XLDateTime(d2));
    
    int y1 = tm1.tm_year + 1900, m1 = tm1.tm_mon + 1, day1 = tm1.tm_mday;
    int y2 = tm2.tm_year + 1900, m2 = tm2.tm_mon + 1, day2 = tm2.tm_mday;
    
    if (method) { // European
        if (day1 == 31) day1 = 30;
        if (day2 == 31) day2 = 30;
    } else { // US (NASD)
        if (m1 == 2 && day1 >= 28) day1 = 30; // approx for end of Feb
        if (day1 == 31) day1 = 30;
        if (day2 == 31 && day1 >= 30) day2 = 30;
    }
    
    double res = (y2 - y1) * 360.0 + (m2 - m1) * 30.0 + (day2 - day1);
    return XLCellValue(res);
}

// =============================================================================
// Easy Additions (Financial)
// =============================================================================

XLCellValue XLFormulaEngine::fnNper(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() < 3) return errValue();
    double rate = nums[0];
    double pmt = nums[1];
    double pv = nums[2];
    double fv = (nums.size() > 3) ? nums[3] : 0.0;
    double type = (nums.size() > 4) ? nums[4] : 0.0;
    
    if (rate == 0) {
        if (pmt == 0) return errNum();
        return XLCellValue(-(pv + fv) / pmt);
    }
    
    double num = pmt * (1 + rate * type) - fv * rate;
    double den = pv * rate + pmt * (1 + rate * type);
    
    double ratio = num / den;
    if (ratio <= 0.0) return errNum();
    
    double res = std::log(ratio) / std::log(1 + rate);
    return XLCellValue(res);
}

XLCellValue XLFormulaEngine::fnDb(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() < 4) return errValue();
    double cost = nums[0];
    double salvage = nums[1];
    double life = nums[2];
    double period = nums[3];
    double month = (nums.size() > 4) ? nums[4] : 12.0;
    
    if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || month <= 0) return errNum();
    if (period > life && (period != life + 1 || month == 12)) return errNum();
    
    double rate = 1.0 - std::pow(salvage / cost, 1.0 / life);
    rate = std::round(rate * 1000.0) / 1000.0; // DB rounds to 3 decimal places
    
    double total = 0.0;
    double dep = cost * rate * month / 12.0;
    total += dep;
    if (period == 1) return XLCellValue(dep);
    
    for (int i = 2; i < period; ++i) {
        dep = (cost - total) * rate;
        total += dep;
    }
    
    if (period == life + 1) {
        dep = (cost - total) * rate * (12.0 - month) / 12.0;
    } else {
        dep = (cost - total) * rate;
    }
    return XLCellValue(dep);
}

XLCellValue XLFormulaEngine::fnDdb(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() < 4) return errValue();
    double cost = nums[0];
    double salvage = nums[1];
    double life = nums[2];
    double period = nums[3];
    double factor = (nums.size() > 4) ? nums[4] : 2.0;
    
    if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || factor <= 0) return errNum();
    
    double total = 0.0;
    double dep = 0.0;
    for (int i = 1; i <= period; ++i) {
        dep = (cost - total) * (factor / life);
        if (total + dep > cost - salvage) {
            dep = cost - salvage - total;
        }
        total += dep;
    }
    return XLCellValue(dep);
}

XLCellValue XLFormulaEngine::fnCovarianceP(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums1 = numerics(args[0]);
    auto nums2 = numerics(args[1]);
    if (nums1.size() != nums2.size() || nums1.empty()) return errNA();
    
    double mean1 = std::accumulate(nums1.begin(), nums1.end(), 0.0) / nums1.size();
    double mean2 = std::accumulate(nums2.begin(), nums2.end(), 0.0) / nums2.size();
    
    double cov = 0.0;
    for (size_t i = 0; i < nums1.size(); ++i) {
        cov += (nums1[i] - mean1) * (nums2[i] - mean2);
    }
    return XLCellValue(cov / nums1.size());
}

XLCellValue XLFormulaEngine::fnCovarianceS(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums1 = numerics(args[0]);
    auto nums2 = numerics(args[1]);
    if (nums1.size() != nums2.size() || nums1.size() < 2) return errDiv0();
    
    double mean1 = std::accumulate(nums1.begin(), nums1.end(), 0.0) / nums1.size();
    double mean2 = std::accumulate(nums2.begin(), nums2.end(), 0.0) / nums2.size();
    
    double cov = 0.0;
    for (size_t i = 0; i < nums1.size(); ++i) {
        cov += (nums1[i] - mean1) * (nums2[i] - mean2);
    }
    return XLCellValue(cov / (nums1.size() - 1));
}

XLCellValue XLFormulaEngine::fnSlope(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums1 = numerics(args[0]); // y
    auto nums2 = numerics(args[1]); // x
    if (nums1.size() != nums2.size() || nums1.empty()) return errNA();
    
    double meanY = std::accumulate(nums1.begin(), nums1.end(), 0.0) / nums1.size();
    double meanX = std::accumulate(nums2.begin(), nums2.end(), 0.0) / nums2.size();
    
    double num = 0.0, den = 0.0;
    for (size_t i = 0; i < nums1.size(); ++i) {
        num += (nums2[i] - meanX) * (nums1[i] - meanY);
        den += (nums2[i] - meanX) * (nums2[i] - meanX);
    }
    if (den == 0.0) return errDiv0();
    return XLCellValue(num / den);
}

XLCellValue XLFormulaEngine::fnIntercept(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums1 = numerics(args[0]); // y
    auto nums2 = numerics(args[1]); // x
    if (nums1.size() != nums2.size() || nums1.empty()) return errNA();
    
    double meanY = std::accumulate(nums1.begin(), nums1.end(), 0.0) / nums1.size();
    double meanX = std::accumulate(nums2.begin(), nums2.end(), 0.0) / nums2.size();
    
    double num = 0.0, den = 0.0;
    for (size_t i = 0; i < nums1.size(); ++i) {
        num += (nums2[i] - meanX) * (nums1[i] - meanY);
        den += (nums2[i] - meanX) * (nums2[i] - meanX);
    }
    if (den == 0.0) return errDiv0();
    double slope = num / den;
    return XLCellValue(meanY - slope * meanX);
}

XLCellValue XLFormulaEngine::fnTrimmean(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums = numerics(args[0]);
    if (nums.empty() || args[1].empty()) return errNum();
    
    double p = toDouble(args[1][0]);
    if (p < 0.0 || p >= 1.0) return errNum();
    
    size_t k = static_cast<size_t>(std::floor(nums.size() * p / 2.0));
    if (nums.size() <= 2 * k) return errNum();
    
    std::sort(nums.begin(), nums.end());
    double sum = 0.0;
    for (size_t i = k; i < nums.size() - k; ++i) {
        sum += nums[i];
    }
    return XLCellValue(sum / (nums.size() - 2 * k));
}

XLCellValue XLFormulaEngine::fnPercentileInc(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums = numerics(args[0]);
    if (nums.empty() || args[1].empty()) return errNum();
    
    double k = toDouble(args[1][0]);
    if (k < 0.0 || k > 1.0) return errNum();
    
    std::sort(nums.begin(), nums.end());
    double pos = k * (nums.size() - 1);
    size_t index = static_cast<size_t>(pos);
    double frac = pos - index;
    
    if (index + 1 < nums.size()) {
        return XLCellValue(nums[index] + frac * (nums[index + 1] - nums[index]));
    } else {
        return XLCellValue(nums[index]);
    }
}

XLCellValue XLFormulaEngine::fnPercentileExc(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums = numerics(args[0]);
    if (nums.empty() || args[1].empty()) return errNum();
    
    double k = toDouble(args[1][0]);
    if (k <= 0.0 || k >= 1.0) return errNum();
    
    std::sort(nums.begin(), nums.end());
    double pos = k * (nums.size() + 1) - 1.0;
    if (pos < 0.0 || pos >= nums.size()) return errNum();
    
    size_t index = static_cast<size_t>(pos);
    double frac = pos - index;
    
    if (index + 1 < nums.size()) {
        return XLCellValue(nums[index] + frac * (nums[index + 1] - nums[index]));
    } else {
        return XLCellValue(nums[index]);
    }
}

XLCellValue XLFormulaEngine::fnQuartileInc(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    if (args[1].empty()) return errNum();
    int quart = static_cast<int>(std::floor(toDouble(args[1][0])));
    if (quart < 0 || quart > 4) return errNum();
    
    std::vector<XLFormulaArg> newArgs = args;
    newArgs[1] = { XLCellValue(quart / 4.0) };
    return fnPercentileInc(newArgs);
}

XLCellValue XLFormulaEngine::fnQuartileExc(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    if (args[1].empty()) return errNum();
    int quart = static_cast<int>(std::floor(toDouble(args[1][0])));
    if (quart < 1 || quart > 3) return errNum();
    
    std::vector<XLFormulaArg> newArgs = args;
    newArgs[1] = { XLCellValue(quart / 4.0) };
    return fnPercentileExc(newArgs);
}

XLCellValue XLFormulaEngine::fnIsnontext(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(true); // blank is non-text
    const auto& v = args[0][0];
    return XLCellValue(v.type() != XLValueType::String);
}



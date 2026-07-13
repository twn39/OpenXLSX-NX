#include "XLFormulaEngine.hpp"
#include "XLException.hpp"
#include <algorithm>
#include <cctype>
#include <fast_float/fast_float.h>

namespace OpenXLSX {

std::vector<XLToken> XLFormulaLexer::tokenize(std::string_view formula)
{
    std::vector<XLToken> tokens;

    // Strip leading '='
    std::size_t i = 0;
    if (!formula.empty() && formula[i] == '=') ++i;

    const auto len = formula.size();

    auto emit = [&](XLTokenKind k, std::size_t startOffset, std::string text = {}, double num = 0.0, bool b = false) {
        XLToken t;
        t.kind    = k;
        t.text    = std::move(text);
        t.number  = num;
        t.boolean = b;
        t.offset  = startOffset;
        tokens.push_back(std::move(t));
    };

    while (i < len) {
        char c = formula[i];

        // --- Whitespace ---
        if (std::isspace(static_cast<unsigned char>(c))) {
            ++i;
            continue;
        }

        std::size_t startOffset = i;

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
            emit(XLTokenKind::String, startOffset, s);
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
            emit(XLTokenKind::Number, startOffset, numStr, val);
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
                if (!second.empty()) { emit(XLTokenKind::CellRef, startOffset, ident + ":" + second); }
                else {
                    emit(XLTokenKind::CellRef, startOffset, ident);
                    emit(XLTokenKind::Colon, i - 1);
                }
                continue;
            }

            // Boolean?
            std::string upper = ident;
            std::transform(upper.begin(), upper.end(), upper.begin(), ::toupper);
            if (upper == "TRUE") {
                emit(XLTokenKind::Bool, startOffset, ident, 1.0, true);
                continue;
            }
            if (upper == "FALSE") {
                emit(XLTokenKind::Bool, startOffset, ident, 0.0, false);
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
                    emit(XLTokenKind::CellRef, startOffset, ident);
                    continue;
                }
                // Sheet-qualified ref like "Sheet1!A1"
                auto bang = ident.find('!');
                if (bang != std::string::npos) {
                    emit(XLTokenKind::CellRef, startOffset, ident);
                    continue;
                }
            }

            emit(XLTokenKind::Ident, startOffset, ident);
            continue;
        }

        // --- Single or double char operators ---
        switch (c) {
            case '+':
                emit(XLTokenKind::Plus, startOffset);
                ++i;
                break;
            case '-':
                emit(XLTokenKind::Minus, startOffset);
                ++i;
                break;
            case '*':
                emit(XLTokenKind::Star, startOffset);
                ++i;
                break;
            case '/':
                emit(XLTokenKind::Slash, startOffset);
                ++i;
                break;
            case '^':
                emit(XLTokenKind::Caret, startOffset);
                ++i;
                break;
            case '%':
                emit(XLTokenKind::Percent, startOffset);
                ++i;
                break;
            case '&':
                emit(XLTokenKind::Amp, startOffset);
                ++i;
                break;
            case '(':
                emit(XLTokenKind::LParen, startOffset);
                ++i;
                break;
            case ')':
                emit(XLTokenKind::RParen, startOffset);
                ++i;
                break;
            case ',':
                emit(XLTokenKind::Comma, startOffset);
                ++i;
                break;
            case ';':
                emit(XLTokenKind::Semicolon, startOffset);
                ++i;
                break;
            case ':':
                emit(XLTokenKind::Colon, startOffset);
                ++i;
                break;
            case '{':
                emit(XLTokenKind::LBrace, startOffset);
                ++i;
                break;
            case '}':
                emit(XLTokenKind::RBrace, startOffset);
                ++i;
                break;
            case '=':
                emit(XLTokenKind::Eq, startOffset);
                ++i;
                break;
            case '<':
                ++i;
                if (i < len && formula[i] == '=') {
                    emit(XLTokenKind::Le, startOffset);
                    ++i;
                }
                else if (i < len && formula[i] == '>') {
                    emit(XLTokenKind::NEq, startOffset);
                    ++i;
                }
                else
                    emit(XLTokenKind::Lt, startOffset);
                break;
            case '>':
                ++i;
                if (i < len && formula[i] == '=') {
                    emit(XLTokenKind::Ge, startOffset);
                    ++i;
                }
                else
                    emit(XLTokenKind::Gt, startOffset);
                break;
            default:
                emit(XLTokenKind::Error, startOffset, std::string(1, c));
                ++i;
                break;
        }
    }

    emit(XLTokenKind::End, i);
    return tokens;
}

// =============================================================================
// Diagnostics
// =============================================================================

void XLFormulaDiagnosticReporter::reportError(std::string message, std::size_t offset)
{
    m_diagnostics.push_back({std::move(message), offset});
}

const std::vector<XLFormulaDiagnostic>& XLFormulaDiagnosticReporter::diagnostics() const
{
    return m_diagnostics;
}

bool XLFormulaDiagnosticReporter::hasErrors() const
{
    return !m_diagnostics.empty();
}

void XLFormulaDiagnosticReporter::clear()
{
    m_diagnostics.clear();
}

std::string XLFormulaDiagnosticReporter::getFullReport() const
{
    std::string report;
    for (const auto& diag : m_diagnostics) {
        if (!report.empty()) report += "\n";
        report += "Error at position " + std::to_string(diag.offset) + ": " + diag.message;
        if (!m_formula.empty()) {
            report += "\n" + m_formula + "\n";
            std::size_t caretPos = std::min(diag.offset, m_formula.size());
            report += std::string(caretPos, ' ') + "^";
        }
    }
    return report;
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

void XLFormulaParser::ParseContext::reportError(std::string message, std::size_t offset)
{
    if (panicMode) return;
    panicMode = true;
    if (reporter) {
        reporter->reportError(std::move(message), offset);
    }
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

std::unique_ptr<XLASTNode> XLFormulaParser::parse(gsl::span<const XLToken> tokens, XLFormulaDiagnosticReporter* reporter)
{
    ParseContext ctx{tokens, 0, reporter, false};
    auto         node = parseExpr(ctx, 0);
    if (ctx.current().kind != XLTokenKind::End) {
        ctx.reportError("Unexpected trailing token '" + ctx.current().text + "'", ctx.current().offset);
    }
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
        if (ctx.current().kind == XLTokenKind::RParen) {
            ctx.consume();
            ctx.panicMode = false;
        } else {
            ctx.reportError("Missing closing parenthesis ')'", ctx.current().offset);
        }
        return inner;
    }

    // Excel array constant: {1,2;3,4}  (comma = column, semicolon = row)
    if (tok.kind == XLTokenKind::LBrace) {
        ctx.consume();
        auto node = std::make_unique<XLASTNode>(XLNodeKind::ArrayLit);
        std::size_t rows = 0;
        std::size_t cols = 0;
        std::size_t curCols = 0;

        auto finishRow = [&]() {
            if (curCols == 0) return;
            if (cols == 0) cols = curCols;
            else if (curCols != cols) {
                ctx.reportError("Array constant rows must have the same length", ctx.current().offset);
            }
            ++rows;
            curCols = 0;
        };

        if (ctx.current().kind != XLTokenKind::RBrace && ctx.current().kind != XLTokenKind::End) {
            while (ctx.current().kind != XLTokenKind::RBrace && ctx.current().kind != XLTokenKind::End) {
                // Array constants: numbers, strings, bools, unary +/- (via parseUnary)
                node->children.push_back(parseUnary(ctx));
                ++curCols;
                if (ctx.current().kind == XLTokenKind::Comma) {
                    ctx.consume();
                    continue;
                }
                if (ctx.current().kind == XLTokenKind::Semicolon) {
                    ctx.consume();
                    finishRow();
                    continue;
                }
                break;
            }
            finishRow();
        }
        if (ctx.current().kind == XLTokenKind::RBrace) {
            ctx.consume();
            ctx.panicMode = false;
        }
        else {
            ctx.reportError("Missing closing '}' for array constant", ctx.current().offset);
        }
        if (rows == 0) {
            rows = 1;
            cols = 0;
        }
        node->number = static_cast<double>(rows);
        node->text   = std::to_string(cols);
        return node;
    }

    if (tok.kind == XLTokenKind::Comma || tok.kind == XLTokenKind::Semicolon || tok.kind == XLTokenKind::RParen ||
        tok.kind == XLTokenKind::RBrace) {
        // Omitted argument or empty primary (e.g. FUNC(a,,b) or FUNC())
        // Return an empty ErrorLit without reporting error, and do NOT consume it
        auto node = std::make_unique<XLASTNode>(XLNodeKind::ErrorLit);
        node->text = "";
        return node;
    }

    if (tok.kind == XLTokenKind::Error) {
        if (tok.text != "{" && tok.text != "}" && tok.text != ";") {
            ctx.reportError("Invalid character '" + tok.text + "'", tok.offset);
        }
        auto node = std::make_unique<XLASTNode>(XLNodeKind::ErrorLit);
        node->text = tok.text;
        ctx.consume();
        return node;
    }

    if (tok.kind == XLTokenKind::End) {
        ctx.reportError("Unexpected end of expression", tok.offset);
        auto node = std::make_unique<XLASTNode>(XLNodeKind::ErrorLit);
        node->text = "#VALUE!";
        return node;
    }

    // Any other unexpected token (e.g. operators)
    ctx.reportError("Expected operand, found '" + tok.text + "'", tok.offset);
    auto node = std::make_unique<XLASTNode>(XLNodeKind::ErrorLit);
    node->text = "#VALUE!";
    ctx.consume();
    return node;
}

std::unique_ptr<XLASTNode> XLFormulaParser::parseFuncCall(std::string name, ParseContext& ctx)
{
    Expects(ctx.current().kind == XLTokenKind::LParen);
    ctx.consume();    // eat '('

    auto node = std::make_unique<XLASTNode>(XLNodeKind::FuncCall);
    // Store function name as uppercase; strip Excel future-function prefixes (_xlfn. / _xlws.)
    std::transform(name.begin(), name.end(), name.begin(), ::toupper);
    auto stripPrefix = [](std::string& n, const char* prefix) {
        const std::size_t len = std::char_traits<char>::length(prefix);
        if (n.size() > len && n.compare(0, len, prefix) == 0) n.erase(0, len);
    };
    stripPrefix(name, "_XLFN.");
    stripPrefix(name, "_XLWS.");
    node->text = std::move(name);

    // Parse argument list (comma or semicolon separated)
    while (ctx.current().kind != XLTokenKind::RParen && ctx.current().kind != XLTokenKind::End) {
        node->children.push_back(parseExpr(ctx, 0));
        if (ctx.current().kind == XLTokenKind::Comma || ctx.current().kind == XLTokenKind::Semicolon) {
            ctx.consume();
            ctx.panicMode = false;
        }
        else {
            break;
        }
    }
    if (ctx.current().kind == XLTokenKind::RParen) {
        ctx.consume();
        ctx.panicMode = false;
    } else {
        ctx.reportError("Missing closing parenthesis ')' for function '" + node->text + "'", ctx.current().offset);
    }
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
        // Single cell — construct a 1x1 LazyRange to preserve position metadata (for ROW/COLUMN)
        std::string refStr(rangeRef);
        // Strip sheet prefix if any
        std::string sheetPart;
        auto        bang = refStr.find('!');
        if (bang != std::string::npos) { sheetPart = refStr.substr(0, bang); refStr = refStr.substr(bang + 1); }
        // Strip $
        refStr.erase(std::remove(refStr.begin(), refStr.end(), '$'), refStr.end());
        try {
            XLCellReference cr(refStr);
            uint32_t r = cr.row();
            uint16_t c = cr.column();
            return XLFormulaArg(r, r, c, c, std::move(sheetPart), &resolver);
        } catch (...) {
            // Not a valid cell ref (e.g. named range) — fall back to resolver
            return XLFormulaArg(resolver(rangeRef));
        }
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


}

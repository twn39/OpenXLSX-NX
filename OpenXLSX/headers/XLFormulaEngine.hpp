#ifndef OPENXLSX_XLFORMULAENGINE_HPP
#define OPENXLSX_XLFORMULAENGINE_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

// ===== Standard Library ===== //
#include <functional>
#include <memory>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

// ===== GSL ===== //
#include <gsl/gsl>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCellReference.hpp"
#include "XLCellValue.hpp"

// Forward declare XLWorksheet so callers can use makeResolver without pulling the full header.
namespace OpenXLSX
{
    class XLWorksheet;
}

namespace OpenXLSX
{
    // =========================================================================
    // Lexer
    // =========================================================================

    /**
     * @brief Kinds of tokens produced by the lexer.
     */
    enum class XLTokenKind : uint8_t {
        Number,       ///< Numeric literal (integer or float)
        String,       ///< Quoted string literal  "hello"
        Bool,         ///< TRUE or FALSE keyword
        CellRef,      ///< Cell reference  A1, $B$2, Sheet1!C3
        Ident,        ///< Function name or named range
        Plus,         ///< +
        Minus,        ///< -
        Star,         ///< *
        Slash,        ///< /
        Caret,        ///< ^  (power)
        Percent,      ///< %
        Amp,          ///< &  (string concat)
        Eq,           ///< =
        NEq,          ///< <>
        Lt,           ///< <
        Le,           ///< <=
        Gt,           ///< >
        Ge,           ///< >=
        LParen,       ///< (
        RParen,       ///< )
        Comma,        ///< ,
        Semicolon,    ///< ;  (alternative argument separator in some locales)
        Colon,        ///< :  (used inside range references parsed by lexer)
        End,          ///< Sentinel – end of input
        Error         ///< Unrecognised character
    };

    /**
     * @brief A single lexical token from a formula string.
     */
    struct OPENXLSX_EXPORT XLToken
    {
        XLTokenKind kind{XLTokenKind::Error};
        std::string text;              ///< Raw text of the token (number string, identifier, …)
        double      number{0.0};       ///< Pre-parsed numeric value when kind == Number
        bool        boolean{false};    ///< Pre-parsed bool when kind == Bool
    };

    /**
     * @brief Tokenises a raw Excel formula string (with or without leading '=').
     * @details Hand-written character scanner; zero heap allocations during scanning.
     *          Range references like A1:B10 are emitted as a *single* CellRef token
     *          whose text is "A1:B10".
     */
    class OPENXLSX_EXPORT XLFormulaLexer
    {
    public:
        /**
         * @brief Tokenise the given formula.
         * @param formula The formula string (may start with '=').
         * @return Ordered vector of tokens, last element always XLTokenKind::End.
         */
        static std::vector<XLToken> tokenize(std::string_view formula);

    private:
        // Stateless helper – everything runs in tokenize().
    };

    // =========================================================================
    // AST
    // =========================================================================

    /** @brief Kind discriminator for AST nodes. */
    enum class XLNodeKind : uint8_t {
        Number,
        StringLit,
        BoolLit,
        CellRef,    ///< e.g. "A1" or "Sheet1!A1"
        Range,      ///< e.g. "A1:B10" – evaluated lazily by the resolver
        BinOp,
        UnaryOp,
        FuncCall,
        ErrorLit    ///< #NAME?, etc. – propagated as-is
    };

    /**
     * @brief Polymorphic AST node.  Uses a tagged-union approach with `std::unique_ptr` children.
     * @details All fields are public for simplicity; the AST is an internal implementation detail.
     */
    struct OPENXLSX_EXPORT XLASTNode
    {
        XLNodeKind kind;

        // ---- Leaf payloads ----
        double      number{0.0};
        std::string text;    ///< string literal, cell-ref text, range text, identifier, error text
        bool        boolean{false};
        XLTokenKind op{XLTokenKind::Error};    ///< operator for BinOp / UnaryOp

        // ---- Children ----
        std::vector<std::unique_ptr<XLASTNode>> children;    ///< operands or function arguments

        explicit XLASTNode(XLNodeKind k) : kind(k) {}

        // Non-copyable due to unique_ptr children; movable.
        XLASTNode(const XLASTNode&)            = delete;
        XLASTNode& operator=(const XLASTNode&) = delete;
        XLASTNode(XLASTNode&&)                 = default;
        XLASTNode& operator=(XLASTNode&&)      = default;
    };

    // =========================================================================
    // Parser
    // =========================================================================

    /**
     * @brief Recursive-descent (Pratt) parser that converts a token stream into an AST.
     */
    class OPENXLSX_EXPORT XLFormulaParser
    {
    public:
        /**
         * @brief Parse the token list produced by XLFormulaLexer::tokenize().
         * @param tokens Span over the token vector (including sentinel End token).
         * @return Root AST node.
         * @throws XLFormulaError on syntax error.
         */
        static std::unique_ptr<XLASTNode> parse(gsl::span<const XLToken> tokens);

    private:
        // All state is local to the recursive parse calls below.
        struct ParseContext
        {
            gsl::span<const XLToken> tokens;
            std::size_t              pos{0};

            [[nodiscard]] const XLToken& current() const;
            [[nodiscard]] const XLToken& peek(std::size_t offset = 1) const;
            const XLToken&               consume();
            bool                         matchKind(XLTokenKind k);
        };

        static std::unique_ptr<XLASTNode> parseExpr(ParseContext& ctx, int minPrec = 0);
        static std::unique_ptr<XLASTNode> parseUnary(ParseContext& ctx);
        static std::unique_ptr<XLASTNode> parsePrimary(ParseContext& ctx);
        static std::unique_ptr<XLASTNode> parseFuncCall(std::string name, ParseContext& ctx);

        static int  precedence(XLTokenKind k);
        static bool isRightAssoc(XLTokenKind k);
    };

    // =========================================================================
    // Evaluator / Engine
    // =========================================================================

    /**
     * @brief Callback type: resolves a cell reference string to a cell value.
     * @details The string is the raw reference text, e.g. "A1", "$B$2", "Sheet1!C3",
     *          or "A1:B10" for a range (the engine calls this once per distinct ref).
     *          Return `XLCellValue{}` (empty) for unknown or out-of-range cells.
     *
     *          **Range refs** are passed as-is.  The engine detects the colon in the
     *          text and will call the resolver with each individual cell in the range
     *          when it needs to expand the range.  The resolver only needs to handle
     *          single-cell refs; range expansion is done internally.
     */
    using XLCellResolver = std::function<XLCellValue(std::string_view ref)>;

    /**
     * @brief Lightweight formula evaluation engine.
     *
     * @details Usage:
     * @code
     *   XLFormulaEngine engine;
     *   auto resolver = XLFormulaEngine::makeResolver(worksheet);
     *   XLCellValue result = engine.evaluate("SUM(A1:C1)", resolver);
     * @endcode
     *
     * The engine is **thread-safe for concurrent evaluate() calls** after construction
     * (the function table is built once in the constructor and is read-only thereafter).
     */

    class OPENXLSX_EXPORT XLFormulaArg
    {
    public:
        enum class Type { Empty, Scalar, Array, LazyRange };

    private:
        Type                                                m_type{Type::Empty};
        XLCellValue                                         m_scalar;
        std::vector<XLCellValue>                            m_array;
        uint32_t                                            m_r1{0}, m_r2{0};
        uint16_t                                            m_c1{0}, m_c2{0};
        std::string                                         m_sheetName;
        const std::function<XLCellValue(std::string_view)>* m_resolver{nullptr};

    public:
        XLFormulaArg() = default;
        XLFormulaArg(XLCellValue v) : m_type(Type::Scalar), m_scalar(std::move(v)) {}
        XLFormulaArg(std::vector<XLCellValue> arr) : m_type(Type::Array), m_array(std::move(arr)) {}
        XLFormulaArg(uint32_t                                            r1,
                     uint32_t                                            r2,
                     uint16_t                                            c1,
                     uint16_t                                            c2,
                     std::string                                         sheetName,
                     const std::function<XLCellValue(std::string_view)>* resolver)
            : m_type(Type::LazyRange),
              m_r1(r1),
              m_r2(r2),
              m_c1(c1),
              m_c2(c2),
              m_sheetName(std::move(sheetName)),
              m_resolver(resolver)
        {}

        Type type() const { return m_type; }

        size_t rows() const
        {
            if (m_type == Type::LazyRange) return static_cast<size_t>(m_r2 - m_r1 + 1);
            if (m_type == Type::Array) return m_array.size();    // simplified, arrays might be 2d but we treat as 1d column
            return m_type == Type::Scalar ? 1 : 0;
        }

        size_t cols() const
        {
            if (m_type == Type::LazyRange) return static_cast<size_t>(m_c2 - m_c1 + 1);
            return m_type == Type::Scalar || m_type == Type::Array ? 1 : 0;
        }

        /// @brief Return the 1-based row of the top-left cell (LazyRange only; 0 for other types).
        uint32_t firstRow() const noexcept { return m_type == Type::LazyRange ? m_r1 : 0; }

        /// @brief Return the 1-based column of the top-left cell (LazyRange only; 0 for other types).
        uint16_t firstCol() const noexcept { return m_type == Type::LazyRange ? m_c1 : 0; }

        bool empty() const
        {
            if (m_type == Type::Empty) return true;
            if (m_type == Type::Scalar) return m_scalar.type() == XLValueType::Empty;
            if (m_type == Type::Array) return m_array.empty();
            return m_r1 > m_r2 || m_c1 > m_c2;
        }
        size_t size() const
        {
            if (m_type == Type::Empty) return 0;
            if (m_type == Type::Scalar) return 1;
            if (m_type == Type::Array) return m_array.size();
            return static_cast<size_t>(m_r2 - m_r1 + 1) * static_cast<size_t>(m_c2 - m_c1 + 1);
        }

        XLCellValue operator[](size_t index) const
        {
            if (m_type == Type::Scalar) return index == 0 ? m_scalar : XLCellValue();
            if (m_type == Type::Array) return index < m_array.size() ? m_array[index] : XLCellValue();
            if (m_type == Type::LazyRange) {
                uint16_t    w   = m_c2 - m_c1 + 1;
                uint32_t    row = m_r1 + static_cast<uint32_t>(index / w);
                uint16_t    col = m_c1 + static_cast<uint16_t>(index % w);
                std::string ref = m_sheetName;
                if (!ref.empty()) ref += "!";
                ref += XLCellReference(row, col).address();
                if (m_resolver && *m_resolver) return (*m_resolver)(ref);
            }
            return XLCellValue();
        }

        class Iterator
        {
            const XLFormulaArg* arg;
            size_t              index;

        public:
            Iterator(const XLFormulaArg* a, size_t i) : arg(a), index(i) {}
            XLCellValue operator*() const { return (*arg)[index]; }
            Iterator&   operator++()
            {
                ++index;
                return *this;
            }
            bool operator!=(const Iterator& other) const { return index != other.index; }
        };
        Iterator begin() const { return Iterator(this, 0); }
        Iterator end() const { return Iterator(this, size()); }
    };

    class OPENXLSX_EXPORT XLFormulaEngine
    {
    public:
        XLFormulaEngine();
        ~XLFormulaEngine() = default;

        XLFormulaEngine(const XLFormulaEngine&)            = delete;
        XLFormulaEngine& operator=(const XLFormulaEngine&) = delete;
        XLFormulaEngine(XLFormulaEngine&&)                 = default;
        XLFormulaEngine& operator=(XLFormulaEngine&&)      = default;

        /**
         * @brief Evaluate a formula string.
         * @param formula The formula text (with or without leading '=').
         * @param resolver Callback to look up cell values.  May be empty if the formula
         *        contains no cell references.
         * @return The computed XLCellValue; an error value on evaluation failure.
         */
        [[nodiscard]] XLCellValue evaluate(std::string_view formula, const XLCellResolver& resolver = {}) const;

        /**
         * @brief Create a CellResolver that reads live values from an XLWorksheet.
         * @param wks The source worksheet.
         * @return A resolver callable capturing a reference to @p wks.
         * @note The returned resolver is only valid while @p wks is alive.
         */
        [[nodiscard]] static XLCellResolver makeResolver(const XLWorksheet& wks);

    private:
        // ---- Internal evaluation helpers ----

        /**
         * @brief Expand a range string "A1:B3" using the resolver into a flat value list.
         */
        static XLFormulaArg expandRange(std::string_view rangeRef, const XLCellResolver& resolver);

        /**
         * @brief Collect numeric values from a mixed argument list (scalars + range vectors).
         */
        static std::vector<double> collectNumbers(const std::vector<XLCellValue>& flat, bool countBlanks = false);

        /**
         * @brief Evaluate a single AST node recursively.
         * @throws XLFormulaError on unrecoverable error.
         */
        [[nodiscard]] XLCellValue evalNode(const XLASTNode& node, const XLCellResolver& resolver) const;

        /**
         * @brief Expand a function argument into a flat vector of XLCellValue.
         *        Handles both scalar values and ranges.
         */
        [[nodiscard]] XLFormulaArg expandArg(const XLASTNode& argNode, const XLCellResolver& resolver) const;

        // ---- Built-in function table ----
        // Each entry maps an uppercase function name to its implementation.
        using FuncArgs = std::vector<XLCellValue>;    ///< all arguments flattened
        using FuncImpl = std::function<XLCellValue(const std::vector<XLFormulaArg>&)>;
        static const std::unordered_map<std::string, FuncImpl>& getBuiltins();
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLFORMULAENGINE_HPP

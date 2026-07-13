#ifndef OPENXLSX_XLFORMULAENGINE_HPP
#define OPENXLSX_XLFORMULAENGINE_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

// ===== Standard Library ===== //
#include <cstddef>
#include <functional>
#include <memory>
#include <mutex>
#include <optional>
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
#include "XLEvaluationContext.hpp"

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
        LBrace,       ///< {  (array constant)
        RBrace,       ///< }  (array constant)
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
        std::size_t offset{0};         ///< Position in the formula string
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
        ArrayLit,   ///< Excel array constant {1,2;3,4} — children row-major; number=rows, text=cols
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
    // Diagnostics
    // =========================================================================

    struct OPENXLSX_EXPORT XLFormulaDiagnostic
    {
        std::string message;
        std::size_t offset{0};
    };

    class OPENXLSX_EXPORT XLFormulaDiagnosticReporter
    {
    public:
        XLFormulaDiagnosticReporter() = default;
        explicit XLFormulaDiagnosticReporter(std::string formula) : m_formula(std::move(formula)) {}

        void reportError(std::string message, std::size_t offset);
        
        [[nodiscard]] const std::vector<XLFormulaDiagnostic>& diagnostics() const;
        [[nodiscard]] bool hasErrors() const;
        void clear();
        [[nodiscard]] std::string getFullReport() const;

    private:
        std::string                      m_formula;
        std::vector<XLFormulaDiagnostic> m_diagnostics;
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
         * @param reporter Optional diagnostic reporter to collect syntax errors.
         * @return Root AST node.
         * @throws XLFormulaError on syntax error if reporter is nullptr.
         */
        static std::unique_ptr<XLASTNode> parse(gsl::span<const XLToken> tokens, XLFormulaDiagnosticReporter* reporter = nullptr);

    private:
        // All state is local to the recursive parse calls below.
        struct ParseContext
        {
            gsl::span<const XLToken>         tokens;
            std::size_t                      pos{0};
            XLFormulaDiagnosticReporter*     reporter{nullptr};
            bool                             panicMode{false};

            [[nodiscard]] const XLToken& current() const;
            [[nodiscard]] const XLToken& peek(std::size_t offset = 1) const;
            const XLToken&               consume();
            bool                         matchKind(XLTokenKind k);

            void reportError(std::string message, std::size_t offset);
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
     * @brief Optional callback: resolve a defined name to a formula/reference string.
     * @details Return e.g. "Sheet1!$A$1:$B$10" or "=$A$1". Empty optional = unknown name.
     */
    using XLNameResolver = std::function<std::optional<std::string>(std::string_view name)>;

    // Forward-declare so XLEvalSession can appear before XLFormulaArg details are complete.
    class XLFormulaArg;
    class XLFormulaEngine;

    /**
     * @brief Per-evaluation session shared by the engine and session-aware functions.
     *
     * @details Phase A extension point for deep formula features:
     *          - cell resolution (via resolver)
     *          - defined-name resolution
     *          - current cell/sheet (for relative / parameterless ROW/COLUMN)
     *          - call-depth guard (INDIRECT recursion)
     *          - access to expandRange / evaluate helpers used by INDIRECT/OFFSET
     *
     *          Lifetime: valid only for the duration of a single `XLFormulaEngine::evaluate` call
     *          (or an explicitly scoped session created by the caller).
     */
    class OPENXLSX_EXPORT XLEvalSession
    {
    public:
        static constexpr int kMaxCallDepth = 128;

        XLEvalSession() = default;
        explicit XLEvalSession(const XLCellResolver& resolver) : m_resolver(&resolver) {}

        XLEvalSession& setResolver(const XLCellResolver& resolver)
        {
            m_resolver = &resolver;
            return *this;
        }

        XLEvalSession& setNameResolver(XLNameResolver nameResolver)
        {
            m_nameResolver = std::move(nameResolver);
            return *this;
        }

        /**
         * @brief Bind the cell that owns the formula (1-based Excel coordinates).
         * @details Enables parameterless ROW()/COLUMN() and future relative-ref features.
         */
        XLEvalSession& setCurrentCell(uint32_t row, uint16_t col)
        {
            m_currentRow     = row;
            m_currentCol     = col;
            m_hasCurrentCell = (row > 0 && col > 0);
            return *this;
        }

        XLEvalSession& setCurrentSheet(std::string sheetName)
        {
            m_currentSheet = std::move(sheetName);
            return *this;
        }

        [[nodiscard]] bool hasResolver() const { return m_resolver != nullptr && static_cast<bool>(*m_resolver); }

        [[nodiscard]] const XLCellResolver* resolverPtr() const { return m_resolver; }

        [[nodiscard]] XLCellValue cellValue(std::string_view ref) const
        {
            if (!hasResolver()) return XLCellValue{};
            return (*m_resolver)(ref);
        }

        [[nodiscard]] std::optional<std::string> resolveName(std::string_view name) const
        {
            if (!m_nameResolver) return std::nullopt;
            return m_nameResolver(name);
        }

        [[nodiscard]] bool     hasCurrentCell() const noexcept { return m_hasCurrentCell; }
        [[nodiscard]] uint32_t currentRow() const noexcept { return m_currentRow; }
        [[nodiscard]] uint16_t currentCol() const noexcept { return m_currentCol; }
        [[nodiscard]] const std::string& currentSheet() const noexcept { return m_currentSheet; }

        /**
         * @brief Enter a nested function/ref resolution (INDIRECT, etc.).
         * @return false if the maximum call depth would be exceeded.
         */
        bool enterCall()
        {
            if (m_callDepth >= kMaxCallDepth) return false;
            ++m_callDepth;
            return true;
        }

        void leaveCall() noexcept
        {
            if (m_callDepth > 0) --m_callDepth;
        }

        [[nodiscard]] int callDepth() const noexcept { return m_callDepth; }

        /**
         * @brief Expand an A1-style cell/range text into a LazyRange (or scalar error).
         * @note Requires a live resolver; used by INDIRECT/OFFSET and the engine.
         */
        [[nodiscard]] XLFormulaArg expandRange(std::string_view rangeRef) const;

    private:
        const XLCellResolver* m_resolver{nullptr};
        XLNameResolver        m_nameResolver;
        uint32_t              m_currentRow{0};
        uint16_t              m_currentCol{0};
        bool                  m_hasCurrentCell{false};
        std::string           m_currentSheet;
        int                   m_callDepth{0};
    };

    /**
     * @brief RAII guard that pairs XLEvalSession::enterCall / leaveCall.
     */
    class OPENXLSX_EXPORT XLEvalCallGuard
    {
    public:
        explicit XLEvalCallGuard(XLEvalSession& session) : m_session(&session), m_ok(session.enterCall()) {}
        ~XLEvalCallGuard()
        {
            if (m_ok && m_session) m_session->leaveCall();
        }
        XLEvalCallGuard(const XLEvalCallGuard&)            = delete;
        XLEvalCallGuard& operator=(const XLEvalCallGuard&) = delete;
        [[nodiscard]] bool ok() const noexcept { return m_ok; }

    private:
        XLEvalSession* m_session{nullptr};
        bool           m_ok{false};
    };

    /**
     * @brief Lightweight formula evaluation engine.
     *
     * @details Usage:
     * @code
     *   XLFormulaEngine engine;
     *   XLWorksheetEvaluationContext ctx(worksheet);   // sheet-read port
     *   auto resolver = XLFormulaEngine::makeResolver(ctx);
     *   XLCellValue result = engine.evaluate("SUM(A1:C1)", resolver);
     * @endcode
     *
     * The engine is **thread-safe for concurrent evaluate() calls** after construction
     * (the function table is built once in the constructor and is read-only thereafter).
     * Cell lookup goes through `XLEvaluationContext` / `XLCellResolver` so formula code
     * never depends on concrete worksheet internals.
     */

    class OPENXLSX_EXPORT XLFormulaArg
    {
    public:
        enum class Type { Empty, Scalar, Array, LazyRange };

    private:
        Type                                                m_type{Type::Empty};
        XLCellValue                                         m_scalar;
        std::vector<XLCellValue>                            m_array;    ///< row-major dense storage when Type::Array
        size_t                                              m_arrayRows{0};
        size_t                                              m_arrayCols{0};
        uint32_t                                            m_r1{0}, m_r2{0};
        uint16_t                                            m_c1{0}, m_c2{0};
        std::string                                         m_sheetName;
        const std::function<XLCellValue(std::string_view)>* m_resolver{nullptr};

    public:
        XLFormulaArg() = default;
        XLFormulaArg(XLCellValue v) : m_type(Type::Scalar), m_scalar(std::move(v)) {}

        /**
         * @brief Construct a column vector (N×1) from a flat list.
         * @details Preserves historical 1-D behaviour used by arithmetic broadcast helpers.
         */
        XLFormulaArg(std::vector<XLCellValue> arr)
            : m_type(Type::Array),
              m_array(std::move(arr)),
              m_arrayRows(m_array.size()),
              m_arrayCols(m_array.empty() ? 0 : 1)
        {}

        /**
         * @brief Construct a dense row-major 2-D array.
         * @param data Must contain rows*cols elements (row-major).
         */
        XLFormulaArg(std::vector<XLCellValue> data, size_t rows, size_t cols)
            : m_type(Type::Array), m_array(std::move(data)), m_arrayRows(rows), m_arrayCols(cols)
        {
            if (m_arrayRows * m_arrayCols != m_array.size()) {
                // Recover a consistent rectangular shape when possible.
                if (m_array.empty()) {
                    m_arrayRows = 0;
                    m_arrayCols = 0;
                }
                else if (cols > 0 && m_array.size() % cols == 0) {
                    m_arrayRows = m_array.size() / cols;
                    m_arrayCols = cols;
                }
                else {
                    m_arrayRows = m_array.size();
                    m_arrayCols = 1;
                }
            }
        }

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
            if (m_type == Type::Array) return m_arrayRows;
            return m_type == Type::Scalar ? 1 : 0;
        }

        size_t cols() const
        {
            if (m_type == Type::LazyRange) return static_cast<size_t>(m_c2 - m_c1 + 1);
            if (m_type == Type::Array) return m_arrayCols;
            return m_type == Type::Scalar ? 1 : 0;
        }

        /// @brief Return the 1-based row of the top-left cell (LazyRange only; 0 for other types).
        uint32_t firstRow() const noexcept { return m_type == Type::LazyRange ? m_r1 : 0; }

        /// @brief Return the 1-based column of the top-left cell (LazyRange only; 0 for other types).
        uint16_t firstCol() const noexcept { return m_type == Type::LazyRange ? m_c1 : 0; }

        /// @brief Bottom-right row for LazyRange (0 otherwise).
        uint32_t lastRow() const noexcept { return m_type == Type::LazyRange ? m_r2 : 0; }

        /// @brief Bottom-right column for LazyRange (0 otherwise).
        uint16_t lastCol() const noexcept { return m_type == Type::LazyRange ? m_c2 : 0; }

        [[nodiscard]] const std::string& sheetName() const noexcept { return m_sheetName; }

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

        /**
         * @brief 0-based (row, col) access. Works for Scalar, Array, and LazyRange.
         */
        [[nodiscard]] XLCellValue at(size_t row, size_t col) const
        {
            if (m_type == Type::Scalar) return (row == 0 && col == 0) ? m_scalar : XLCellValue();
            if (m_type == Type::Array) {
                if (row >= m_arrayRows || col >= m_arrayCols) return XLCellValue();
                return m_array[row * m_arrayCols + col];
            }
            if (m_type == Type::LazyRange) {
                if (row >= rows() || col >= cols()) return XLCellValue();
                return (*this)[row * cols() + col];
            }
            return XLCellValue();
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

        /**
         * @brief Materialize any arg into a dense row-major Array (copy).
         * @details Scalars become 1×1; LazyRange is fully resolved via the cell resolver.
         */
        [[nodiscard]] XLFormulaArg materialize() const
        {
            if (m_type == Type::Array) return *this;
            if (m_type == Type::Scalar) {
                std::vector<XLCellValue> d{m_scalar};
                return XLFormulaArg(std::move(d), 1, 1);
            }
            if (m_type == Type::LazyRange) {
                const size_t r = rows();
                const size_t c = cols();
                std::vector<XLCellValue> d;
                d.reserve(r * c);
                for (size_t i = 0; i < r * c; ++i) d.push_back((*this)[i]);
                return XLFormulaArg(std::move(d), r, c);
            }
            return XLFormulaArg();
        }

        /**
         * @brief Implicit-intersection / top-left scalar projection.
         */
        [[nodiscard]] XLCellValue asScalar() const
        {
            if (empty()) return XLCellValue{};
            if (m_type == Type::Scalar) return m_scalar;
            return (*this)[0];
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

    /**
     * @brief Callback used by spillArray to write one spill cell (1-based Excel coordinates).
     */
    using XLSpillWriter = std::function<void(uint32_t row, uint16_t col, const XLCellValue& value)>;

    /**
     * @brief Callback: return true if (row,col) is occupied and must not be overwritten by a spill
     *        (except the formula anchor cell itself).
     */
    using XLSpillOccupancy = std::function<bool(uint32_t row, uint16_t col)>;

    /**
     * @brief Result of a checked spill attempt (dynamic-array #SPILL! support).
     */
    struct OPENXLSX_EXPORT XLSpillResult
    {
        bool        ok{false};             ///< false when the spill range is blocked
        size_t      cellsWritten{0};       ///< number of cells written (0 on #SPILL!)
        XLCellValue error;                 ///< #SPILL! when !ok, otherwise empty
        size_t      rows{0};               ///< spill height
        size_t      cols{0};               ///< spill width
    };

    /**
     * @brief Write a (possibly multi-cell) formula result into a rectangular spill range.
     * @param result Materialized or lazy array/range/scalar result.
     * @param topRow 1-based top-left row of the spill anchor.
     * @param leftCol 1-based top-left column of the spill anchor.
     * @param write Invoked once per output cell (row-major).
     * @return Number of cells written.
     */
    OPENXLSX_EXPORT size_t spillArray(const XLFormulaArg& result, uint32_t topRow, uint16_t leftCol, const XLSpillWriter& write);

    /**
     * @brief Check whether a spill rectangle is free of blocking occupants.
     * @param isOccupied Return true for cells that block spill (non-empty foreign content).
     * @param anchorRow/anchorCol The formula cell itself is never treated as a blocker.
     */
    OPENXLSX_EXPORT bool spillRangeIsClear(const XLFormulaArg&   result,
                                           uint32_t              topRow,
                                           uint16_t              leftCol,
                                           const XLSpillOccupancy& isOccupied,
                                           uint32_t              anchorRow = 0,
                                           uint16_t              anchorCol = 0);

    /**
     * @brief Spill with #SPILL! detection: writes only when the target range is clear.
     * @details On conflict, no cells are written and result.error is set to #SPILL!.
     */
    OPENXLSX_EXPORT XLSpillResult spillArrayChecked(const XLFormulaArg&     result,
                                                    uint32_t                topRow,
                                                    uint16_t                leftCol,
                                                    const XLSpillWriter&    write,
                                                    const XLSpillOccupancy& isOccupied,
                                                    uint32_t                anchorRow = 0,
                                                    uint16_t                anchorCol = 0);

    class OPENXLSX_EXPORT XLFormulaEngine
    {
    public:
        XLFormulaEngine();
        ~XLFormulaEngine() = default;

        XLFormulaEngine(const XLFormulaEngine&)            = delete;
        XLFormulaEngine& operator=(const XLFormulaEngine&) = delete;
        XLFormulaEngine(XLFormulaEngine&&)                 = delete;
        XLFormulaEngine& operator=(XLFormulaEngine&&)      = delete;

        // ---- Phase D: AST cache (formula string → parsed tree) ----

        /** Enable/disable parse-tree caching (default: enabled). */
        void setAstCacheEnabled(bool enabled) noexcept { m_astCacheEnabled = enabled; }
        [[nodiscard]] bool astCacheEnabled() const noexcept { return m_astCacheEnabled; }

        /** Max distinct formulas retained in the cache (default 512). LRU-ish drop of arbitrary entry when full. */
        void setAstCacheCapacity(std::size_t capacity) noexcept;
        [[nodiscard]] std::size_t astCacheCapacity() const noexcept { return m_astCacheCapacity; }

        /** Number of cached ASTs. */
        [[nodiscard]] std::size_t astCacheSize() const;

        /** Drop all cached parse trees. */
        void clearAstCache();

        /**
         * @brief Evaluate a formula string.
         * @param formula The formula text (with or without leading '=').
         * @param resolver Callback to look up cell values.  May be empty if the formula
         *        contains no cell references.
         * @return The computed XLCellValue; an error value on evaluation failure.
         *
         * @note Prefer pairing with an `XLEvaluationContext` via `makeResolver(context)` so
         *       formula code depends only on the sheet-read port, not a concrete worksheet type:
         * @code
         *   XLWorksheetEvaluationContext ctx(wks);
         *   engine.evaluate("=A1+B1", XLFormulaEngine::makeResolver(ctx));
         * @endcode
         */
        [[nodiscard]] XLCellValue evaluate(std::string_view formula, const XLCellResolver& resolver = {}, XLFormulaDiagnosticReporter* reporter = nullptr) const;

        /**
         * @brief Evaluate with a full evaluation session (names, current cell, depth, …).
         * @details Multi-cell array results are reduced via implicit intersection (top-left).
         *          Use evaluateArray() when the full spill shape is required.
         */
        [[nodiscard]] XLCellValue evaluate(std::string_view formula, XLEvalSession& session, XLFormulaDiagnosticReporter* reporter = nullptr) const;

        /**
         * @brief Evaluate a formula preserving full array/range shape (Phase B spill path).
         */
        [[nodiscard]] XLFormulaArg evaluateArray(std::string_view formula, const XLCellResolver& resolver = {}, XLFormulaDiagnosticReporter* reporter = nullptr) const;

        /**
         * @brief Session-aware array evaluation (FILTER/UNIQUE/SORT/SEQUENCE, range arithmetic, …).
         */
        [[nodiscard]] XLFormulaArg evaluateArray(std::string_view formula, XLEvalSession& session, XLFormulaDiagnosticReporter* reporter = nullptr) const;

        /**
         * @brief Create a CellResolver that reads live values from an evaluation context.
         * @param context The abstract sheet-read port.
         * @return A resolver callable capturing a reference to @p context.
         * @note The returned resolver is only valid while @p context is alive.
         */
        [[nodiscard]] static XLCellResolver makeResolver(const XLEvaluationContext& context);

        /**
         * @brief Create a CellResolver that reads live values from an XLWorksheet.
         * @param wks The source worksheet.
         * @return A resolver callable capturing a reference to @p wks.
         * @note The returned resolver is only valid while @p wks is alive.
         * @note Prefer `XLWorksheetEvaluationContext` + `evaluate(formula, context)` for new code.
         */
        [[nodiscard]] static XLCellResolver makeResolver(const XLWorksheet& wks);

        /**
         * @brief Expand an A1 cell/range reference using the given resolver.
         * @details Public so session-aware functions (INDIRECT, OFFSET) can build LazyRanges.
         */
        [[nodiscard]] static XLFormulaArg expandRange(std::string_view rangeRef, const XLCellResolver& resolver);

    private:
        // ---- Internal evaluation helpers ----

        /**
         * @brief Collect numeric values from a mixed argument list (scalars + range vectors).
         */
        static std::vector<double> collectNumbers(const std::vector<XLCellValue>& flat, bool countBlanks = false);

        /**
         * @brief Evaluate a single AST node recursively under a session (scalar projection).
         */
        [[nodiscard]] XLCellValue evalNode(const XLASTNode& node, XLEvalSession& session) const;

        /**
         * @brief Evaluate a node as a full-shaped argument (ranges / arrays preserved).
         */
        [[nodiscard]] XLFormulaArg evalAsArg(const XLASTNode& node, XLEvalSession& session) const;

        /**
         * @brief Expand a function argument (range / cell / INDIRECT / OFFSET / scalar).
         */
        [[nodiscard]] XLFormulaArg expandArg(const XLASTNode& argNode, XLEvalSession& session) const;

        /**
         * @brief Evaluate reference-returning functions (INDIRECT, OFFSET) as XLFormulaArg.
         */
        [[nodiscard]] XLFormulaArg evalRefFunction(const XLASTNode& node, XLEvalSession& session) const;

        /**
         * @brief Dispatch a FuncCall through the builtin table, returning full array shape.
         */
        [[nodiscard]] XLFormulaArg evalFunctionAsArg(const XLASTNode& node, XLEvalSession& session) const;

        [[nodiscard]] static bool isRefReturningFunction(std::string_view name);

        // ---- Built-in function table ----
        // Session-aware: every builtin returns XLFormulaArg (scalar or multi-cell array).
        using FuncImpl = std::function<XLFormulaArg(const std::vector<XLFormulaArg>&, XLEvalSession&)>;
        static const std::unordered_map<std::string, FuncImpl>& getBuiltins();

        /** Normalize formula key for the AST cache (strip leading '=', trim). */
        [[nodiscard]] static std::string cacheKey(std::string_view formula);

        /** Parse (or fetch cached) AST. When @p reporter is non-null, cache is bypassed. */
        [[nodiscard]] std::shared_ptr<XLASTNode> getOrParseAst(std::string_view formula,
                                                               XLFormulaDiagnosticReporter* reporter) const;

        bool                                              m_astCacheEnabled{true};
        std::size_t                                       m_astCacheCapacity{512};
        mutable std::mutex                                m_astCacheMutex;
        mutable std::unordered_map<std::string, std::shared_ptr<XLASTNode>> m_astCache;
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLFORMULAENGINE_HPP

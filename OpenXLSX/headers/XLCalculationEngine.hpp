#ifndef OPENXLSX_XLCALCULATIONENGINE_HPP
#define OPENXLSX_XLCALCULATIONENGINE_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"
#include "XLFormulaEngine.hpp"

#include <cstddef>
#include <cstdint>
#include <queue>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

namespace OpenXLSX
{
    class XLWorksheet;
    class XLDocument;

    /**
     * @brief Normalized cell identity used by the calculation graph.
     * @details In single-sheet mode, `sheet` is empty for the engine default sheet.
     *          In multi-sheet (workbook) mode, `sheet` is always set to the home sheet name.
     *          Coordinates are 1-based Excel indices.
     */
    struct OPENXLSX_EXPORT XLCellKey
    {
        std::string sheet;
        uint32_t    row{0};
        uint16_t    col{0};

        [[nodiscard]] bool valid() const noexcept { return row > 0 && col > 0; }

        [[nodiscard]] bool operator==(const XLCellKey& other) const noexcept
        {
            return row == other.row && col == other.col && sheet == other.sheet;
        }

        [[nodiscard]] std::string address() const;
        [[nodiscard]] std::string qualifiedAddress() const;

        /**
         * @brief Parse "A1", "$B$2", "Sheet1!C3", or "'My Sheet'!A1".
         * @param defaultSheet Optional default sheet name for unqualified refs (left empty on the key
         *        unless @p forceDefaultSheet is true).
         * @param forceDefaultSheet When true, unqualified refs get sheet=defaultSheet (workbook mode).
         */
        [[nodiscard]] static XLCellKey parse(std::string_view ref,
                                             std::string_view defaultSheet     = {},
                                             bool             forceDefaultSheet = false);
    };

    struct OPENXLSX_EXPORT XLCellKeyHash
    {
        size_t operator()(const XLCellKey& k) const noexcept;
    };

    /**
     * @brief Options for workbook/sheet formula recalculation (Phase C / C-deep).
     */
    struct OPENXLSX_EXPORT XLCalculationOptions
    {
        /** Write computed values back into worksheet / value map (cached result). */
        bool writeBack{true};

        /** Maximum recursion depth for single-cell evaluation. */
        int maxDepth{256};

        /**
         * @brief Cap on expanded range dependencies per formula.
         * @details Larger ranges are not expanded into the graph; the cell still
         *          reads live values through the resolver during evaluation.
         */
        size_t maxExpandedDeps{4096};

        /**
         * @brief When markDirty is called, also dirty all transitive dependents.
         */
        bool propagateDirty{true};

        /**
         * @brief After recalculate / recalculateAll on a document-backed engine,
         *        set the workbook fullCalcOnLoad flag (Excel recalc prompt on open).
         */
        bool setFullCalcOnLoad{false};

        /**
         * @brief Error token returned for circular references (default Excel-ish #CIRCREF!).
         * @details Stored as XLValueType::Error text.
         */
        std::string circularErrorToken{"#CIRCREF!"};

        /**
         * @brief When true (default), register with XLDocument to auto-dirty on cell value/formula writes.
         */
        bool autoTrackChanges{true};

        /**
         * @brief Load workbook defined names into the dependency extractor (default true for sheet/doc engines).
         */
        bool useDefinedNames{true};
    };

    /**
     * @brief Result status for a single calcCellValue / recalculate step.
     */
    enum class XLCalcStatus : uint8_t {
        Ok = 0,
        Circular,
        Error,
        Empty
    };

    /**
     * @brief Sheet- or workbook-scoped formula calculation engine (Phase C).
     *
     * @details Builds a dependency graph over formula cells, evaluates with caching
     *          and circular detection, propagates dirty flags to dependents, and can
     *          recalculate only dirty cells or the full graph.
     *
     * @code
     *   XLCalculationEngine calc(wks);
     *   calc.recalculateAll();
     *   calc.setInputValue("A1", 99);   // dirties B1, C1, …
     *   calc.recalculate();             // only dirty formulas
     *
     *   XLCalculationEngine book(doc);  // all worksheets
     *   book.recalculateAll();
     * @endcode
     */
    class OPENXLSX_EXPORT XLCalculationEngine
    {
    public:
        using FormulaMap = std::unordered_map<std::string, std::string>;    ///< "A1" or "Sheet2!A1" -> formula
        using ValueMap   = std::unordered_map<std::string, XLCellValue>;

        /** Single worksheet (sheet-local graph; empty sheet keys). */
        explicit XLCalculationEngine(XLWorksheet& worksheet, XLCalculationOptions options = {});

        /**
         * @brief All worksheets in a document (multi-sheet graph; keys always carry sheet name).
         * @note The document must outlive this engine.
         */
        explicit XLCalculationEngine(XLDocument& document, XLCalculationOptions options = {});

        /**
         * @brief In-memory model for unit tests.
         * @param formulas Map of address -> formula. Use "Sheet!A1" for multi-sheet graphs.
         * @param values   Constant / seed values (same key conventions).
         * @param defaultSheet Optional default sheet name for unqualified addresses.
         */
        XLCalculationEngine(FormulaMap             formulas,
                            ValueMap               values,
                            XLCalculationOptions   options      = {},
                            std::string            defaultSheet = {});

        ~XLCalculationEngine();

        XLCalculationEngine(const XLCalculationEngine&)            = delete;
        XLCalculationEngine& operator=(const XLCalculationEngine&) = delete;
        XLCalculationEngine(XLCalculationEngine&&)                 = delete;
        XLCalculationEngine& operator=(XLCalculationEngine&&)      = delete;

        [[nodiscard]] const XLCalculationOptions& options() const noexcept { return m_options; }
        void setOptions(XLCalculationOptions options) { m_options = std::move(options); }

        [[nodiscard]] bool isMultiSheet() const noexcept { return m_multiSheet; }

        /** (Re)scan formula cells and rebuild dependency + reverse-dependency graphs. */
        void rebuild();

        [[nodiscard]] size_t formulaCount() const noexcept { return m_nodes.size(); }

        /** Number of formula cells currently marked dirty. */
        [[nodiscard]] size_t dirtyCount() const noexcept;

        /** Direct precedents (cells this formula reads). */
        [[nodiscard]] std::vector<XLCellKey> dependencies(const XLCellKey& cell) const;
        [[nodiscard]] std::vector<XLCellKey> dependencies(std::string_view a1) const;

        /** Direct dependents (formulas that read this cell). Includes non-formula seeds as keys. */
        [[nodiscard]] std::vector<XLCellKey> dependents(const XLCellKey& cell) const;
        [[nodiscard]] std::vector<XLCellKey> dependents(std::string_view a1) const;

        /**
         * @brief Evaluate one cell, recursively calculating precedent formula cells.
         * @return Computed value; circular refs return Error with circularErrorToken.
         */
        [[nodiscard]] XLCellValue calcCellValue(const XLCellKey& cell);
        [[nodiscard]] XLCellValue calcCellValue(std::string_view a1);

        [[nodiscard]] XLCalcStatus lastStatus() const noexcept { return m_lastStatus; }

        /**
         * @brief Recalculate only dirty formula cells.
         * @return Number of formula cells successfully evaluated (non-circular).
         */
        size_t recalculate();

        /**
         * @brief Mark all formulas dirty, then recalculate.
         */
        size_t recalculateAll();

        /**
         * @brief Mark a cell dirty; with propagateDirty, also all transitive dependents.
         * @param cell Any cell (input seed or formula).
         * @param propagate Override options.propagateDirty when set.
         */
        void markDirty(const XLCellKey& cell, bool propagate);
        void markDirty(const XLCellKey& cell);
        void markDirty(std::string_view a1, bool propagate);
        void markDirty(std::string_view a1);

        void markAllDirty();

        /**
         * @brief Update an input value and dirty all dependent formulas.
         * @details Writes to the value map or worksheet cell, then markDirty(propagate).
         */
        void setInputValue(const XLCellKey& cell, const XLCellValue& value);
        void setInputValue(std::string_view a1, const XLCellValue& value);

        /**
         * @brief Notify that a cell changed externally (same as markDirty with propagation).
         */
        void notifyChanged(const XLCellKey& cell);
        void notifyChanged(std::string_view a1);

        /** Clear evaluation caches and visit colours (keeps formula text / deps). */
        void clearCache();

        /**
         * @brief Inject defined names for map-mode engines (name → refersTo formula, e.g. "Sales" → "=$B$2").
         * @details Names are matched case-insensitively. Call before rebuild() or rebuild after set.
         */
        void setDefinedNames(std::unordered_map<std::string, std::string> names);

        /**
         * @brief Reload defined names from the bound workbook (sheet/document engines).
         */
        void reloadDefinedNames();

        /**
         * @brief Locally update one formula cell in the graph without full rebuild.
         * @details Reads the live formula from the worksheet/map; removes the node if no formula.
         * @return true if the node exists after the update (still a formula cell).
         */
        bool updateFormulaCell(const XLCellKey& cell);
        bool updateFormulaCell(std::string_view a1);

        /**
         * @brief Extract cell/range dependencies from a formula string (public for tests).
         * @param forceDefaultSheet When true, unqualified refs get sheet=defaultSheet.
         * @param definedNames Optional uppercase-name → refersTo map for expanding named ranges.
         */
        [[nodiscard]] static std::vector<XLCellKey> extractDependencies(
            std::string_view                                              formula,
            std::string_view                                              defaultSheet      = {},
            size_t                                                        maxExpanded       = 4096,
            bool                                                          forceDefaultSheet = false,
            const std::unordered_map<std::string, std::string>*           definedNames      = nullptr);

    private:
        enum class VisitColor : uint8_t { White = 0, Gray, Black };

        struct Node
        {
            std::string            formula;
            std::vector<XLCellKey> deps;
            XLCellValue            cached;
            bool                   hasCache{false};
            bool                   dirty{true};
            VisitColor             color{VisitColor::White};
        };

        XLCellValue calcCellValueImpl(const XLCellKey& cell, int depth);
        XLCellValue readInputValue(const XLCellKey& cell) const;
        void        writeBackValue(const XLCellKey& cell, const XLCellValue& value);
        void        rebuildFromWorksheet();
        void        rebuildFromDocument();
        void        rebuildFromMaps();
        void        rebuildDependentsIndex();
        void        collectFormulasFromSheet(XLWorksheet& wks, const std::string& sheetName);
        XLCellKey   makeKey(uint32_t row, uint16_t col, std::string_view sheet = {}) const;
        XLCellKey   normalizeLookupKey(const XLCellKey& key) const;
        XLCellValue circularErrorValue() const;
        void        maybeSetFullCalcOnLoad();
        size_t      recalculateKeys(const std::vector<XLCellKey>& keys);
        void        attachDocumentListener();
        void        detachDocumentListener();
        void        onExternalCellChange(std::string_view sheet, uint32_t row, uint16_t col, bool formulaChanged);
        std::vector<XLCellKey> extractDepsForContext(std::string_view formula, std::string_view homeSheet) const;
        std::string            readLiveFormula(const XLCellKey& key) const;

        XLWorksheet*         m_worksheet{nullptr};
        XLDocument*          m_document{nullptr};
        std::string          m_defaultSheet;
        FormulaMap           m_mapFormulas;
        ValueMap             m_mapValues;
        bool                 m_useMaps{false};
        bool                 m_multiSheet{false};
        bool                 m_listenerAttached{false};
        XLCalculationOptions m_options{};
        XLFormulaEngine      m_engine{};
        std::unordered_map<std::string, std::string> m_definedNames;    ///< UPPER name → refersTo

        std::unordered_map<XLCellKey, Node, XLCellKeyHash>              m_nodes;
        std::unordered_map<XLCellKey, std::vector<XLCellKey>, XLCellKeyHash> m_dependents;
        XLCalcStatus                                                    m_lastStatus{XLCalcStatus::Empty};
        bool                                                            m_circularSeen{false};
        bool                                                            m_suppressExternalDirty{false};    ///< ignore self write-back notifies
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLCALCULATIONENGINE_HPP

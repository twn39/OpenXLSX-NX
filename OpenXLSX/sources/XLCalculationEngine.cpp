#include "XLCalculationEngine.hpp"

#include "XLCell.hpp"
#include "XLCellRange.hpp"
#include "XLCellReference.hpp"
#include "XLCommandQuery.hpp"
#include "XLDocument.hpp"
#include "XLEvaluationContext.hpp"
#include "XLFormulaUtils.hpp"
#include "XLWorkbook.hpp"
#include "XLWorksheet.hpp"

#include <algorithm>
#include <cctype>
#include <functional>
#include <optional>
#include <queue>

namespace OpenXLSX
{
    namespace
    {
        std::string_view unquoteSheet(std::string_view name)
        {
            if (name.size() >= 2 && name.front() == '\'' && name.back() == '\'') {
                return name.substr(1, name.size() - 2);
            }
            return name;
        }

        void stripDollars(std::string& s) { s.erase(std::remove(s.begin(), s.end(), '$'), s.end()); }

        bool sameSheet(std::string_view a, std::string_view b)
        {
            if (a.empty() || b.empty()) return true;
            return a == b;
        }

        bool keyLess(const XLCellKey& a, const XLCellKey& b)
        {
            if (a.sheet != b.sheet) return a.sheet < b.sheet;
            if (a.row != b.row) return a.row < b.row;
            return a.col < b.col;
        }
    }    // namespace

    // =============================================================================
    // XLCellKey
    // =============================================================================

    std::string XLCellKey::address() const
    {
        if (!valid()) return {};
        return XLCellReference(row, col).address();
    }

    std::string XLCellKey::qualifiedAddress() const
    {
        auto addr = address();
        if (sheet.empty()) return addr;
        bool needQuote = false;
        for (char c : sheet) {
            if (!(std::isalnum(static_cast<unsigned char>(c)) || c == '_')) {
                needQuote = true;
                break;
            }
        }
        if (needQuote) return "'" + sheet + "'!" + addr;
        return sheet + "!" + addr;
    }

    XLCellKey XLCellKey::parse(std::string_view ref, std::string_view defaultSheet, bool forceDefaultSheet)
    {
        XLCellKey key;

        std::string text(ref);
        if (!text.empty() && text.front() == '=') text.erase(text.begin());

        auto bang = text.find('!');
        if (bang != std::string::npos) {
            key.sheet = std::string(unquoteSheet(text.substr(0, bang)));
            text      = text.substr(bang + 1);
        }
        else if (forceDefaultSheet && !defaultSheet.empty()) {
            key.sheet = std::string(defaultSheet);
        }
        stripDollars(text);
        if (text.find(':') != std::string::npos) return {};

        try {
            XLCellReference cr(text);
            key.row = cr.row();
            key.col = cr.column();
        }
        catch (...) {
            return {};
        }
        return key;
    }

    size_t XLCellKeyHash::operator()(const XLCellKey& k) const noexcept
    {
        size_t h = std::hash<uint32_t>{}(k.row);
        h ^= std::hash<uint16_t>{}(k.col) + 0x9e3779b9 + (h << 6) + (h >> 2);
        h ^= std::hash<std::string>{}(k.sheet) + 0x9e3779b9 + (h << 6) + (h >> 2);
        return h;
    }

    // =============================================================================
    // Dependency extraction
    // =============================================================================

    std::vector<XLCellKey> XLCalculationEngine::extractDependencies(
        std::string_view                                    formula,
        std::string_view                                    defaultSheet,
        size_t                                              maxExpanded,
        bool                                                forceDefaultSheet,
        const std::unordered_map<std::string, std::string>* definedNames)
    {
        std::vector<XLCellKey>                       deps;
        std::unordered_set<XLCellKey, XLCellKeyHash> seen;
        std::unordered_set<std::string>              expandingNames;

        auto addKey = [&](XLCellKey k) {
            if (!k.valid()) return;
            if (!forceDefaultSheet && !defaultSheet.empty() && k.sheet == defaultSheet) k.sheet.clear();
            if (seen.insert(k).second) deps.push_back(std::move(k));
        };

        auto expandRangeText = [&](std::string_view rangeText) {
            std::string text(rangeText);
            std::string sheetPart;
            auto        bang = text.find('!');
            if (bang != std::string::npos) {
                sheetPart = std::string(unquoteSheet(text.substr(0, bang)));
                text      = text.substr(bang + 1);
            }
            else if (forceDefaultSheet && !defaultSheet.empty()) {
                sheetPart = std::string(defaultSheet);
            }
            auto colon = text.find(':');
            if (colon == std::string::npos) {
                addKey(XLCellKey::parse(rangeText, defaultSheet, forceDefaultSheet));
                return;
            }
            std::string startRef = text.substr(0, colon);
            std::string endRef   = text.substr(colon + 1);
            auto        endBang  = endRef.find('!');
            if (endBang != std::string::npos) endRef = endRef.substr(endBang + 1);
            stripDollars(startRef);
            stripDollars(endRef);
            try {
                XLCellReference a(startRef);
                XLCellReference b(endRef);
                uint32_t        r1 = a.row(), r2 = b.row();
                uint16_t        c1 = a.column(), c2 = b.column();
                if (r1 > r2) std::swap(r1, r2);
                if (c1 > c2) std::swap(c1, c2);
                const size_t cells = static_cast<size_t>(r2 - r1 + 1) * static_cast<size_t>(c2 - c1 + 1);
                // Phase D: full expansion when small; otherwise stride-sample + corners
                // so large ranges (e.g. SUM(A:A)) still participate in dirty tracking.
                if (cells <= maxExpanded) {
                    for (uint32_t r = r1; r <= r2; ++r) {
                        for (uint16_t c = c1; c <= c2; ++c) {
                            XLCellKey k;
                            k.sheet = sheetPart;
                            k.row   = r;
                            k.col   = c;
                            addKey(std::move(k));
                        }
                    }
                }
                else {
                    auto pushRC = [&](uint32_t r, uint16_t c) {
                        XLCellKey k;
                        k.sheet = sheetPart;
                        k.row   = r;
                        k.col   = c;
                        addKey(std::move(k));
                    };
                    // Four corners always
                    pushRC(r1, c1);
                    pushRC(r1, c2);
                    pushRC(r2, c1);
                    pushRC(r2, c2);
                    // Stride sample ~maxExpanded cells
                    const size_t rowSpan = static_cast<size_t>(r2 - r1) + 1;
                    const size_t colSpan = static_cast<size_t>(c2 - c1) + 1;
                    const size_t stride  = std::max<size_t>(1, cells / maxExpanded);
                    size_t       idx     = 0;
                    size_t       added   = 0;
                    for (uint32_t r = r1; r <= r2 && added < maxExpanded; ++r) {
                        for (uint16_t c = c1; c <= c2 && added < maxExpanded; ++c, ++idx) {
                            if (idx % stride == 0) {
                                pushRC(r, c);
                                ++added;
                            }
                        }
                    }
                    (void)rowSpan;
                    (void)colSpan;
                }
            }
            catch (...) {
            }
        };

        std::function<void(std::string_view)> expandFormula;
        expandFormula = [&](std::string_view fml) {
            auto tokens = XLFormulaLexer::tokenize(fml);
            for (const auto& t : tokens) {
                if (t.kind == XLTokenKind::CellRef) {
                    if (t.text.find(':') != std::string::npos) {
                        expandRangeText(t.text);
                    }
                    else {
                        auto k = XLCellKey::parse(t.text, defaultSheet, forceDefaultSheet);
                        if (k.valid()) {
                            addKey(std::move(k));
                        }
                        else if (definedNames != nullptr) {
                            std::string un = t.text;
                            std::transform(un.begin(), un.end(), un.begin(), [](unsigned char c) {
                                return static_cast<char>(std::toupper(c));
                            });
                            auto it = definedNames->find(un);
                            if (it != definedNames->end() && expandingNames.insert(un).second) {
                                expandFormula(it->second);
                                expandingNames.erase(un);
                            }
                        }
                    }
                }
                else if (t.kind == XLTokenKind::Ident && definedNames != nullptr) {
                    std::string un = t.text;
                    std::transform(un.begin(), un.end(), un.begin(), [](unsigned char c) {
                        return static_cast<char>(std::toupper(c));
                    });
                    auto it = definedNames->find(un);
                    if (it != definedNames->end() && expandingNames.insert(un).second) {
                        expandFormula(it->second);
                        expandingNames.erase(un);
                    }
                }
            }
        };

        expandFormula(formula);
        return deps;
    }

    std::vector<XLCellKey> XLCalculationEngine::extractDepsForContext(std::string_view formula,
                                                                      std::string_view homeSheet) const
    {
        const auto* names = m_definedNames.empty() ? nullptr : &m_definedNames;
        return extractDependencies(formula, homeSheet, m_options.maxExpandedDeps, m_multiSheet, names);
    }

    // =============================================================================
    // Construction
    // =============================================================================

    XLCalculationEngine::XLCalculationEngine(XLWorksheet& worksheet, XLCalculationOptions options)
        : m_worksheet(&worksheet), m_useMaps(false), m_multiSheet(false), m_options(std::move(options))
    {
        try {
            m_defaultSheet = worksheet.name();
            m_document     = &worksheet.parentDoc();
        }
        catch (...) {
            m_defaultSheet.clear();
        }
        if (m_options.useDefinedNames) reloadDefinedNames();
        rebuild();
        attachDocumentListener();
    }

    XLCalculationEngine::XLCalculationEngine(XLDocument& document, XLCalculationOptions options)
        : m_worksheet(nullptr), m_document(&document), m_useMaps(false), m_multiSheet(true), m_options(std::move(options))
    {
        if (m_options.useDefinedNames) reloadDefinedNames();
        rebuild();
        attachDocumentListener();
    }

    XLCalculationEngine::XLCalculationEngine(FormulaMap           formulas,
                                             ValueMap             values,
                                             XLCalculationOptions options,
                                             std::string          defaultSheet)
        : m_worksheet(nullptr),
          m_mapFormulas(std::move(formulas)),
          m_mapValues(std::move(values)),
          m_useMaps(true),
          m_options(std::move(options)),
          m_defaultSheet(std::move(defaultSheet))
    {
        // Multi-sheet map graph if any key or formula contains '!' or defaultSheet set with qualified keys.
        m_multiSheet = false;
        for (const auto& [addr, _] : m_mapFormulas) {
            if (addr.find('!') != std::string::npos) {
                m_multiSheet = true;
                break;
            }
        }
        if (!m_multiSheet) {
            for (const auto& [addr, _] : m_mapValues) {
                if (addr.find('!') != std::string::npos) {
                    m_multiSheet = true;
                    break;
                }
            }
        }
        rebuild();
    }

    XLCalculationEngine::~XLCalculationEngine() { detachDocumentListener(); }

    void XLCalculationEngine::attachDocumentListener()
    {
        if (!m_options.autoTrackChanges || m_document == nullptr || m_listenerAttached) return;
        m_document->registerCellChangeListener(this, [this](std::string_view sheet, uint32_t row, uint16_t col, bool formulaChanged) {
            onExternalCellChange(sheet, row, col, formulaChanged);
        });
        m_listenerAttached = true;
    }

    void XLCalculationEngine::detachDocumentListener()
    {
        if (!m_listenerAttached || m_document == nullptr) return;
        try {
            m_document->unregisterCellChangeListener(this);
        }
        catch (...) {
        }
        m_listenerAttached = false;
    }

    void XLCalculationEngine::onExternalCellChange(std::string_view sheet, uint32_t row, uint16_t col, bool formulaChanged)
    {
        if (m_suppressExternalDirty) return;

        XLCellKey key = makeKey(row, col, sheet);
        // Ignore single-sheet engine events for other sheets when multi is false.
        if (!m_multiSheet && !sheet.empty() && !m_defaultSheet.empty() && sheet != m_defaultSheet) return;

        if (formulaChanged) {
            updateFormulaCell(key);
        }
        else {
            markDirty(key, true);
        }
    }

    void XLCalculationEngine::setDefinedNames(std::unordered_map<std::string, std::string> names)
    {
        m_definedNames.clear();
        for (auto& [k, v] : names) {
            std::string uk = k;
            std::transform(uk.begin(), uk.end(), uk.begin(), [](unsigned char c) { return static_cast<char>(std::toupper(c)); });
            m_definedNames[std::move(uk)] = std::move(v);
        }
    }

    void XLCalculationEngine::reloadDefinedNames()
    {
        m_definedNames.clear();
        if (m_document == nullptr && m_worksheet != nullptr) {
            try {
                m_document = &m_worksheet->parentDoc();
            }
            catch (...) {
            }
        }
        if (m_document == nullptr) return;
        try {
            auto names = m_document->workbook().definedNames().all();
            for (const auto& dn : names) {
                if (!dn.valid()) continue;
                std::string n = dn.name();
                std::transform(n.begin(), n.end(), n.begin(), [](unsigned char c) { return static_cast<char>(std::toupper(c)); });
                // Skip Excel internals like _xlnm.Print_Area
                if (n.rfind("_XLNM.", 0) == 0) continue;
                m_definedNames[std::move(n)] = dn.refersTo();
            }
        }
        catch (...) {
        }
    }

    XLCellKey XLCalculationEngine::makeKey(uint32_t row, uint16_t col, std::string_view sheet) const
    {
        XLCellKey k;
        k.row = row;
        k.col = col;
        if (m_multiSheet) {
            k.sheet = sheet.empty() ? m_defaultSheet : std::string(sheet);
        }
        else {
            k.sheet.clear();
        }
        return k;
    }

    XLCellKey XLCalculationEngine::normalizeLookupKey(const XLCellKey& key) const
    {
        if (!m_multiSheet) {
            XLCellKey k = key;
            k.sheet.clear();
            return k;
        }
        XLCellKey k = key;
        if (k.sheet.empty()) k.sheet = m_defaultSheet;
        return k;
    }

    XLCellValue XLCalculationEngine::circularErrorValue() const
    {
        XLCellValue e;
        e.setError(m_options.circularErrorToken.empty() ? std::string("#CIRCREF!") : m_options.circularErrorToken);
        return e;
    }

    // =============================================================================
    // Rebuild
    // =============================================================================

    void XLCalculationEngine::rebuild()
    {
        m_nodes.clear();
        m_dependents.clear();
        m_lastStatus = XLCalcStatus::Empty;
        if (m_useMaps)
            rebuildFromMaps();
        else if (m_document != nullptr)
            rebuildFromDocument();
        else
            rebuildFromWorksheet();
        rebuildDependentsIndex();
    }

    void XLCalculationEngine::rebuildDependentsIndex()
    {
        m_dependents.clear();
        for (const auto& [cell, node] : m_nodes) {
            for (const auto& dep : node.deps) {
                m_dependents[dep].push_back(cell);
            }
        }
    }

    void XLCalculationEngine::rebuildFromMaps()
    {
        const bool forceSheet = m_multiSheet;
        for (const auto& [addr, formula] : m_mapFormulas) {
            auto key = XLCellKey::parse(addr, m_defaultSheet, forceSheet);
            if (!key.valid()) continue;
            // Unqualified map keys in multi-sheet mode with empty defaultSheet stay local-empty;
            // force at least empty sheet identity.
            if (m_multiSheet && key.sheet.empty() && !m_defaultSheet.empty()) key.sheet = m_defaultSheet;

            Node node;
            node.formula               = formula;
            const std::string ctxSheet = key.sheet.empty() ? m_defaultSheet : key.sheet;
            node.deps                  = extractDepsForContext(formula, ctxSheet);
            if (m_multiSheet) {
                for (auto& d : node.deps) {
                    if (d.sheet.empty()) d.sheet = ctxSheet;
                }
            }
            else {
                for (auto& d : node.deps) d.sheet.clear();
            }
            node.dirty = true;
            m_nodes.emplace(std::move(key), std::move(node));
        }
    }

    void XLCalculationEngine::collectFormulasFromSheet(XLWorksheet& wks, const std::string& sheetName)
    {
        try {
            auto rng = wks.range();
            for (auto& cell : rng) {
                if (!cell.hasFormula()) continue;
                std::string formula;
                try {
                    formula = cell.formula().get();
                }
                catch (...) {
                    continue;
                }
                if (formula.empty()) continue;

                auto ref = cell.cellReference();
                auto key = makeKey(ref.row(), ref.column(), sheetName);

                Node node;
                node.formula = std::move(formula);
                const std::string home = m_multiSheet ? sheetName : (sheetName.empty() ? m_defaultSheet : sheetName);
                node.deps              = extractDepsForContext(node.formula, home);
                if (m_multiSheet) {
                    for (auto& d : node.deps) {
                        if (d.sheet.empty()) d.sheet = sheetName;
                    }
                }
                else {
                    for (auto& d : node.deps) d.sheet.clear();
                }
                node.dirty = true;
                m_nodes.emplace(std::move(key), std::move(node));
            }
        }
        catch (...) {
        }
    }

    std::string XLCalculationEngine::readLiveFormula(const XLCellKey& key) const
    {
        if (m_useMaps) {
            auto it = m_mapFormulas.find(key.address());
            if (it != m_mapFormulas.end()) return it->second;
            it = m_mapFormulas.find(key.qualifiedAddress());
            if (it != m_mapFormulas.end()) return it->second;
            return {};
        }
        try {
            if (m_document != nullptr) {
                const std::string sheet = key.sheet.empty() ? m_defaultSheet : key.sheet;
                auto              wks   = m_document->workbook().worksheet(sheet);
                if (!wks.cell(key.row, key.col).hasFormula()) return {};
                return wks.cell(key.row, key.col).formula().get();
            }
            if (m_worksheet != nullptr) {
                if (!m_worksheet->cell(key.row, key.col).hasFormula()) return {};
                return m_worksheet->cell(key.row, key.col).formula().get();
            }
        }
        catch (...) {
        }
        return {};
    }

    bool XLCalculationEngine::updateFormulaCell(std::string_view a1)
    {
        return updateFormulaCell(XLCellKey::parse(a1, m_defaultSheet, m_multiSheet));
    }

    bool XLCalculationEngine::updateFormulaCell(const XLCellKey& cellIn)
    {
        XLCellKey key = normalizeLookupKey(cellIn);
        if (!key.valid()) return false;

        std::string formula = readLiveFormula(key);
        // Map-mode: also allow caller to have updated m_mapFormulas already.
        if (m_useMaps && formula.empty()) {
            // leave as empty → remove node
        }

        if (formula.empty()) {
            // Formula cleared: drop node and rebuild reverse index.
            m_nodes.erase(key);
            if (!m_multiSheet) {
                XLCellKey local = key;
                local.sheet.clear();
                m_nodes.erase(local);
            }
            rebuildDependentsIndex();
            // Dependents of this cell may still need recalc (they referenced a value/formula that changed).
            markDirty(key, true);
            return false;
        }

        const std::string home = key.sheet.empty() ? m_defaultSheet : key.sheet;
        Node              node;
        node.formula = std::move(formula);
        node.deps    = extractDepsForContext(node.formula, home);
        if (m_multiSheet) {
            for (auto& d : node.deps)
                if (d.sheet.empty()) d.sheet = home;
        }
        else {
            for (auto& d : node.deps) d.sheet.clear();
        }
        node.dirty    = true;
        node.hasCache = false;
        node.color    = VisitColor::White;

        m_nodes[key] = std::move(node);
        rebuildDependentsIndex();
        markDirty(key, true);
        return true;
    }

    void XLCalculationEngine::rebuildFromWorksheet()
    {
        if (m_worksheet == nullptr) return;
        collectFormulasFromSheet(*m_worksheet, m_multiSheet ? m_defaultSheet : std::string{});
    }

    void XLCalculationEngine::rebuildFromDocument()
    {
        if (m_document == nullptr) return;
        try {
            auto names = m_document->workbook().worksheetNames();
            if (!names.empty()) m_defaultSheet = names.front();
            for (const auto& name : names) {
                try {
                    auto wks = m_document->workbook().worksheet(name);
                    collectFormulasFromSheet(wks, name);
                }
                catch (...) {
                }
            }
        }
        catch (...) {
        }
    }

    // =============================================================================
    // Dirty / cache / dependents
    // =============================================================================

    size_t XLCalculationEngine::dirtyCount() const noexcept
    {
        size_t n = 0;
        for (const auto& [_, node] : m_nodes)
            if (node.dirty) ++n;
        return n;
    }

    void XLCalculationEngine::clearCache()
    {
        for (auto& [_, node] : m_nodes) {
            node.hasCache = false;
            node.dirty    = true;
            node.color    = VisitColor::White;
            node.cached   = XLCellValue{};
        }
        m_lastStatus = XLCalcStatus::Empty;
    }

    void XLCalculationEngine::markAllDirty()
    {
        for (auto& [_, node] : m_nodes) {
            node.dirty    = true;
            node.hasCache = false;
            node.color    = VisitColor::White;
        }
    }

    void XLCalculationEngine::markDirty(const XLCellKey& cell)
    {
        markDirty(cell, m_options.propagateDirty);
    }

    void XLCalculationEngine::markDirty(std::string_view a1)
    {
        markDirty(XLCellKey::parse(a1, m_defaultSheet, m_multiSheet), m_options.propagateDirty);
    }

    void XLCalculationEngine::markDirty(std::string_view a1, bool propagate)
    {
        markDirty(XLCellKey::parse(a1, m_defaultSheet, m_multiSheet), propagate);
    }

    void XLCalculationEngine::markDirty(const XLCellKey& cell, bool propagate)
    {
        XLCellKey start = normalizeLookupKey(cell);
        if (!start.valid()) return;

        if (!propagate) {
            auto it = m_nodes.find(start);
            if (it != m_nodes.end()) {
                it->second.dirty    = true;
                it->second.hasCache = false;
                it->second.color    = VisitColor::White;
            }
            // Also try empty-sheet form for single-sheet engines.
            if (!m_multiSheet && !cell.sheet.empty()) {
                XLCellKey local = cell;
                local.sheet.clear();
                auto it2 = m_nodes.find(local);
                if (it2 != m_nodes.end()) {
                    it2->second.dirty    = true;
                    it2->second.hasCache = false;
                    it2->second.color    = VisitColor::White;
                }
            }
            return;
        }

        // BFS over reverse edges (and the seed itself if it is a formula).
        std::queue<XLCellKey>                 q;
        std::unordered_set<XLCellKey, XLCellKeyHash> visited;
        q.push(start);
        visited.insert(start);

        while (!q.empty()) {
            XLCellKey cur = q.front();
            q.pop();

            auto itNode = m_nodes.find(cur);
            if (itNode == m_nodes.end() && !m_multiSheet) {
                XLCellKey local = cur;
                local.sheet.clear();
                itNode = m_nodes.find(local);
                if (itNode != m_nodes.end()) cur = local;
            }
            if (itNode != m_nodes.end()) {
                itNode->second.dirty    = true;
                itNode->second.hasCache = false;
                itNode->second.color    = VisitColor::White;
            }

            // Look up dependents under both multi and local key forms.
            auto pushDeps = [&](const XLCellKey& k) {
                auto it = m_dependents.find(k);
                if (it == m_dependents.end()) return;
                for (const auto& dep : it->second) {
                    if (visited.insert(dep).second) q.push(dep);
                }
            };
            pushDeps(cur);
            if (!cur.sheet.empty()) {
                XLCellKey local = cur;
                local.sheet.clear();
                pushDeps(local);
            }
            if (cur.sheet.empty() && !m_defaultSheet.empty()) {
                XLCellKey named = cur;
                named.sheet     = m_defaultSheet;
                pushDeps(named);
            }
        }
    }

    void XLCalculationEngine::notifyChanged(const XLCellKey& cell) { markDirty(cell, true); }
    void XLCalculationEngine::notifyChanged(std::string_view a1) { markDirty(a1, true); }

    void XLCalculationEngine::setInputValue(std::string_view a1, const XLCellValue& value)
    {
        setInputValue(XLCellKey::parse(a1, m_defaultSheet, m_multiSheet), value);
    }

    void XLCalculationEngine::setInputValue(const XLCellKey& cell, const XLCellValue& value)
    {
        XLCellKey key = normalizeLookupKey(cell);
        if (!key.valid()) return;

        if (m_useMaps) {
            // Prefer qualified key when multi-sheet.
            if (m_multiSheet && !key.sheet.empty())
                m_mapValues[key.qualifiedAddress()] = value;
            m_mapValues[key.address()] = value;
            if (m_multiSheet && !key.sheet.empty()) m_mapValues[key.qualifiedAddress()] = value;
        }
        else if (m_document != nullptr) {
            try {
                const std::string sheet = key.sheet.empty() ? m_defaultSheet : key.sheet;
                auto              wks   = m_document->workbook().worksheet(sheet);
                wks.cell(key.row, key.col).value() = value;
            }
            catch (...) {
            }
        }
        else if (m_worksheet != nullptr) {
            try {
                m_worksheet->cell(key.row, key.col).value() = value;
            }
            catch (...) {
            }
        }

        markDirty(key, true);
    }

    std::vector<XLCellKey> XLCalculationEngine::dependencies(const XLCellKey& cell) const
    {
        auto it = m_nodes.find(normalizeLookupKey(cell));
        if (it == m_nodes.end()) {
            XLCellKey local = cell;
            local.sheet.clear();
            it = m_nodes.find(local);
        }
        if (it == m_nodes.end()) return {};
        return it->second.deps;
    }

    std::vector<XLCellKey> XLCalculationEngine::dependencies(std::string_view a1) const
    {
        return dependencies(XLCellKey::parse(a1, m_defaultSheet, m_multiSheet));
    }

    std::vector<XLCellKey> XLCalculationEngine::dependents(const XLCellKey& cell) const
    {
        XLCellKey key = normalizeLookupKey(cell);
        auto      it  = m_dependents.find(key);
        if (it == m_dependents.end() && !key.sheet.empty()) {
            XLCellKey local = key;
            local.sheet.clear();
            it = m_dependents.find(local);
        }
        if (it == m_dependents.end() && key.sheet.empty() && !m_defaultSheet.empty()) {
            XLCellKey named = key;
            named.sheet     = m_defaultSheet;
            it              = m_dependents.find(named);
        }
        if (it == m_dependents.end()) return {};
        return it->second;
    }

    std::vector<XLCellKey> XLCalculationEngine::dependents(std::string_view a1) const
    {
        return dependents(XLCellKey::parse(a1, m_defaultSheet, m_multiSheet));
    }

    // =============================================================================
    // Value I/O
    // =============================================================================

    XLCellValue XLCalculationEngine::readInputValue(const XLCellKey& cell) const
    {
        if (m_useMaps) {
            if (m_multiSheet && !cell.sheet.empty()) {
                auto itQ = m_mapValues.find(cell.qualifiedAddress());
                if (itQ != m_mapValues.end()) return itQ->second;
            }
            auto it = m_mapValues.find(cell.address());
            if (it != m_mapValues.end()) return it->second;
            it = m_mapValues.find(cell.qualifiedAddress());
            if (it != m_mapValues.end()) return it->second;
            return XLCellValue{};
        }

        if (m_document != nullptr) {
            try {
                const std::string sheet = cell.sheet.empty() ? m_defaultSheet : cell.sheet;
                auto              wks   = m_document->workbook().worksheet(sheet);
                return wks.cell(cell.row, cell.col).value();
            }
            catch (...) {
                return XLCellValue{};
            }
        }

        if (m_worksheet == nullptr) return XLCellValue{};

        try {
            if (!cell.sheet.empty() && !sameSheet(cell.sheet, m_defaultSheet)) {
                auto other = m_worksheet->parentDoc().workbook().worksheet(cell.sheet);
                return other.cell(cell.row, cell.col).value();
            }
            XLWorksheetEvaluationContext ctx(*m_worksheet);
            return ctx.cellValue(cell.address());
        }
        catch (...) {
            return XLCellValue{};
        }
    }

    void XLCalculationEngine::writeBackValue(const XLCellKey& cell, const XLCellValue& value)
    {
        if (!m_options.writeBack) return;

        // Prevent write-back from re-dirtying the graph via document listeners.
        const bool prev              = m_suppressExternalDirty;
        m_suppressExternalDirty      = true;

        if (m_useMaps) {
            if (m_multiSheet && !cell.sheet.empty())
                m_mapValues[cell.qualifiedAddress()] = value;
            else
                m_mapValues[cell.address()] = value;
            m_suppressExternalDirty = prev;
            return;
        }

        if (m_document != nullptr) {
            try {
                const std::string sheet = cell.sheet.empty() ? m_defaultSheet : cell.sheet;
                auto              wks   = m_document->workbook().worksheet(sheet);
                wks.cell(cell.row, cell.col).value() = value;
            }
            catch (...) {
            }
            m_suppressExternalDirty = prev;
            return;
        }

        if (m_worksheet == nullptr) {
            m_suppressExternalDirty = prev;
            return;
        }
        try {
            if (!cell.sheet.empty() && !sameSheet(cell.sheet, m_defaultSheet)) {
                auto other = m_worksheet->parentDoc().workbook().worksheet(cell.sheet);
                other.cell(cell.row, cell.col).value() = value;
            }
            else {
                m_worksheet->cell(cell.row, cell.col).value() = value;
            }
        }
        catch (...) {
        }
        m_suppressExternalDirty = prev;
    }

    void XLCalculationEngine::maybeSetFullCalcOnLoad()
    {
        if (!m_options.setFullCalcOnLoad) return;
        try {
            if (m_document != nullptr) {
                m_document->setFormulaNeedsRecalculation(true);
                m_document->execCommand(XLCommand(XLCommandType::SetFullCalcOnLoad));
            }
            else if (m_worksheet != nullptr) {
                m_worksheet->parentDoc().setFormulaNeedsRecalculation(true);
                m_worksheet->parentDoc().execCommand(XLCommand(XLCommandType::SetFullCalcOnLoad));
            }
        }
        catch (...) {
        }
    }

    // =============================================================================
    // Evaluation
    // =============================================================================

    XLCellValue XLCalculationEngine::calcCellValue(std::string_view a1)
    {
        return calcCellValue(XLCellKey::parse(a1, m_defaultSheet, m_multiSheet));
    }

    XLCellValue XLCalculationEngine::calcCellValue(const XLCellKey& cell)
    {
        for (auto& [_, node] : m_nodes) {
            if (node.color == VisitColor::Gray) node.color = VisitColor::White;
        }
        m_circularSeen = false;
        auto result    = calcCellValueImpl(normalizeLookupKey(cell), 0);
        if (m_circularSeen) m_lastStatus = XLCalcStatus::Circular;
        return result;
    }

    XLCellValue XLCalculationEngine::calcCellValueImpl(const XLCellKey& cellIn, int depth)
    {
        XLCellKey cell = normalizeLookupKey(cellIn);
        if (!cell.valid()) {
            if (!m_circularSeen) m_lastStatus = XLCalcStatus::Empty;
            return XLCellValue{};
        }
        if (depth > m_options.maxDepth) {
            if (!m_circularSeen) m_lastStatus = XLCalcStatus::Error;
            return formula::errRef();
        }

        auto it = m_nodes.find(cell);
        if (it == m_nodes.end() && !m_multiSheet) {
            XLCellKey local = cell;
            local.sheet.clear();
            it = m_nodes.find(local);
            if (it != m_nodes.end()) cell = local;
        }
        // Multi-sheet: try empty sheet if default
        if (it == m_nodes.end() && m_multiSheet && !cell.sheet.empty()) {
            // already full key
        }

        if (it == m_nodes.end()) {
            if (!m_circularSeen) m_lastStatus = XLCalcStatus::Ok;
            return readInputValue(cell);
        }

        Node& node = it->second;

        if (node.color == VisitColor::Gray) {
            m_circularSeen = true;
            m_lastStatus   = XLCalcStatus::Circular;
            return circularErrorValue();
        }

        if (node.hasCache && !node.dirty) {
            if (!m_circularSeen) m_lastStatus = XLCalcStatus::Ok;
            return node.cached;
        }

        node.color = VisitColor::Gray;

        const std::string homeSheet = cell.sheet.empty() ? m_defaultSheet : cell.sheet;

        // Live formula refresh from worksheet / document.
        if (!m_useMaps) {
            try {
                if (m_document != nullptr) {
                    auto wks  = m_document->workbook().worksheet(homeSheet);
                    auto live = wks.cell(cell.row, cell.col).formula().get();
                    if (!live.empty()) node.formula = std::move(live);
                }
                else if (m_worksheet != nullptr) {
                    auto live = m_worksheet->cell(cell.row, cell.col).formula().get();
                    if (!live.empty()) node.formula = std::move(live);
                }
            }
            catch (...) {
            }
        }

        XLCellResolver resolver = [this, depth, homeSheet](std::string_view ref) -> XLCellValue {
            auto key = XLCellKey::parse(ref, homeSheet, m_multiSheet);
            if (!key.valid()) {
                if (m_worksheet) return XLWorksheetEvaluationContext(*m_worksheet).cellValue(ref);
                if (m_document) {
                    try {
                        auto wks = m_document->workbook().worksheet(homeSheet);
                        return XLWorksheetEvaluationContext(wks).cellValue(ref);
                    }
                    catch (...) {
                    }
                }
                return XLCellValue{};
            }
            if (m_multiSheet && key.sheet.empty()) key.sheet = homeSheet;

            auto itNode = m_nodes.find(key);
            if (itNode == m_nodes.end() && !m_multiSheet) {
                XLCellKey local = key;
                local.sheet.clear();
                itNode = m_nodes.find(local);
                if (itNode != m_nodes.end()) key = local;
            }
            if (itNode != m_nodes.end()) return calcCellValueImpl(key, depth + 1);
            return readInputValue(key);
        };

        XLEvalSession session(resolver);
        session.setCurrentCell(cell.row, cell.col);
        if (!homeSheet.empty()) session.setCurrentSheet(homeSheet);
        if (!m_definedNames.empty()) {
            session.setNameResolver([this](std::string_view name) -> std::optional<std::string> {
                std::string un(name);
                std::transform(un.begin(), un.end(), un.begin(), [](unsigned char c) {
                    return static_cast<char>(std::toupper(c));
                });
                auto it = m_definedNames.find(un);
                if (it == m_definedNames.end()) return std::nullopt;
                return it->second;
            });
        }

        XLCellValue result;
        try {
            result = m_engine.evaluate(node.formula, session);
            if (m_circularSeen) {
                m_lastStatus = XLCalcStatus::Circular;
            }
            else if (result.type() == XLValueType::Error) {
                m_lastStatus = XLCalcStatus::Error;
            }
            else {
                m_lastStatus = XLCalcStatus::Ok;
            }
        }
        catch (...) {
            result = formula::errValue();
            if (!m_circularSeen) m_lastStatus = XLCalcStatus::Error;
        }

        if (m_circularSeen) {
            node.color    = VisitColor::White;
            node.dirty    = true;
            node.hasCache = false;
            return result.type() == XLValueType::Error ? result : circularErrorValue();
        }

        node.cached   = result;
        node.hasCache = true;
        node.dirty    = false;
        node.color    = VisitColor::Black;

        writeBackValue(cell, result);
        return result;
    }

    size_t XLCalculationEngine::recalculateKeys(const std::vector<XLCellKey>& keys)
    {
        size_t ok = 0;
        for (const auto& key : keys) {
            for (auto& [_, n] : m_nodes) {
                if (n.color == VisitColor::Gray) n.color = VisitColor::White;
            }
            m_circularSeen = false;
            (void)calcCellValueImpl(key, 0);
            if (m_circularSeen) m_lastStatus = XLCalcStatus::Circular;
            if (m_lastStatus != XLCalcStatus::Circular) ++ok;
        }
        return ok;
    }

    size_t XLCalculationEngine::recalculate()
    {
        if (m_nodes.empty()) rebuild();

        std::vector<XLCellKey> keys;
        keys.reserve(m_nodes.size());
        for (const auto& [k, node] : m_nodes) {
            if (node.dirty) keys.push_back(k);
        }
        std::sort(keys.begin(), keys.end(), keyLess);

        size_t ok = recalculateKeys(keys);
        maybeSetFullCalcOnLoad();
        return ok;
    }

    size_t XLCalculationEngine::recalculateAll()
    {
        if (m_nodes.empty()) rebuild();
        markAllDirty();
        return recalculate();
    }

}    // namespace OpenXLSX

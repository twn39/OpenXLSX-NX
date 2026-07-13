#include "XLFormulaFunctionsInternal.hpp"
#include "XLFormulaRegistry.hpp"
#include "XLDateTime.hpp"
#include "XLNumberFormatter.hpp"
#include "XLFormulaUtils.hpp"
#include <algorithm>
#include <array>
#include <cmath>
#include <limits>
#include <numeric>
#include <optional>
#include <random>
#include <regex>
#include <unordered_set>
#include <cctype>
using namespace OpenXLSX;

namespace
{
    std::tm info_serialToTm(double serial)
    {
        try {
            return XLDateTime(serial).tm();
        }
        catch (...) {
            return {};
        }
    }

    double info_tmToSerial(std::tm t)
    {
        try {
            return XLDateTime(t).serial();
        }
        catch (...) {
            return 0.0;
        }
    }

    int info_daysInMonth(int year, int month)
    {
        constexpr std::array<int, 12> days = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
        bool leap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
        int d = days.at(static_cast<std::size_t>(month - 1));
        if (month == 2 && leap) ++d;
        return d;
    }
} // namespace

XLCellValue Formula::fnIsnumber(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    return XLCellValue(isNumeric(args[0][0]));
}

XLCellValue Formula::fnIsblank(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(true);
    return XLCellValue(isEmpty(args[0][0]));
}

XLCellValue Formula::fnIserror(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    return XLCellValue(isError(args[0][0]));
}

XLCellValue Formula::fnIstext(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    return XLCellValue(args[0][0].type() == XLValueType::String);
}

XLCellValue Formula::fnIsna(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    const auto& v = args[0][0];
    if (v.type() != XLValueType::Error) return XLCellValue(false);
    return XLCellValue(v.get<std::string>() == "#N/A");
}

XLCellValue Formula::fnIfna(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const auto& v = args[0][0];
    if (v.type() == XLValueType::Error && v.get<std::string>() == "#N/A") return args[1][0];
    return v;
}

XLCellValue Formula::fnIslogical(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    return XLCellValue(args[0][0].type() == XLValueType::Boolean);
}

XLCellValue Formula::fnIsnontext(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(true);    // blank is non-text
    const auto& v = args[0][0];
    return XLCellValue(v.type() != XLValueType::String);
}

XLCellValue Formula::fnIserr(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(false);
    const auto& v = args[0][0];
    if (v.type() != XLValueType::Error) return XLCellValue(false);
    return XLCellValue(v.get<std::string>() != "#N/A");
}

namespace
{
    bool lookupWildcardMatch(const std::string& text, const std::string& pattern)
    {
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
    }

    std::string lookupLower(std::string s)
    {
        std::transform(s.begin(), s.end(), s.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
        return s;
    }

    /** Exact equality for lookup: numeric exact, text case-insensitive, optional wildcards in pattern. */
    bool lookupExactEqual(const XLCellValue& key, const XLCellValue& lookup, bool allowWildcard)
    {
        if (isError(key) || isError(lookup)) return false;
        if (isNumeric(key) && isNumeric(lookup)) return toDouble(key) == toDouble(lookup);
        // Type mismatch: do not coerce number vs pure text for exact match (Excel-ish).
        if (isNumeric(key) != isNumeric(lookup)) {
            // Exception: blank never equals a number.
            if (isEmpty(key) || isEmpty(lookup)) return false;
            // Both non-numeric path below for bool/string.
            if (isNumeric(key) || isNumeric(lookup)) return false;
        }
        std::string keyStr = lookupLower(toString(key));
        std::string lookStr = lookupLower(toString(lookup));
        if (allowWildcard && (lookStr.find('*') != std::string::npos || lookStr.find('?') != std::string::npos)) {
            return lookupWildcardMatch(keyStr, lookStr);
        }
        return keyStr == lookStr;
    }

    /**
     * @brief Compare key to lookup for ordering: <0 key<lookup, 0 equal, >0 key>lookup.
     * @return optional empty if incomparable.
     */
    std::optional<int> lookupCompare(const XLCellValue& key, const XLCellValue& lookup)
    {
        if (isEmpty(key) || isEmpty(lookup)) return std::nullopt;
        if (isNumeric(key) && isNumeric(lookup)) {
            double d = toDouble(key) - toDouble(lookup);
            if (d < 0) return -1;
            if (d > 0) return 1;
            return 0;
        }
        if (isNumeric(key) != isNumeric(lookup)) return std::nullopt;
        std::string a = lookupLower(toString(key));
        std::string b = lookupLower(toString(lookup));
        if (a < b) return -1;
        if (a > b) return 1;
        return 0;
    }

    /**
     * @brief Approximate match index (0-based): largest key <= lookup (ascending table).
     * @details Scans left-to-right; last qualifying key wins (Excel VLOOKUP TRUE / MATCH 1).
     */
    int lookupApproxLessEqual(const XLFormulaArg& keys, const XLCellValue& lookup)
    {
        int best = -1;
        for (std::size_t i = 0; i < keys.size(); ++i) {
            auto cmp = lookupCompare(keys[i], lookup);
            if (!cmp.has_value()) continue;
            if (*cmp <= 0) best = static_cast<int>(i);
            // If table is sorted ascending, once key > lookup we can stop early for MATCH/VLOOKUP.
            // Keep scanning for robustness with unsorted data (still last <= wins).
        }
        return best;
    }

    /** Smallest key >= lookup (MATCH type -1). */
    int lookupApproxGreaterEqual(const XLFormulaArg& keys, const XLCellValue& lookup)
    {
        int best = -1;
        for (std::size_t i = 0; i < keys.size(); ++i) {
            auto cmp = lookupCompare(keys[i], lookup);
            if (!cmp.has_value()) continue;
            if (*cmp >= 0) {
                if (best < 0) best = static_cast<int>(i);
                else {
                    auto bestCmp = lookupCompare(keys[static_cast<std::size_t>(best)], keys[i]);
                    // Prefer the smallest key that is still >= lookup (first in descending-sorted table).
                    if (bestCmp.has_value() && *bestCmp > 0) best = static_cast<int>(i);
                }
            }
        }
        // For descending tables Excel walks until values become smaller; prefer first match from left
        // that is >= lookup among those that qualify with minimal key.
        // Recompute: among all keys >= lookup, pick the one with smallest key; if ties, first index.
        int    chosen = -1;
        double bestKey = std::numeric_limits<double>::infinity();
        bool   bestIsNum = false;
        std::string bestStr;
        for (std::size_t i = 0; i < keys.size(); ++i) {
            auto cmp = lookupCompare(keys[i], lookup);
            if (!cmp.has_value() || *cmp < 0) continue;
            if (isNumeric(keys[i])) {
                double d = toDouble(keys[i]);
                if (chosen < 0 || !bestIsNum || d < bestKey) {
                    chosen   = static_cast<int>(i);
                    bestKey  = d;
                    bestIsNum = true;
                }
            }
            else if (!bestIsNum) {
                std::string s = lookupLower(toString(keys[i]));
                if (chosen < 0 || s < bestStr) {
                    chosen  = static_cast<int>(i);
                    bestStr = std::move(s);
                }
            }
        }
        return chosen;
    }

    XLCellValue lookupReturnAt(const XLFormulaArg& table, int row, int col, int nCols)
    {
        if (row < 0 || col < 0) return errRef();
        std::size_t idx = static_cast<std::size_t>(row * nCols + col);
        if (idx >= table.size()) return errRef();
        return table[idx];
    }
}    // namespace

XLCellValue Formula::fnVlookup(const std::vector<XLFormulaArg>& args)
{
    // VLOOKUP(lookup_value, table_array, col_index, [range_lookup])
    // range_lookup: FALSE/0 = exact (wildcards ok); TRUE/1/omitted historically TRUE in Excel
    // but many engines default exact — we follow Excel: omitted = TRUE (approximate).
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();

    const XLCellValue& lookupVal = args[0][0];
    auto               table     = args[1].materialize();
    int                colIdx    = static_cast<int>(toDouble(args[2][0]));    // 1-based
    // Excel default range_lookup is TRUE when omitted.
    bool exact = false;
    if (args.size() >= 4 && !args[3].empty()) {
        if (args[3][0].type() == XLValueType::Boolean)
            exact = !args[3][0].get<bool>();
        else
            exact = (toDouble(args[3][0]) == 0.0);
    }

    if (colIdx < 1) return errValue();
    int nCols = static_cast<int>(table.cols());
    if (nCols < 1) nCols = 1;
    if (colIdx > nCols) return errRef();
    int nRows = static_cast<int>(table.rows());

    // Build first-column key vector as column arg for helpers.
    std::vector<XLCellValue> keyCol;
    keyCol.reserve(static_cast<std::size_t>(nRows));
    for (int r = 0; r < nRows; ++r) keyCol.push_back(table.at(static_cast<std::size_t>(r), 0));
    XLFormulaArg keys(std::move(keyCol), static_cast<std::size_t>(nRows), 1);

    if (exact) {
        for (int r = 0; r < nRows; ++r) {
            if (lookupExactEqual(keys[static_cast<std::size_t>(r)], lookupVal, /*wildcard=*/true)) {
                return lookupReturnAt(table, r, colIdx - 1, nCols);
            }
        }
        return errNA();
    }

    // Approximate: last key <= lookup (assumes ascending first column).
    int best = lookupApproxLessEqual(keys, lookupVal);
    if (best < 0) return errNA();
    return lookupReturnAt(table, best, colIdx - 1, nCols);
}

XLCellValue Formula::fnHlookup(const std::vector<XLFormulaArg>& args)
{
    // HLOOKUP(lookup_value, table_array, row_index, [range_lookup])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();

    const XLCellValue& lookupVal = args[0][0];
    auto               table     = args[1].materialize();
    int                rowIdx    = static_cast<int>(toDouble(args[2][0]));    // 1-based
    bool               exact     = false;
    if (args.size() >= 4 && !args[3].empty()) {
        if (args[3][0].type() == XLValueType::Boolean)
            exact = !args[3][0].get<bool>();
        else
            exact = (toDouble(args[3][0]) == 0.0);
    }

    if (rowIdx < 1) return errValue();
    int nRows = static_cast<int>(table.rows());
    if (rowIdx > nRows) return errRef();
    int nCols = static_cast<int>(table.cols());

    std::vector<XLCellValue> keyRow;
    keyRow.reserve(static_cast<std::size_t>(nCols));
    for (int c = 0; c < nCols; ++c) keyRow.push_back(table.at(0, static_cast<std::size_t>(c)));
    XLFormulaArg keys(std::move(keyRow), 1, static_cast<std::size_t>(nCols));

    if (exact) {
        for (int c = 0; c < nCols; ++c) {
            if (lookupExactEqual(keys[static_cast<std::size_t>(c)], lookupVal, true)) {
                return lookupReturnAt(table, rowIdx - 1, c, nCols);
            }
        }
        return errNA();
    }

    int best = lookupApproxLessEqual(keys, lookupVal);
    if (best < 0) return errNA();
    return lookupReturnAt(table, rowIdx - 1, best, nCols);
}

XLCellValue Formula::fnXlookup(const std::vector<XLFormulaArg>& args)
{
    // XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
    // match_mode: 0 exact, -1 exact or next smaller, 1 exact or next larger, 2 wildcard
    // search_mode: 1 first-to-last, -1 last-to-first, 2 binary asc, -2 binary desc
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();

    const auto& lookupVal = args[0][0];
    auto        lookupArr = args[1].materialize();
    auto        returnArr = args[2].materialize();

    int matchMode  = (args.size() > 4 && !args[4].empty()) ? static_cast<int>(toDouble(args[4][0])) : 0;
    int searchMode = (args.size() > 5 && !args[5].empty()) ? static_cast<int>(toDouble(args[5][0])) : 1;

    const int n = static_cast<int>(lookupArr.size());
    if (n == 0) {
        if (args.size() > 3 && !args[3].empty()) return args[3][0];
        return errNA();
    }

    int bestMatchIdx = -1;

    auto compareValues = [](const XLCellValue& key, const XLCellValue& val) -> double {
        auto cmp = lookupCompare(key, val);
        if (!cmp.has_value()) return std::numeric_limits<double>::infinity();
        return static_cast<double>(*cmp);
    };

    if (searchMode == 2 || searchMode == -2) {
        // Binary search requires match_mode 0/-1/1 (not wildcard).
        if (matchMode == 2) return errValue();
        int low  = 0;
        int high = n - 1;

        while (low <= high) {
            int         mid  = low + (high - low) / 2;
            const auto& key  = lookupArr[static_cast<std::size_t>(mid)];
            double      diff = compareValues(key, lookupVal);
            if (!std::isfinite(diff)) {
                // Incomparable: skip by shrinking toward high (conservative).
                high = mid - 1;
                continue;
            }

            if (diff == 0.0) {
                bestMatchIdx = mid;
                break;
            }

            if (searchMode == 2) {    // Ascending
                if (diff < 0) {
                    if (matchMode == -1) bestMatchIdx = mid;
                    low = mid + 1;
                }
                else {
                    if (matchMode == 1) bestMatchIdx = mid;
                    high = mid - 1;
                }
            }
            else {    // Descending
                if (diff < 0) {
                    if (matchMode == -1) bestMatchIdx = mid;
                    high = mid - 1;
                }
                else {
                    if (matchMode == 1) bestMatchIdx = mid;
                    low = mid + 1;
                }
            }
        }
    }
    else {
        // Linear search
        int startIdx = (searchMode == -1) ? n - 1 : 0;
        int endIdx   = (searchMode == -1) ? -1 : n;
        int step     = (searchMode == -1) ? -1 : 1;

        if (matchMode == 0 || matchMode == 2) {
            const bool wild = (matchMode == 2);
            for (int i = startIdx; i != endIdx; i += step) {
                if (lookupExactEqual(lookupArr[static_cast<std::size_t>(i)], lookupVal, wild || matchMode == 0)) {
                    // match_mode 0 also allows wildcards in Excel when pattern has *?
                    bestMatchIdx = i;
                    break;
                }
            }
            // For match_mode 0, also allow wildcards if lookup contains them
            if (bestMatchIdx < 0 && matchMode == 0) {
                for (int i = startIdx; i != endIdx; i += step) {
                    if (lookupExactEqual(lookupArr[static_cast<std::size_t>(i)], lookupVal, true)) {
                        bestMatchIdx = i;
                        break;
                    }
                }
            }
        }
        else if (matchMode == -1) {
            // exact or next smaller
            for (int i = startIdx; i != endIdx; i += step) {
                if (lookupExactEqual(lookupArr[static_cast<std::size_t>(i)], lookupVal, false)) {
                    bestMatchIdx = i;
                    break;
                }
            }
            if (bestMatchIdx < 0) bestMatchIdx = lookupApproxLessEqual(lookupArr, lookupVal);
        }
        else if (matchMode == 1) {
            for (int i = startIdx; i != endIdx; i += step) {
                if (lookupExactEqual(lookupArr[static_cast<std::size_t>(i)], lookupVal, false)) {
                    bestMatchIdx = i;
                    break;
                }
            }
            if (bestMatchIdx < 0) bestMatchIdx = lookupApproxGreaterEqual(lookupArr, lookupVal);
        }
    }

    if (bestMatchIdx >= 0) {
        // Return array may be a column (n×1), row (1×n), or multi-column with same row count.
        if (returnArr.rows() == static_cast<std::size_t>(n) && returnArr.cols() >= 1) {
            // Vertical alignment: take row bestMatchIdx, all columns — scalar if 1 col.
            if (returnArr.cols() == 1) return returnArr.at(static_cast<std::size_t>(bestMatchIdx), 0);
            // Multi-column: return first column cell for scalar API (full row via evaluateArray later).
            return returnArr.at(static_cast<std::size_t>(bestMatchIdx), 0);
        }
        if (static_cast<std::size_t>(bestMatchIdx) < returnArr.size()) {
            return returnArr[static_cast<std::size_t>(bestMatchIdx)];
        }
        return errRef();
    }

    if (args.size() > 3 && !args[3].empty()) return args[3][0];
    return errNA();
}

XLCellValue Formula::fnIndex(const std::vector<XLFormulaArg>& args)
{
    // INDEX(array, row_num, [col_num])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const auto& arr = args[0];

    int r = static_cast<int>(toDouble(args[1][0]));
    int c = 1;

    if (args.size() > 2 && !args[2].empty()) { c = static_cast<int>(toDouble(args[2][0])); }
    else if (arr.rows() == 1 && arr.cols() > 1) {
        // Excel spec: if the array is 1D horizontal and only 1 index is provided,
        // it acts as the column index, not the row index.
        c = r;
        r = 1;
    }

    if (r < 1 || c < 1) return errValue();

    std::size_t nCols = arr.cols();
    if (nCols == 0) nCols = 1;    // Fallback protection

    std::size_t idx = static_cast<std::size_t>((r - 1) * nCols + (c - 1));

    if (idx >= arr.size()) return errRef();
    return arr[idx];
}

XLCellValue Formula::fnMatch(const std::vector<XLFormulaArg>& args)
{
    // MATCH(lookup_value, lookup_array, [match_type])
    // match_type: 1 (default) largest <= ; 0 exact (+wildcards); -1 smallest >=
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const XLCellValue& lookupVal = args[0][0];
    auto               arr       = args[1].materialize();
    int                matchType = (args.size() > 2 && !args[2].empty()) ? static_cast<int>(toDouble(args[2][0])) : 1;

    if (matchType == 0) {
        for (std::size_t i = 0; i < arr.size(); ++i) {
            if (lookupExactEqual(arr[i], lookupVal, /*wildcard=*/true)) {
                return XLCellValue(static_cast<int64_t>(i + 1));
            }
        }
        return errNA();
    }
    if (matchType == 1) {
        int best = lookupApproxLessEqual(arr, lookupVal);
        if (best < 0) return errNA();
        return XLCellValue(static_cast<int64_t>(best + 1));
    }
    if (matchType == -1) {
        int best = lookupApproxGreaterEqual(arr, lookupVal);
        if (best < 0) return errNA();
        return XLCellValue(static_cast<int64_t>(best + 1));
    }
    return errValue();
}

namespace
{
    XLCellValue rowFromArg(const XLFormulaArg& arg)
    {
        if (arg.type() == XLFormulaArg::Type::LazyRange) {
            uint32_t r = arg.firstRow();
            if (r > 0) return XLCellValue(static_cast<int64_t>(r));
        }
        const auto& v   = arg[0];
        std::string ref = v.type() == XLValueType::String ? v.get<std::string>() : toString(v);
        if (ref.empty()) return errValue();
        auto bang = ref.find('!');
        if (bang != std::string::npos) ref = ref.substr(bang + 1);
        ref.erase(std::remove(ref.begin(), ref.end(), '$'), ref.end());
        auto colon = ref.find(':');
        if (colon != std::string::npos) ref = ref.substr(0, colon);
        try {
            XLCellReference cellRef(ref);
            return XLCellValue(static_cast<int64_t>(cellRef.row()));
        }
        catch (...) {
            return errValue();
        }
    }

    XLCellValue columnFromArg(const XLFormulaArg& arg)
    {
        if (arg.type() == XLFormulaArg::Type::LazyRange) {
            uint16_t c = arg.firstCol();
            if (c > 0) return XLCellValue(static_cast<int64_t>(c));
        }
        const auto& v   = arg[0];
        std::string ref = v.type() == XLValueType::String ? v.get<std::string>() : toString(v);
        if (ref.empty()) return errValue();
        auto bang = ref.find('!');
        if (bang != std::string::npos) ref = ref.substr(bang + 1);
        ref.erase(std::remove(ref.begin(), ref.end(), '$'), ref.end());
        auto colon = ref.find(':');
        if (colon != std::string::npos) ref = ref.substr(0, colon);
        try {
            XLCellReference cellRef(ref);
            return XLCellValue(static_cast<int64_t>(cellRef.column()));
        }
        catch (...) {
            return errValue();
        }
    }
}    // namespace

XLCellValue Formula::fnRow(const std::vector<XLFormulaArg>& args)
{
    // ROW([reference]) — return the 1-based row of the top-left cell of the argument
    if (args.empty() || args[0].empty()) return errValue();
    return rowFromArg(args[0]);
}

/** Session-aware ROW: supports parameterless ROW() via XLEvalSession::currentCell. */
static XLCellValue fnRowSession(const std::vector<XLFormulaArg>& args, XLEvalSession& session)
{
    if (args.empty() || args[0].empty()) {
        if (session.hasCurrentCell()) return XLCellValue(static_cast<int64_t>(session.currentRow()));
        return errValue();
    }
    return rowFromArg(args[0]);
}

XLCellValue Formula::fnColumn(const std::vector<XLFormulaArg>& args)
{
    // COLUMN([reference]) — return the 1-based column of the top-left cell
    if (args.empty() || args[0].empty()) return errValue();
    return columnFromArg(args[0]);
}

static XLCellValue fnColumnSession(const std::vector<XLFormulaArg>& args, XLEvalSession& session)
{
    if (args.empty() || args[0].empty()) {
        if (session.hasCurrentCell()) return XLCellValue(static_cast<int64_t>(session.currentCol()));
        return errValue();
    }
    return columnFromArg(args[0]);
}

/** Scalar wrappers that delegate to the ref-function path via session.expandRange. */
static XLCellValue fnIndirectSession(const std::vector<XLFormulaArg>& args, XLEvalSession& session)
{
    // Scalar context is handled by evalRefFunction when used as FuncCall; this path is a
    // safety net if registered as a normal function.
    if (args.empty() || args[0].empty()) return errRef();
    if (isError(args[0][0])) return args[0][0];
    std::string refText = toString(args[0][0]);
    if (!refText.empty() && refText.front() == '=') refText.erase(refText.begin());
    if (auto named = session.resolveName(refText)) {
        refText = *named;
        if (!refText.empty() && refText.front() == '=') refText.erase(refText.begin());
    }
    auto arg = session.expandRange(refText);
    return arg.empty() ? errRef() : arg[0];
}

static XLCellValue fnOffsetSession(const std::vector<XLFormulaArg>& args, XLEvalSession& session)
{
    if (args.size() < 3 || args[0].type() != XLFormulaArg::Type::LazyRange) return errValue();
    if (args[1].empty() || args[2].empty() || !isNumeric(args[1][0]) || !isNumeric(args[2][0])) return errValue();

    const int64_t rowOff = static_cast<int64_t>(std::trunc(toDouble(args[1][0])));
    const int64_t colOff = static_cast<int64_t>(std::trunc(toDouble(args[2][0])));
    int64_t       height = static_cast<int64_t>(args[0].rows());
    int64_t       width  = static_cast<int64_t>(args[0].cols());
    if (args.size() >= 4 && !args[3].empty()) {
        if (!isNumeric(args[3][0])) return errValue();
        height = static_cast<int64_t>(std::trunc(toDouble(args[3][0])));
    }
    if (args.size() >= 5 && !args[4].empty()) {
        if (!isNumeric(args[4][0])) return errValue();
        width = static_cast<int64_t>(std::trunc(toDouble(args[4][0])));
    }
    if (height == 0 || width == 0) return errRef();
    if (height < 0) height = -height;
    if (width < 0) width = -width;

    const int64_t r1 = static_cast<int64_t>(args[0].firstRow()) + rowOff;
    const int64_t c1 = static_cast<int64_t>(args[0].firstCol()) + colOff;
    if (r1 < 1 || c1 < 1) return errRef();

    XLFormulaArg range(static_cast<uint32_t>(r1),
                       static_cast<uint32_t>(r1 + height - 1),
                       static_cast<uint16_t>(c1),
                       static_cast<uint16_t>(c1 + width - 1),
                       args[0].sheetName(),
                       session.resolverPtr());
    return range.empty() ? errRef() : range[0];
}

XLCellValue Formula::fnDatedif(const std::vector<XLFormulaArg>& args)
{
    // DATEDIF(start_date, end_date, unit)
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    if (!isNumeric(args[0][0]) || !isNumeric(args[1][0])) return errValue();

    double d1Serial = toDouble(args[0][0]);
    double d2Serial = toDouble(args[1][0]);
    if (d1Serial > d2Serial) return errNum();  // start must be <= end

    auto tm1 = info_serialToTm(d1Serial);
    auto tm2 = info_serialToTm(d2Serial);
    int  y1  = tm1.tm_year + 1900, m1 = tm1.tm_mon + 1, day1 = tm1.tm_mday;
    int  y2  = tm2.tm_year + 1900, m2 = tm2.tm_mon + 1, day2 = tm2.tm_mday;

    std::string unit = toString(args[2][0]);
    // Normalize to upper
    for (char& c : unit) c = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));

    if (unit == "Y") {
        // Complete years
        int yrs = y2 - y1;
        if (m2 * 100 + day2 < m1 * 100 + day1) --yrs;
        return XLCellValue(static_cast<int64_t>(yrs));
    }
    if (unit == "M") {
        // Complete months
        int months = (y2 - y1) * 12 + (m2 - m1);
        if (day2 < day1) --months;
        return XLCellValue(static_cast<int64_t>(months));
    }
    if (unit == "D") {
        // Total days
        return XLCellValue(static_cast<int64_t>(std::floor(d2Serial) - std::floor(d1Serial)));
    }
    if (unit == "MD") {
        // Days ignoring months and years
        int days = day2 - day1;
        if (days < 0) days += info_daysInMonth(y2, m2 == 1 ? 12 : m2 - 1);
        return XLCellValue(static_cast<int64_t>(days));
    }
    if (unit == "YM") {
        // Complete months ignoring years
        int months = (m2 - m1 + 12) % 12;
        if (day2 < day1 && months > 0) --months;
        else if (day2 < day1 && months == 0) months = 11;
        return XLCellValue(static_cast<int64_t>(months));
    }
    if (unit == "YD") {
        // Days ignoring years (between same-year anniversary dates)
        std::tm tmAnn{}; tmAnn.tm_year = tm2.tm_year; tmAnn.tm_mon = tm1.tm_mon; tmAnn.tm_mday = tm1.tm_mday;
        double annSerial = info_tmToSerial(tmAnn);
        if (annSerial > d2Serial) {
            tmAnn.tm_year -= 1;
            annSerial = info_tmToSerial(tmAnn);
        }
        return XLCellValue(static_cast<int64_t>(std::floor(d2Serial) - std::floor(annSerial)));
    }
    return errValue();
}

XLCellValue Formula::fnIseven(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double val = toDouble(args[0][0]);
    if (std::isnan(val)) return errValue();
    int64_t i = static_cast<int64_t>(std::trunc(val));
    return XLCellValue(i % 2 == 0);
}

XLCellValue Formula::fnIsodd(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double val = toDouble(args[0][0]);
    if (std::isnan(val)) return errValue();
    int64_t i = static_cast<int64_t>(std::trunc(val));
    return XLCellValue(i % 2 != 0);
}

void OpenXLSX::registerInfoFunctions(XLFormulaRegistry& registry)
{
    registry.registerFunction("ISNUMBER", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIsnumber));
    registry.registerFunction("ISBLANK", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIsblank));
    registry.registerFunction("ISERROR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIserror));
    registry.registerFunction("ISTEXT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIstext));
    registry.registerFunction("ISNA", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIsna));
    registry.registerFunction("IFNA", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIfna));
    registry.registerFunction("ISLOGICAL", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIslogical));
    registry.registerFunction("ISNONTEXT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIsnontext));
    registry.registerFunction("ISERR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIserr));
    registry.registerFunction("VLOOKUP", std::make_shared<XLSimpleFormulaFunction>(Formula::fnVlookup));
    registry.registerFunction("HLOOKUP", std::make_shared<XLSimpleFormulaFunction>(Formula::fnHlookup));
    registry.registerFunction("XLOOKUP", std::make_shared<XLSimpleFormulaFunction>(Formula::fnXlookup));
    registry.registerFunction("INDEX", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIndex));
    registry.registerFunction("MATCH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMatch));
    registry.registerFunction("ROW", std::make_shared<XLSessionFormulaFunction>(fnRowSession));
    registry.registerFunction("COLUMN", std::make_shared<XLSessionFormulaFunction>(fnColumnSession));
    registry.registerFunction("INDIRECT", std::make_shared<XLSessionFormulaFunction>(fnIndirectSession));
    registry.registerFunction("OFFSET", std::make_shared<XLSessionFormulaFunction>(fnOffsetSession));
    registry.registerFunction("DATEDIF", std::make_shared<XLSimpleFormulaFunction>(Formula::fnDatedif));
    registry.registerFunction("ISEVEN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIseven));
    registry.registerFunction("ISODD", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIsodd));
}

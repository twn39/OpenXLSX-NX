/**
 * @file XLFormulaArray.cpp
 * @brief Phase B dynamic-array functions: FILTER, UNIQUE, SORT, SEQUENCE, TRANSPOSE.
 */

#include "XLFormulaFunctionsInternal.hpp"
#include "XLFormulaRegistry.hpp"
#include "XLFormulaUtils.hpp"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <random>
#include <string>
#include <unordered_map>
#include <vector>

using namespace OpenXLSX;

namespace
{
    bool truthy(const XLCellValue& v)
    {
        if (v.type() == XLValueType::Empty) return false;
        if (v.type() == XLValueType::Boolean) return v.get<bool>();
        if (v.type() == XLValueType::Error) return false;
        if (isNumeric(v)) return toDouble(v) != 0.0;
        // Non-empty text is true (Excel-ish for include masks built from text is rare).
        return !toString(v).empty();
    }

    bool valuesEqual(const XLCellValue& a, const XLCellValue& b)
    {
        if (a.type() == XLValueType::Error || b.type() == XLValueType::Error) {
            return isError(a) && isError(b) && toString(a) == toString(b);
        }
        if (isNumeric(a) && isNumeric(b)) return toDouble(a) == toDouble(b);
        if (a.type() == XLValueType::Boolean && b.type() == XLValueType::Boolean) return a.get<bool>() == b.get<bool>();
        return toString(a) == toString(b);
    }

    int compareValues(const XLCellValue& a, const XLCellValue& b)
    {
        // Excel SORT: numbers before text; blanks last-ish — keep simple numeric/string order.
        if (isNumeric(a) && isNumeric(b)) {
            double da = toDouble(a), db = toDouble(b);
            if (da < db) return -1;
            if (da > db) return 1;
            return 0;
        }
        if (isNumeric(a) && !isNumeric(b)) return -1;
        if (!isNumeric(a) && isNumeric(b)) return 1;
        const std::string sa = toString(a);
        const std::string sb = toString(b);
        if (sa < sb) return -1;
        if (sa > sb) return 1;
        return 0;
    }

    XLFormulaArg makeErrorArg(std::string_view token)
    {
        XLCellValue e;
        e.setError(std::string(token));
        return XLFormulaArg(std::move(e));
    }

    // =========================================================================
    // FILTER(array, include, [if_empty])
    // =========================================================================
    XLFormulaArg fnFilter(const std::vector<XLFormulaArg>& args, XLEvalSession& /*session*/)
    {
        if (args.size() < 2 || args[0].empty() || args[1].empty()) return makeErrorArg("#VALUE!");

        auto array   = args[0].materialize();
        auto include = args[1].materialize();

        const size_t nRows = array.rows();
        const size_t nCols = array.cols();
        if (nRows == 0 || nCols == 0) return makeErrorArg("#CALC!");

        // Row filter when include is a column (nRows×1) or a flat vector of length nRows.
        // Column filter when include is a row (1×nCols).
        const bool byRow = (include.rows() == nRows && include.cols() == 1) ||
                           (include.size() == nRows && include.cols() <= 1) ||
                           (include.rows() == nRows && include.size() == nRows);
        const bool byCol = (include.rows() == 1 && include.cols() == nCols) ||
                           (include.size() == nCols && include.rows() == 1);

        if (!byRow && !byCol) {
            // Allow flat include matching row count as row-filter default.
            if (include.size() == nRows)
                ;    // treat as byRow
            else
                return makeErrorArg("#VALUE!");
        }

        const bool filterRows = byRow || (!byCol && include.size() == nRows);

        std::vector<size_t> keep;
        if (filterRows) {
            keep.reserve(nRows);
            for (size_t r = 0; r < nRows; ++r) {
                XLCellValue mask = (include.rows() >= 1 && include.cols() >= 1) ? include.at(r, 0) : include[r];
                if (include.rows() == nRows && include.cols() > 1) mask = include.at(r, 0);
                if (truthy(mask)) keep.push_back(r);
            }
        }
        else {
            keep.reserve(nCols);
            for (size_t c = 0; c < nCols; ++c) {
                XLCellValue mask = include.at(0, c);
                if (truthy(mask)) keep.push_back(c);
            }
        }

        if (keep.empty()) {
            if (args.size() >= 3 && !args[2].empty()) {
                auto fallback = args[2].materialize();
                return fallback;
            }
            return makeErrorArg("#CALC!");
        }

        if (filterRows) {
            std::vector<XLCellValue> out;
            out.reserve(keep.size() * nCols);
            for (size_t r : keep) {
                for (size_t c = 0; c < nCols; ++c) out.push_back(array.at(r, c));
            }
            return XLFormulaArg(std::move(out), keep.size(), nCols);
        }

        std::vector<XLCellValue> out;
        out.reserve(nRows * keep.size());
        for (size_t r = 0; r < nRows; ++r) {
            for (size_t c : keep) out.push_back(array.at(r, c));
        }
        return XLFormulaArg(std::move(out), nRows, keep.size());
    }

    // =========================================================================
    // UNIQUE(array, [by_col], [exactly_once])
    // =========================================================================
    XLFormulaArg fnUnique(const std::vector<XLFormulaArg>& args, XLEvalSession& /*session*/)
    {
        if (args.empty() || args[0].empty()) return makeErrorArg("#VALUE!");
        auto array = args[0].materialize();
        const size_t nRows = array.rows();
        const size_t nCols = array.cols();
        if (nRows == 0 || nCols == 0) return makeErrorArg("#VALUE!");

        bool byCol = false;
        if (args.size() >= 2 && !args[1].empty()) byCol = truthy(args[1][0]);
        bool exactlyOnce = false;
        if (args.size() >= 3 && !args[2].empty()) exactlyOnce = truthy(args[2][0]);

        auto rowKey = [&](size_t r) {
            std::string k;
            for (size_t c = 0; c < nCols; ++c) {
                k.push_back('\x1f');
                k += toString(array.at(r, c));
            }
            return k;
        };
        auto colKey = [&](size_t c) {
            std::string k;
            for (size_t r = 0; r < nRows; ++r) {
                k.push_back('\x1f');
                k += toString(array.at(r, c));
            }
            return k;
        };

        if (!byCol) {
            std::vector<size_t> order;
            std::unordered_map<std::string, size_t> first;
            std::unordered_map<std::string, size_t> counts;
            order.reserve(nRows);
            for (size_t r = 0; r < nRows; ++r) {
                auto k = rowKey(r);
                counts[k]++;
                if (first.find(k) == first.end()) {
                    first[k] = r;
                    order.push_back(r);
                }
            }
            std::vector<size_t> keep;
            for (size_t r : order) {
                auto k = rowKey(r);
                if (!exactlyOnce || counts[k] == 1) keep.push_back(r);
            }
            if (keep.empty()) return makeErrorArg("#CALC!");
            std::vector<XLCellValue> out;
            out.reserve(keep.size() * nCols);
            for (size_t r : keep)
                for (size_t c = 0; c < nCols; ++c) out.push_back(array.at(r, c));
            return XLFormulaArg(std::move(out), keep.size(), nCols);
        }

        std::vector<size_t> order;
        std::unordered_map<std::string, size_t> first;
        std::unordered_map<std::string, size_t> counts;
        for (size_t c = 0; c < nCols; ++c) {
            auto k = colKey(c);
            counts[k]++;
            if (first.find(k) == first.end()) {
                first[k] = c;
                order.push_back(c);
            }
        }
        std::vector<size_t> keep;
        for (size_t c : order) {
            auto k = colKey(c);
            if (!exactlyOnce || counts[k] == 1) keep.push_back(c);
        }
        if (keep.empty()) return makeErrorArg("#CALC!");
        std::vector<XLCellValue> out;
        out.reserve(nRows * keep.size());
        for (size_t r = 0; r < nRows; ++r)
            for (size_t c : keep) out.push_back(array.at(r, c));
        return XLFormulaArg(std::move(out), nRows, keep.size());
    }

    // =========================================================================
    // SORT(array, [sort_index], [sort_order], [by_col])
    // sort_index: 1-based column (or row if by_col) to sort on
    // sort_order: 1 ascending (default), -1 descending
    // =========================================================================
    XLFormulaArg fnSort(const std::vector<XLFormulaArg>& args, XLEvalSession& /*session*/)
    {
        if (args.empty() || args[0].empty()) return makeErrorArg("#VALUE!");
        auto array = args[0].materialize();
        const size_t nRows = array.rows();
        const size_t nCols = array.cols();
        if (nRows == 0 || nCols == 0) return makeErrorArg("#VALUE!");

        int  sortIndex = 1;
        int  sortOrder = 1;
        bool byCol     = false;
        if (args.size() >= 2 && !args[1].empty() && isNumeric(args[1][0])) sortIndex = static_cast<int>(toDouble(args[1][0]));
        if (args.size() >= 3 && !args[2].empty() && isNumeric(args[2][0])) sortOrder = static_cast<int>(toDouble(args[2][0]));
        if (args.size() >= 4 && !args[3].empty()) byCol = truthy(args[3][0]);
        if (sortIndex < 1) return makeErrorArg("#VALUE!");

        if (!byCol) {
            if (static_cast<size_t>(sortIndex) > nCols) return makeErrorArg("#VALUE!");
            std::vector<size_t> idx(nRows);
            for (size_t i = 0; i < nRows; ++i) idx[i] = i;
            const size_t keyCol = static_cast<size_t>(sortIndex - 1);
            std::stable_sort(idx.begin(), idx.end(), [&](size_t a, size_t b) {
                int cmp = compareValues(array.at(a, keyCol), array.at(b, keyCol));
                return sortOrder < 0 ? (cmp > 0) : (cmp < 0);
            });
            std::vector<XLCellValue> out;
            out.reserve(nRows * nCols);
            for (size_t r : idx)
                for (size_t c = 0; c < nCols; ++c) out.push_back(array.at(r, c));
            return XLFormulaArg(std::move(out), nRows, nCols);
        }

        if (static_cast<size_t>(sortIndex) > nRows) return makeErrorArg("#VALUE!");
        std::vector<size_t> idx(nCols);
        for (size_t i = 0; i < nCols; ++i) idx[i] = i;
        const size_t keyRow = static_cast<size_t>(sortIndex - 1);
        std::stable_sort(idx.begin(), idx.end(), [&](size_t a, size_t b) {
            int cmp = compareValues(array.at(keyRow, a), array.at(keyRow, b));
            return sortOrder < 0 ? (cmp > 0) : (cmp < 0);
        });
        std::vector<XLCellValue> out;
        out.reserve(nRows * nCols);
        for (size_t r = 0; r < nRows; ++r)
            for (size_t c : idx) out.push_back(array.at(r, c));
        return XLFormulaArg(std::move(out), nRows, nCols);
    }

    // =========================================================================
    // SEQUENCE(rows, [columns], [start], [step])
    // =========================================================================
    XLFormulaArg fnSequence(const std::vector<XLFormulaArg>& args, XLEvalSession& /*session*/)
    {
        if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return makeErrorArg("#VALUE!");
        int64_t nRows = static_cast<int64_t>(std::trunc(toDouble(args[0][0])));
        int64_t nCols = 1;
        double  start = 1.0;
        double  step  = 1.0;
        if (args.size() >= 2 && !args[1].empty()) {
            if (!isNumeric(args[1][0])) return makeErrorArg("#VALUE!");
            nCols = static_cast<int64_t>(std::trunc(toDouble(args[1][0])));
        }
        if (args.size() >= 3 && !args[2].empty()) {
            if (!isNumeric(args[2][0])) return makeErrorArg("#VALUE!");
            start = toDouble(args[2][0]);
        }
        if (args.size() >= 4 && !args[3].empty()) {
            if (!isNumeric(args[3][0])) return makeErrorArg("#VALUE!");
            step = toDouble(args[3][0]);
        }
        if (nRows < 1 || nCols < 1) return makeErrorArg("#VALUE!");
        // Safety cap to avoid accidental huge allocations in evaluate.
        if (nRows > 1048576 || nCols > 16384 || nRows * nCols > 10'000'000) return makeErrorArg("#NUM!");

        std::vector<XLCellValue> out;
        out.reserve(static_cast<size_t>(nRows * nCols));
        double v = start;
        for (int64_t r = 0; r < nRows; ++r) {
            for (int64_t c = 0; c < nCols; ++c) {
                // Prefer integer when exact.
                if (std::floor(v) == v && std::abs(v) < static_cast<double>(INT64_MAX))
                    out.emplace_back(static_cast<int64_t>(v));
                else
                    out.emplace_back(v);
                v += step;
            }
        }
        return XLFormulaArg(std::move(out), static_cast<size_t>(nRows), static_cast<size_t>(nCols));
    }

    // =========================================================================
    // TRANSPOSE(array)
    // =========================================================================
    XLFormulaArg fnTranspose(const std::vector<XLFormulaArg>& args, XLEvalSession& /*session*/)
    {
        if (args.empty() || args[0].empty()) return makeErrorArg("#VALUE!");
        auto array = args[0].materialize();
        const size_t nRows = array.rows();
        const size_t nCols = array.cols();
        if (nRows == 0 || nCols == 0) return makeErrorArg("#VALUE!");

        std::vector<XLCellValue> out;
        out.reserve(nRows * nCols);
        for (size_t c = 0; c < nCols; ++c)
            for (size_t r = 0; r < nRows; ++r) out.push_back(array.at(r, c));
        return XLFormulaArg(std::move(out), nCols, nRows);
    }

    // =========================================================================
    // SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], ...)
    // =========================================================================
    XLFormulaArg fnSortBy(const std::vector<XLFormulaArg>& args, XLEvalSession& /*session*/)
    {
        if (args.size() < 2 || args[0].empty() || args[1].empty()) return makeErrorArg("#VALUE!");
        auto array = args[0].materialize();
        const size_t nRows = array.rows();
        const size_t nCols = array.cols();
        if (nRows == 0 || nCols == 0) return makeErrorArg("#VALUE!");

        struct KeySpec
        {
            XLFormulaArg by;
            int          order{1};    // 1 asc, -1 desc
        };
        std::vector<KeySpec> keys;
        // args: array, by1, [order1], by2, [order2], ...
        size_t i = 1;
        while (i < args.size()) {
            if (args[i].empty()) return makeErrorArg("#VALUE!");
            KeySpec ks;
            ks.by = args[i].materialize();
            ++i;
            if (i < args.size() && !args[i].empty() && isNumeric(args[i][0])) {
                // Ambiguous: could be next by-array (range) or sort order (scalar).
                // Excel: sort_order is -1 or 1. If scalar numeric ±1 (or any number), treat as order.
                // If multi-cell, treat as next by-array.
                if (args[i].size() == 1) {
                    ks.order = static_cast<int>(toDouble(args[i][0]));
                    if (ks.order != 1 && ks.order != -1) ks.order = (ks.order < 0) ? -1 : 1;
                    ++i;
                }
            }
            // by array must align with rows of array (sort by row).
            if (ks.by.size() != nRows && !(ks.by.rows() == nRows && ks.by.cols() >= 1)) {
                // Allow column vector of nRows, or flat size nRows.
                if (ks.by.size() != nRows) return makeErrorArg("#VALUE!");
            }
            keys.push_back(std::move(ks));
        }
        if (keys.empty()) return makeErrorArg("#VALUE!");

        std::vector<size_t> idx(nRows);
        for (size_t r = 0; r < nRows; ++r) idx[r] = r;
        std::stable_sort(idx.begin(), idx.end(), [&](size_t a, size_t b) {
            for (const auto& ks : keys) {
                XLCellValue va = (ks.by.cols() >= 1 && ks.by.rows() >= 1) ? ks.by.at(a, 0) : ks.by[a];
                XLCellValue vb = (ks.by.cols() >= 1 && ks.by.rows() >= 1) ? ks.by.at(b, 0) : ks.by[b];
                // If by is flat size nRows with cols==1 rows==nRows, at works.
                if (ks.by.rows() == 1 && ks.by.cols() == nRows) {
                    va = ks.by.at(0, a);
                    vb = ks.by.at(0, b);
                }
                int cmp = compareValues(va, vb);
                if (cmp != 0) return ks.order < 0 ? (cmp > 0) : (cmp < 0);
            }
            return false;
        });

        std::vector<XLCellValue> out;
        out.reserve(nRows * nCols);
        for (size_t r : idx)
            for (size_t c = 0; c < nCols; ++c) out.push_back(array.at(r, c));
        return XLFormulaArg(std::move(out), nRows, nCols);
    }

    // =========================================================================
    // RANDARRAY([rows],[columns],[min],[max],[integer])
    // =========================================================================
    XLFormulaArg fnRandArray(const std::vector<XLFormulaArg>& args, XLEvalSession& /*session*/)
    {
        int64_t nRows = 1, nCols = 1;
        double  minV = 0.0, maxV = 1.0;
        bool    asInt = false;
        if (args.size() >= 1 && !args[0].empty()) {
            if (!isNumeric(args[0][0])) return makeErrorArg("#VALUE!");
            nRows = static_cast<int64_t>(std::trunc(toDouble(args[0][0])));
        }
        if (args.size() >= 2 && !args[1].empty()) {
            if (!isNumeric(args[1][0])) return makeErrorArg("#VALUE!");
            nCols = static_cast<int64_t>(std::trunc(toDouble(args[1][0])));
        }
        if (args.size() >= 3 && !args[2].empty()) {
            if (!isNumeric(args[2][0])) return makeErrorArg("#VALUE!");
            minV = toDouble(args[2][0]);
        }
        if (args.size() >= 4 && !args[3].empty()) {
            if (!isNumeric(args[3][0])) return makeErrorArg("#VALUE!");
            maxV = toDouble(args[3][0]);
        }
        if (args.size() >= 5 && !args[4].empty()) asInt = truthy(args[4][0]);
        if (nRows < 1 || nCols < 1) return makeErrorArg("#VALUE!");
        if (nRows * nCols > 10'000'000) return makeErrorArg("#NUM!");
        if (minV > maxV) std::swap(minV, maxV);

        auto& rng = formulaRng();
        std::vector<XLCellValue> out;
        out.reserve(static_cast<size_t>(nRows * nCols));
        if (asInt) {
            int64_t lo = static_cast<int64_t>(std::ceil(minV));
            int64_t hi = static_cast<int64_t>(std::floor(maxV));
            if (lo > hi) return makeErrorArg("#VALUE!");
            std::uniform_int_distribution<int64_t> dist(lo, hi);
            for (int64_t i = 0; i < nRows * nCols; ++i) out.emplace_back(dist(rng));
        }
        else {
            std::uniform_real_distribution<double> dist(minV, maxV);
            for (int64_t i = 0; i < nRows * nCols; ++i) out.emplace_back(dist(rng));
        }
        return XLFormulaArg(std::move(out), static_cast<size_t>(nRows), static_cast<size_t>(nCols));
    }

    // =========================================================================
    // TOCOL(array, [ignore], [scan_by_col])
    // ignore: 0 none, 1 blanks, 2 errors, 3 blanks+errors
    // =========================================================================
    XLFormulaArg fnTocol(const std::vector<XLFormulaArg>& args, XLEvalSession& /*session*/)
    {
        if (args.empty() || args[0].empty()) return makeErrorArg("#VALUE!");
        auto array = args[0].materialize();
        int  ignore = 0;
        bool byCol  = false;
        if (args.size() >= 2 && !args[1].empty() && isNumeric(args[1][0])) ignore = static_cast<int>(toDouble(args[1][0]));
        if (args.size() >= 3 && !args[2].empty()) byCol = truthy(args[2][0]);

        auto skip = [ignore](const XLCellValue& v) {
            if (ignore == 1 || ignore == 3) {
                if (isEmpty(v) || (v.type() == XLValueType::String && v.get<std::string>().empty())) return true;
            }
            if (ignore == 2 || ignore == 3) {
                if (isError(v)) return true;
            }
            return false;
        };

        std::vector<XLCellValue> out;
        const size_t nRows = array.rows();
        const size_t nCols = array.cols();
        if (!byCol) {
            for (size_t r = 0; r < nRows; ++r)
                for (size_t c = 0; c < nCols; ++c) {
                    auto v = array.at(r, c);
                    if (!skip(v)) out.push_back(std::move(v));
                }
        }
        else {
            for (size_t c = 0; c < nCols; ++c)
                for (size_t r = 0; r < nRows; ++r) {
                    auto v = array.at(r, c);
                    if (!skip(v)) out.push_back(std::move(v));
                }
        }
        if (out.empty()) return makeErrorArg("#CALC!");
        return XLFormulaArg(std::move(out), out.size(), 1);
    }

    XLFormulaArg fnTorow(const std::vector<XLFormulaArg>& args, XLEvalSession& session)
    {
        auto col = fnTocol(args, session);
        if (col.type() == XLFormulaArg::Type::Empty) return col;
        if (isError(col.asScalar()) && col.size() <= 1) return col;
        // Convert flat values to 1×n
        auto                     mat = col.materialize();
        std::vector<XLCellValue> out;
        out.reserve(mat.size());
        for (size_t i = 0; i < mat.size(); ++i) out.push_back(mat[i]);
        if (out.empty()) return makeErrorArg("#CALC!");
        const size_t n = out.size();    // capture before move (argument evaluation order)
        return XLFormulaArg(std::move(out), /*rows=*/1, /*cols=*/n);
    }

    // =========================================================================
    // HSTACK / VSTACK
    // =========================================================================
    XLFormulaArg fnHstack(const std::vector<XLFormulaArg>& args, XLEvalSession& /*session*/)
    {
        if (args.empty()) return makeErrorArg("#VALUE!");
        std::vector<XLFormulaArg> mats;
        size_t                    rows = 0;
        size_t                    cols = 0;
        for (const auto& a : args) {
            if (a.empty() && a.type() != XLFormulaArg::Type::Array) continue;
            auto m = a.materialize();
            if (m.rows() == 0 || m.cols() == 0) {
                // scalar
                if (m.size() == 1 || m.type() == XLFormulaArg::Type::Scalar) {
                    std::vector<XLCellValue> one{m.asScalar()};
                    m = XLFormulaArg(std::move(one), 1, 1);
                }
                else
                    continue;
            }
            if (rows == 0)
                rows = m.rows();
            else if (m.rows() != rows)
                return makeErrorArg("#VALUE!");    // Excel pads in some cases; we require equal rows
            cols += m.cols();
            mats.push_back(std::move(m));
        }
        if (mats.empty() || rows == 0 || cols == 0) return makeErrorArg("#VALUE!");
        std::vector<XLCellValue> out(rows * cols);
        size_t                   colOff = 0;
        for (const auto& m : mats) {
            for (size_t r = 0; r < rows; ++r)
                for (size_t c = 0; c < m.cols(); ++c) out[r * cols + colOff + c] = m.at(r, c);
            colOff += m.cols();
        }
        return XLFormulaArg(std::move(out), rows, cols);
    }

    XLFormulaArg fnVstack(const std::vector<XLFormulaArg>& args, XLEvalSession& /*session*/)
    {
        if (args.empty()) return makeErrorArg("#VALUE!");
        std::vector<XLFormulaArg> mats;
        size_t                    rows = 0;
        size_t                    cols = 0;
        for (const auto& a : args) {
            if (a.empty() && a.type() != XLFormulaArg::Type::Array) continue;
            auto m = a.materialize();
            if (m.rows() == 0 || m.cols() == 0) {
                if (m.size() == 1 || m.type() == XLFormulaArg::Type::Scalar) {
                    std::vector<XLCellValue> one{m.asScalar()};
                    m = XLFormulaArg(std::move(one), 1, 1);
                }
                else
                    continue;
            }
            if (cols == 0)
                cols = m.cols();
            else if (m.cols() != cols)
                return makeErrorArg("#VALUE!");
            rows += m.rows();
            mats.push_back(std::move(m));
        }
        if (mats.empty() || rows == 0 || cols == 0) return makeErrorArg("#VALUE!");
        std::vector<XLCellValue> out;
        out.reserve(rows * cols);
        for (const auto& m : mats)
            for (size_t r = 0; r < m.rows(); ++r)
                for (size_t c = 0; c < cols; ++c) out.push_back(m.at(r, c));
        return XLFormulaArg(std::move(out), rows, cols);
    }
}    // namespace

void OpenXLSX::registerArrayFunctions(XLFormulaRegistry& registry)
{
    registry.registerFunction("FILTER", std::make_shared<XLArrayFormulaFunction>(fnFilter));
    registry.registerFunction("UNIQUE", std::make_shared<XLArrayFormulaFunction>(fnUnique));
    registry.registerFunction("SORT", std::make_shared<XLArrayFormulaFunction>(fnSort));
    registry.registerFunction("SORTBY", std::make_shared<XLArrayFormulaFunction>(fnSortBy));
    registry.registerFunction("SEQUENCE", std::make_shared<XLArrayFormulaFunction>(fnSequence));
    registry.registerFunction("RANDARRAY", std::make_shared<XLArrayFormulaFunction>(fnRandArray));
    registry.registerFunction("TRANSPOSE", std::make_shared<XLArrayFormulaFunction>(fnTranspose));
    registry.registerFunction("TOCOL", std::make_shared<XLArrayFormulaFunction>(fnTocol));
    registry.registerFunction("TOROW", std::make_shared<XLArrayFormulaFunction>(fnTorow));
    registry.registerFunction("HSTACK", std::make_shared<XLArrayFormulaFunction>(fnHstack));
    registry.registerFunction("VSTACK", std::make_shared<XLArrayFormulaFunction>(fnVstack));
}

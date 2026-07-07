#include "XLFormulaFunctionsInternal.hpp"
#include "XLFormulaRegistry.hpp"
#include "XLDateTime.hpp"
#include "XLNumberFormatter.hpp"
#include "XLFormulaUtils.hpp"
#include <algorithm>
#include <array>
#include <cmath>
#include <numeric>
#include <random>
#include <regex>
#include <unordered_set>
#include <cctype>

using namespace OpenXLSX;

namespace
{
    bool stat_matchesCriteria(const XLCellValue& cell, const std::string& criteria)
    {
        if (criteria.empty()) return false;
        std::string_view op;
        std::string_view rhs;
        auto s = std::string_view(criteria);
        if (s.size() >= 2 && s[0] == '<' && s[1] == '>') {
            op = "<>";
            rhs = s.substr(2);
        } else if (s.size() >= 2 && s[0] == '<' && s[1] == '=') {
            op = "<=";
            rhs = s.substr(2);
        } else if (s.size() >= 2 && s[0] == '>' && s[1] == '=') {
            op = ">=";
            rhs = s.substr(2);
        } else if (s[0] == '<') {
            op = "<";
            rhs = s.substr(1);
        } else if (s[0] == '>') {
            op = ">";
            rhs = s.substr(1);
        } else if (s[0] == '=') {
            op = "=";
            rhs = s.substr(1);
        } else {
            op = "=";
            rhs = s;
        }
        double rhsNum = 0.0;
        bool rhsIsNum = false;
        try {
            std::size_t idx = 0;
            rhsNum = std::stod(std::string(rhs), &idx);
            rhsIsNum = (idx == rhs.size());
        } catch (...) {}
        if (rhsIsNum && isNumeric(cell)) {
            double lhsNum = toDouble(cell);
            if (op == "=") return lhsNum == rhsNum;
            if (op == "<>") return lhsNum != rhsNum;
            if (op == "<") return lhsNum < rhsNum;
            if (op == "<=") return lhsNum <= rhsNum;
            if (op == ">") return lhsNum > rhsNum;
            if (op == ">=") return lhsNum >= rhsNum;
        }
        std::string cellStr = toString(cell);
        std::string rhsStr = std::string(rhs);
        std::string cellLo = cellStr, rhsLo = rhsStr;
        std::transform(cellLo.begin(), cellLo.end(), cellLo.begin(), ::tolower);
        std::transform(rhsLo.begin(), rhsLo.end(), rhsLo.begin(), ::tolower);
        bool hasWildcard = rhsLo.find('*') != std::string::npos || rhsLo.find('?') != std::string::npos;
        auto wildcardMatch = [](const std::string& text, const std::string& pattern) -> bool {
            std::size_t t = 0, p = 0;
            std::size_t starT = std::string::npos, starP = std::string::npos;
            while (t < text.size()) {
                if (p < pattern.size() && (pattern[p] == '?' || pattern[p] == text[t])) {
                    ++t; ++p;
                } else if (p < pattern.size() && pattern[p] == '*') {
                    starP = p++; starT = t;
                } else if (starP != std::string::npos) {
                    p = starP + 1; t = ++starT;
                } else return false;
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
} // namespace

namespace
{
    constexpr double stat_kPi = 3.14159265358979323846;
    constexpr double stat_kSqrt2Pi = 2.50662827463100050242;

    static double stat_xlMathNormSInv(double p)
    {
        if (p <= 0.0 || p >= 1.0) return std::numeric_limits<double>::quiet_NaN();
        constexpr double a[] = {-3.969683028665376e+01,  2.209460984245205e+02,
                                 -2.759285104469687e+02,  1.383577518672690e+02,
                                 -3.066479806614716e+01,  2.506628277459239e+00};
        constexpr double b[] = {-5.447609879822406e+01,  1.615858368580409e+02,
                                 -1.556989798598866e+02,  6.680131188771972e+01,
                                 -1.328068155288572e+01};
        constexpr double c[] = {-7.784894002430293e-03, -3.223964580411365e-01,
                                 -2.400758277161838e+00, -2.549732539343734e+00,
                                  4.374664141464968e+00,  2.938163982698783e+00};
        constexpr double d[] = { 7.784695709041462e-03,  3.224671290700398e-01,
                                  2.445134137142996e+00,  3.754408661907416e+00};
        constexpr double pLow = 0.02425;
        constexpr double pHigh = 1.0 - pLow;
        double q = 0.0, r = 0.0, x = 0.0;
        if (p < pLow) {
            q = std::sqrt(-2.0 * std::log(p));
            x = (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) /
                ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1.0);
        } else if (p <= pHigh) {
            q = p - 0.5;
            r = q * q;
            x = (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q /
                (((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r+1.0);
        } else {
            q = std::sqrt(-2.0 * std::log(1.0 - p));
            x = -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) /
                  ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1.0);
        }
        return x;
    }

    static double stat_xlMathIncGamma(double a, double x)
    {
        if (x < 0.0 || a <= 0.0) return std::numeric_limits<double>::quiet_NaN();
        if (x == 0.0) return 0.0;
        const double logGammaA = std::lgamma(a);
        if (x < a + 1.0) {
            double ap = a;
            double del = 1.0 / a;
            double sum = del;
            for (int i = 0; i < 200; ++i) {
                ap += 1.0;
                del *= x / ap;
                sum += del;
                if (std::abs(del) < std::abs(sum) * 1e-12) break;
            }
            return sum * std::exp(-x + a * std::log(x) - logGammaA);
        } else {
            double fpmin = 1e-300;
            double b = x + 1.0 - a;
            double c = 1.0 / fpmin;
            double d = 1.0 / b;
            double h = d;
            for (int i = 1; i <= 200; ++i) {
                double an = -i * (i - a);
                b += 2.0;
                d = an * d + b; if (std::abs(d) < fpmin) d = fpmin;
                c = b + an / c; if (std::abs(c) < fpmin) c = fpmin;
                d = 1.0 / d;
                double del2 = d * c;
                h *= del2;
                if (std::abs(del2 - 1.0) < 1e-12) break;
            }
            return 1.0 - std::exp(-x + a * std::log(x) - logGammaA) * h;
        }
    }

    static double stat_xlMathIncBetaCF(double a, double b, double x)
    {
        constexpr double fpmin = 1e-300;
        double qab = a + b;
        double qap = a + 1.0;
        double qam = a - 1.0;
        double c = 1.0;
        double d = 1.0 - qab * x / qap;
        if (std::abs(d) < fpmin) d = fpmin;
        d = 1.0 / d;
        double h = d;
        for (int m = 1; m <= 200; ++m) {
            int m2 = 2 * m;
            double aa = m * (b - m) * x / ((qam + m2) * (a + m2));
            d = 1.0 + aa * d; if (std::abs(d) < fpmin) d = fpmin;
            c = 1.0 + aa / c; if (std::abs(c) < fpmin) c = fpmin;
            d = 1.0 / d;
            h *= d * c;
            aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
            d = 1.0 + aa * d; if (std::abs(d) < fpmin) d = fpmin;
            c = 1.0 + aa / c; if (std::abs(c) < fpmin) c = fpmin;
            d = 1.0 / d;
            double del = d * c;
            h *= del;
            if (std::abs(del - 1.0) < 1e-12) break;
        }
        return h;
    }

    static double stat_xlMathIncBeta(double a, double b, double x)
    {
        if (x < 0.0 || x > 1.0) return std::numeric_limits<double>::quiet_NaN();
        if (x == 0.0) return 0.0;
        if (x == 1.0) return 1.0;
        double logBeta = std::lgamma(a) + std::lgamma(b) - std::lgamma(a + b);
        double front = std::exp(std::log(x) * a + std::log(1.0 - x) * b - logBeta) / a;
        if (x < (a + 1.0) / (a + b + 2.0))
            return front * stat_xlMathIncBetaCF(a, b, x);
        else
            return 1.0 - (std::exp(std::log(1.0-x) * b + std::log(x) * a - logBeta) / b)
                         * stat_xlMathIncBetaCF(b, a, 1.0 - x);
    }

    static double stat_xlMathTCDF(double t, double df)
    {
        double x = df / (df + t * t);
        return 1.0 - 0.5 * stat_xlMathIncBeta(df / 2.0, 0.5, x);
    }

    static double stat_xlMathTInv1Tail(double p, double df)
    {
        if (p <= 0.0 || p >= 1.0 || df < 1.0)
            return std::numeric_limits<double>::quiet_NaN();
        double lo = -1e9, hi = 1e9;
        for (int i = 0; i < 200; ++i) {
            double mid = 0.5 * (lo + hi);
            double val = stat_xlMathTCDF(mid, df);
            if (std::abs(val - p) < 1e-12) return mid;
            if (val < p) lo = mid;
            else         hi = mid;
        }
        return 0.5 * (lo + hi);
    }

    static double stat_xlMathChiSqInv(double p, double df)
    {
        if (p <= 0.0 || p >= 1.0 || df <= 0.0)
            return std::numeric_limits<double>::quiet_NaN();
        double lo = 0.0, hi = df * 10.0 + 100.0;
        for (int i = 0; i < 300; ++i) {
            double mid = 0.5 * (lo + hi);
            double cdf = stat_xlMathIncGamma(df / 2.0, mid / 2.0);
            if (std::abs(cdf - p) < 1e-12) return mid;
            if (cdf < p) lo = mid;
            else         hi = mid;
        }
        return 0.5 * (lo + hi);
    }

    std::vector<double> stat_extractArrayValuesA(const std::vector<XLFormulaArg>& args)
    {
        std::vector<double> out;
        for (const auto& arg : args) {
            for (const auto& v : arg) {
                if (v.type() == XLValueType::Integer) { out.push_back(static_cast<double>(v.get<int64_t>())); }
                else if (v.type() == XLValueType::Float) {
                    out.push_back(v.get<double>());
                }
                else if (v.type() == XLValueType::Boolean) {
                    out.push_back(v.get<bool>() ? 1.0 : 0.0);
                }
                else if (v.type() == XLValueType::String) {
                    out.push_back(0.0);
                }
            }
        }
        return out;
    }
} // namespace

XLCellValue Formula::fnCount(const std::vector<XLFormulaArg>& args)
{
    int64_t cnt = 0;
    for (const auto& arg : args) {
        for (const auto& v : arg) {
            if (isNumeric(v)) ++cnt;
        }
    }
    return XLCellValue(cnt);
}

XLCellValue Formula::fnCounta(const std::vector<XLFormulaArg>& args)
{
    int64_t cnt = 0;
    for (const auto& arg : args) {
        for (const auto& v : arg) {
            if (!isEmpty(v)) {
                if (v.type() == XLValueType::String && v.get<std::string>().empty()) { continue; }
                ++cnt;
            }
        }
    }
    return XLCellValue(cnt);
}

XLCellValue Formula::fnCountif(const std::vector<XLFormulaArg>& args)
{
    // COUNTIF(range, criteria)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::string criteria = toString(args[1][0]);
    int64_t     cnt      = 0;
    for (const auto& v : args[0])
        if (stat_matchesCriteria(v, criteria)) ++cnt;
    return XLCellValue(cnt);
}

XLCellValue Formula::fnSumif(const std::vector<XLFormulaArg>& args)
{
    // SUMIF(range, criteria, [sum_range])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const auto& range    = args[0];
    std::string criteria = toString(args[1][0]);
    const auto& sumRange = (args.size() > 2) ? args[2] : args[0];
    double      total    = 0.0;
    for (std::size_t i = 0; i < range.size(); ++i) {
        if (stat_matchesCriteria(range[i], criteria)) {
            if (i < sumRange.size() && isNumeric(sumRange[i])) total += toDouble(sumRange[i]);
        }
    }
    return XLCellValue(total);
}

XLCellValue Formula::fnCountifs(const std::vector<XLFormulaArg>& args)
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
            if (!stat_matchesCriteria(args[p][i], toString(args[p + 1][0]))) {
                match = false;
                break;
            }
        }
        if (match) ++cnt;
    }
    return XLCellValue(cnt);
}

XLCellValue Formula::fnSumifs(const std::vector<XLFormulaArg>& args)
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
            if (!stat_matchesCriteria(args[p][i], toString(args[p + 1][0]))) {
                match = false;
                break;
            }
        }
        if (match && isNumeric(sumRange[i])) total += toDouble(sumRange[i]);
    }
    return XLCellValue(total);
}

XLCellValue Formula::fnMaxifs(const std::vector<XLFormulaArg>& args)
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
            if (!stat_matchesCriteria(args[p][i], toString(args[p + 1][0]))) {
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

XLCellValue Formula::fnMinifs(const std::vector<XLFormulaArg>& args)
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
            if (!stat_matchesCriteria(args[p][i], toString(args[p + 1][0]))) {
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

XLCellValue Formula::fnAverageif(const std::vector<XLFormulaArg>& args)
{
    // AVERAGEIF(range, criteria, [average_range])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const auto& range    = args[0];
    std::string criteria = toString(args[1][0]);
    const auto& avgRange = (args.size() > 2 && !args[2].empty()) ? args[2] : args[0];
    double      total    = 0.0;
    int64_t     cnt      = 0;
    for (std::size_t i = 0; i < range.size(); ++i) {
        if (stat_matchesCriteria(range[i], criteria)) {
            if (i < avgRange.size() && isNumeric(avgRange[i])) {
                total += toDouble(avgRange[i]);
                ++cnt;
            }
        }
    }
    if (cnt == 0) return errDiv0();
    return XLCellValue(total / static_cast<double>(cnt));
}

XLCellValue Formula::fnRank(const std::vector<XLFormulaArg>& args)
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

XLCellValue Formula::fnLarge(const std::vector<XLFormulaArg>& args)
{
    // LARGE(array, k)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    auto nums = numerics(args[0]);
    auto k    = static_cast<std::ptrdiff_t>(toDouble(args[1][0]));
    if (k < 1 || k > static_cast<std::ptrdiff_t>(nums.size())) return errNum();
    std::nth_element(nums.begin(), nums.begin() + (static_cast<std::ptrdiff_t>(nums.size()) - k), nums.end());
    return XLCellValue(nums[static_cast<std::size_t>(static_cast<std::ptrdiff_t>(nums.size()) - k)]);
}

XLCellValue Formula::fnSmall(const std::vector<XLFormulaArg>& args)
{
    // SMALL(array, k)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    auto nums = numerics(args[0]);
    auto k    = static_cast<std::ptrdiff_t>(toDouble(args[1][0]));
    if (k < 1 || k > static_cast<std::ptrdiff_t>(nums.size())) return errNum();
    std::nth_element(nums.begin(), nums.begin() + (k - 1), nums.end());
    return XLCellValue(nums[static_cast<std::size_t>(k - 1)]);
}

XLCellValue Formula::fnStdev(const std::vector<XLFormulaArg>& args)
{
    // STDEV (sample) – two-pass for numerical stability
    auto nums = numerics(args);
    if (nums.size() < 2) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq   = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(std::sqrt(sq / static_cast<double>(nums.size() - 1)));
}

XLCellValue Formula::fnVar(const std::vector<XLFormulaArg>& args)
{
    // VAR (sample variance)
    auto nums = numerics(args);
    if (nums.size() < 2) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq   = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(sq / static_cast<double>(nums.size() - 1));
}

XLCellValue Formula::fnMedian(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errNum();
    std::sort(nums.begin(), nums.end());
    std::size_t n = nums.size();
    if (n % 2 == 1) return XLCellValue(nums[n / 2]);
    return XLCellValue((nums[n / 2 - 1] + nums[n / 2]) / 2.0);
}

XLCellValue Formula::fnCountblank(const std::vector<XLFormulaArg>& args)
{
    int64_t cnt = 0;
    for (const auto& arg : args) {
        for (const auto& v : arg) {
            if (isEmpty(v) || (v.type() == XLValueType::String && v.get<std::string>().empty())) ++cnt;
        }
    }
    return XLCellValue(cnt);
}

XLCellValue Formula::fnVarp(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq   = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(sq / static_cast<double>(nums.size()));
}

XLCellValue Formula::fnStdevp(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq   = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(std::sqrt(sq / static_cast<double>(nums.size())));
}

XLCellValue Formula::fnVara(const std::vector<XLFormulaArg>& args)
{
    auto nums = stat_extractArrayValuesA(args);
    if (nums.size() < 2) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq   = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(sq / static_cast<double>(nums.size() - 1));
}

XLCellValue Formula::fnVarpa(const std::vector<XLFormulaArg>& args)
{
    auto nums = stat_extractArrayValuesA(args);
    if (nums.empty()) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq   = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(sq / static_cast<double>(nums.size()));
}

XLCellValue Formula::fnStdeva(const std::vector<XLFormulaArg>& args)
{
    auto nums = stat_extractArrayValuesA(args);
    if (nums.size() < 2) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq   = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(std::sqrt(sq / static_cast<double>(nums.size() - 1)));
}

XLCellValue Formula::fnStdevpa(const std::vector<XLFormulaArg>& args)
{
    auto nums = stat_extractArrayValuesA(args);
    if (nums.empty()) return errDiv0();
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / static_cast<double>(nums.size());
    double sq   = 0.0;
    for (double v : nums) sq += (v - mean) * (v - mean);
    return XLCellValue(std::sqrt(sq / static_cast<double>(nums.size())));
}

XLCellValue Formula::fnPermut(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 2) return errValue();
    double n = std::trunc(nums[0]);
    double k = std::trunc(nums[1]);
    if (n < 0 || k < 0 || n < k) return errNum();
    double res = 1.0;
    for (double i = n; i > n - k; --i) { res *= i; }
    return XLCellValue(res);
}

XLCellValue Formula::fnPermutationa(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 2) return errValue();
    double n = std::trunc(nums[0]);
    double k = std::trunc(nums[1]);
    if (n < 0 || k < 0) return errNum();
    return XLCellValue(std::pow(n, k));
}

XLCellValue Formula::fnFisher(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 1) return errValue();
    double x = nums[0];
    if (x <= -1 || x >= 1) return errNum();
    return XLCellValue(0.5 * std::log((1 + x) / (1 - x)));
}

XLCellValue Formula::fnFisherinv(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 1) return errValue();
    double y = nums[0];
    return XLCellValue((std::exp(2 * y) - 1) / (std::exp(2 * y) + 1));
}

XLCellValue Formula::fnStandardize(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 3) return errValue();
    double x     = nums[0];
    double mean  = nums[1];
    double stdev = nums[2];
    if (stdev <= 0) return errNum();
    return XLCellValue((x - mean) / stdev);
}

XLCellValue Formula::fnPearson(const std::vector<XLFormulaArg>& args)
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

XLCellValue Formula::fnAverageifs(const std::vector<XLFormulaArg>& args)
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
    for (size_t i = 1; i < args.size(); ++i) { countArgs.push_back(args[i]); }
    XLCellValue count = fnCountifs(countArgs);
    if (count.type() == XLValueType::Error) return count;

    double dCount = count.get<double>();
    if (dCount == 0) return errDiv0();

    return XLCellValue(sum.get<double>() / dCount);
}

XLCellValue Formula::fnCovarianceP(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums1 = numerics(args[0]);
    auto nums2 = numerics(args[1]);
    if (nums1.size() != nums2.size() || nums1.empty()) return errNA();

    double mean1 = std::accumulate(nums1.begin(), nums1.end(), 0.0) / nums1.size();
    double mean2 = std::accumulate(nums2.begin(), nums2.end(), 0.0) / nums2.size();

    double cov = 0.0;
    for (size_t i = 0; i < nums1.size(); ++i) { cov += (nums1[i] - mean1) * (nums2[i] - mean2); }
    return XLCellValue(cov / nums1.size());
}

XLCellValue Formula::fnCovarianceS(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums1 = numerics(args[0]);
    auto nums2 = numerics(args[1]);
    if (nums1.size() != nums2.size() || nums1.size() < 2) return errDiv0();

    double mean1 = std::accumulate(nums1.begin(), nums1.end(), 0.0) / nums1.size();
    double mean2 = std::accumulate(nums2.begin(), nums2.end(), 0.0) / nums2.size();

    double cov = 0.0;
    for (size_t i = 0; i < nums1.size(); ++i) { cov += (nums1[i] - mean1) * (nums2[i] - mean2); }
    return XLCellValue(cov / (nums1.size() - 1));
}

XLCellValue Formula::fnTrimmean(const std::vector<XLFormulaArg>& args)
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
    for (size_t i = k; i < nums.size() - k; ++i) { sum += nums[i]; }
    return XLCellValue(sum / (nums.size() - 2 * k));
}

XLCellValue Formula::fnSlope(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums1 = numerics(args[0]);    // y
    auto nums2 = numerics(args[1]);    // x
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

XLCellValue Formula::fnIntercept(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums1 = numerics(args[0]);    // y
    auto nums2 = numerics(args[1]);    // x
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

XLCellValue Formula::fnPercentileInc(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    auto nums = numerics(args[0]);
    if (nums.empty() || args[1].empty()) return errNum();

    double k = toDouble(args[1][0]);
    if (k < 0.0 || k > 1.0) return errNum();

    std::sort(nums.begin(), nums.end());
    double pos   = k * (nums.size() - 1);
    size_t index = static_cast<size_t>(pos);
    double frac  = pos - index;

    if (index + 1 < nums.size()) { return XLCellValue(nums[index] + frac * (nums[index + 1] - nums[index])); }
    else {
        return XLCellValue(nums[index]);
    }
}

XLCellValue Formula::fnPercentileExc(const std::vector<XLFormulaArg>& args)
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
    double frac  = pos - index;

    if (index + 1 < nums.size()) { return XLCellValue(nums[index] + frac * (nums[index + 1] - nums[index])); }
    else {
        return XLCellValue(nums[index]);
    }
}

XLCellValue Formula::fnQuartileInc(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    if (args[1].empty()) return errNum();
    int quart = static_cast<int>(std::floor(toDouble(args[1][0])));
    if (quart < 0 || quart > 4) return errNum();

    std::vector<XLFormulaArg> newArgs = args;
    newArgs[1]                        = {XLCellValue(quart / 4.0)};
    return fnPercentileInc(newArgs);
}

XLCellValue Formula::fnQuartileExc(const std::vector<XLFormulaArg>& args)
{
    if (args.size() != 2) return errValue();
    if (args[1].empty()) return errNum();
    int quart = static_cast<int>(std::floor(toDouble(args[1][0])));
    if (quart < 1 || quart > 3) return errNum();

    std::vector<XLFormulaArg> newArgs = args;
    newArgs[1]                        = {XLCellValue(quart / 4.0)};
    return fnPercentileExc(newArgs);
}

XLCellValue Formula::fnRsq(const std::vector<XLFormulaArg>& args)
{
    XLCellValue pearson = fnPearson(args);
    if (pearson.type() == XLValueType::Error) return pearson;
    double p = pearson.get<double>();
    return XLCellValue(p * p);
}

XLCellValue Formula::fnAvedev(const std::vector<XLFormulaArg>& args)
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

XLCellValue Formula::fnDevsq(const std::vector<XLFormulaArg>& args)
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

XLCellValue Formula::fnAveragea(const std::vector<XLFormulaArg>& args)
{
    double total = 0.0;
    int    count = 0;
    for (const auto& arg : args) {
        bool isScalar = (arg.type() == XLFormulaArg::Type::Scalar);
        for (const auto& v : arg) {
            if (v.type() == XLValueType::Empty) continue;
            count++;
            if (v.type() == XLValueType::Integer || v.type() == XLValueType::Float) { total += toDouble(v); }
            else if (v.type() == XLValueType::Boolean) {
                total += (v.get<bool>() ? 1.0 : 0.0);
            }
            else if (v.type() == XLValueType::String) {
                if (isScalar) {
                    try {
                        total += std::stod(v.get<std::string>());
                    }
                    catch (...) {
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

XLCellValue Formula::fnModeSngl(const std::vector<XLFormulaArg>& args)
{
    // MODE / MODE.SNGL — returns the most frequently occurring value
    auto nums = numerics(args);
    if (nums.empty()) return errNA();

    std::sort(nums.begin(), nums.end());

    double bestVal   = nums[0];
    int    bestCount = 1;
    double curVal    = nums[0];
    int    curCount  = 1;

    for (std::size_t i = 1; i < nums.size(); ++i) {
        if (nums[i] == curVal) {
            ++curCount;
        } else {
            if (curCount > bestCount) {
                bestCount = curCount;
                bestVal   = curVal;
            }
            curVal   = nums[i];
            curCount = 1;
        }
    }
    // Check the last run
    if (curCount > bestCount) {
        bestCount = curCount;
        bestVal   = curVal;
    }

    if (bestCount < 2) return errNA();  // No repeated value → #N/A like Excel
    return XLCellValue(bestVal);
}

XLCellValue Formula::fnSkew(const std::vector<XLFormulaArg>& args)
{
    // SKEW(number1, number2, ...) — sample skewness
    auto nums = numerics(args);
    auto n    = static_cast<double>(nums.size());
    if (n < 3.0) return errDiv0();

    double mean = 0.0;
    for (double v : nums) mean += v;
    mean /= n;

    double m2 = 0.0;
    for (double v : nums) {
        double d = v - mean;
        m2 += d * d;
    }
    m2 /= n;

    double stdev = std::sqrt(m2 * n / (n - 1.0));  // sample stdev
    if (stdev == 0.0) return errDiv0();

    // Excel SKEW: n/((n-1)*(n-2)) * Σ((xi-mean)/s)^3
    double sum3 = 0.0;
    for (double v : nums) {
        double z = (v - mean) / stdev;
        sum3 += z * z * z;
    }
    return XLCellValue(n / ((n - 1.0) * (n - 2.0)) * sum3);
}

XLCellValue Formula::fnKurt(const std::vector<XLFormulaArg>& args)
{
    // KURT(number1, ...) — sample excess kurtosis
    auto nums = numerics(args);
    auto n    = static_cast<double>(nums.size());
    if (n < 4.0) return errDiv0();

    double mean = 0.0;
    for (double v : nums) mean += v;
    mean /= n;

    double m2 = 0.0;
    for (double v : nums) { double d = v - mean; m2 += d * d; }
    m2 /= n;

    double stdev = std::sqrt(m2 * n / (n - 1.0));
    if (stdev == 0.0) return errDiv0();

    double sum4 = 0.0;
    for (double v : nums) {
        double z = (v - mean) / stdev;
        sum4 += z * z * z * z;
    }
    // Excel KURT: n*(n+1)/((n-1)*(n-2)*(n-3)) * sum4 - 3*(n-1)^2/((n-2)*(n-3))
    double k = n * (n + 1.0) / ((n - 1.0) * (n - 2.0) * (n - 3.0)) * sum4
               - 3.0 * (n - 1.0) * (n - 1.0) / ((n - 2.0) * (n - 3.0));
    return XLCellValue(k);
}

XLCellValue Formula::fnForecastLinear(const std::vector<XLFormulaArg>& args)
{
    // FORECAST.LINEAR(x, known_y's, known_x's)
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double      x     = toDouble(args[0][0]);
    auto        ys    = numerics(args[1]);
    auto        xs    = numerics(args[2]);
    std::size_t n     = std::min(ys.size(), xs.size());
    if (n < 2) return errNA();

    double meanX = 0.0, meanY = 0.0;
    for (std::size_t i = 0; i < n; ++i) { meanX += xs[i]; meanY += ys[i]; }
    meanX /= static_cast<double>(n);
    meanY /= static_cast<double>(n);

    double ssXX = 0.0, ssXY = 0.0;
    for (std::size_t i = 0; i < n; ++i) {
        ssXX += (xs[i] - meanX) * (xs[i] - meanX);
        ssXY += (xs[i] - meanX) * (ys[i] - meanY);
    }
    if (ssXX == 0.0) return errDiv0();

    double slope     = ssXY / ssXX;
    double intercept = meanY - slope * meanX;
    return XLCellValue(intercept + slope * x);
}

XLCellValue Formula::fnNormSDist(const std::vector<XLFormulaArg>& args)
{
    // NORM.S.DIST(z, cumulative)
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double z      = toDouble(args[0][0]);
    bool   cumul  = (args.size() < 2 || args[1].empty()) ? true
                  : (toDouble(args[1][0]) != 0.0);
    if (!cumul) {
        // PDF: φ(z) = e^(-z²/2) / √(2π)
        return XLCellValue(std::exp(-0.5 * z * z) / stat_kSqrt2Pi);
    }
    // CDF: Φ(z) = 0.5 * erfc(-z / √2)
    return XLCellValue(0.5 * std::erfc(-z / std::sqrt(2.0)));
}

XLCellValue Formula::fnNormDist(const std::vector<XLFormulaArg>& args)
{
    // NORM.DIST(x, mean, standard_dev, cumulative)
    if (args.size() < 4 || args[0].empty() || args[1].empty() || args[2].empty() || args[3].empty())
        return errValue();
    double x     = toDouble(args[0][0]);
    double mu    = toDouble(args[1][0]);
    double sigma = toDouble(args[2][0]);
    bool   cumul = (toDouble(args[3][0]) != 0.0);
    if (sigma <= 0.0) return errNum();

    double z = (x - mu) / sigma;
    if (!cumul) {
        return XLCellValue(std::exp(-0.5 * z * z) / (sigma * stat_kSqrt2Pi));
    }
    return XLCellValue(0.5 * std::erfc(-z / std::sqrt(2.0)));
}

XLCellValue Formula::fnNormSInv(const std::vector<XLFormulaArg>& args)
{
    // NORM.S.INV(probability)
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double p = toDouble(args[0][0]);
    if (p <= 0.0 || p >= 1.0) return errNum();
    double result = stat_xlMathNormSInv(p);
    return std::isnan(result) ? errNum() : XLCellValue(result);
}

XLCellValue Formula::fnNormInv(const std::vector<XLFormulaArg>& args)
{
    // NORM.INV(probability, mean, standard_dev)
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double p     = toDouble(args[0][0]);
    double mu    = toDouble(args[1][0]);
    double sigma = toDouble(args[2][0]);
    if (p <= 0.0 || p >= 1.0) return errNum();
    if (sigma <= 0.0) return errNum();
    double z = stat_xlMathNormSInv(p);
    return std::isnan(z) ? errNum() : XLCellValue(mu + sigma * z);
}

XLCellValue Formula::fnTDist(const std::vector<XLFormulaArg>& args)
{
    // T.DIST(x, deg_freedom, cumulative)
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double x    = toDouble(args[0][0]);
    double df   = toDouble(args[1][0]);
    bool   cumul = (toDouble(args[2][0]) != 0.0);
    if (df < 1.0) return errNum();

    if (!cumul) {
        // PDF: Γ((df+1)/2) / (√(df·π) · Γ(df/2)) · (1 + x²/df)^(-(df+1)/2)
        double coeff = std::exp(std::lgamma((df + 1.0) / 2.0)
                              - std::lgamma(df / 2.0)
                              - 0.5 * std::log(df * stat_kPi));
        return XLCellValue(coeff * std::pow(1.0 + x * x / df, -(df + 1.0) / 2.0));
    }
    // CDF using incomplete beta
    double betaVal = stat_xlMathIncBeta(df / 2.0, 0.5, df / (df + x * x));
    double cdf     = (x >= 0.0) ? 1.0 - 0.5 * betaVal : 0.5 * betaVal;
    return XLCellValue(cdf);
}

XLCellValue Formula::fnTDistRT(const std::vector<XLFormulaArg>& args)
{
    // T.DIST.RT(x, deg_freedom) — right-tail = 1 - CDF
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double x  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (df < 1.0 || x < 0.0) return errNum();
    double betaVal = stat_xlMathIncBeta(df / 2.0, 0.5, df / (df + x * x));
    return XLCellValue(0.5 * betaVal);
}

XLCellValue Formula::fnTDist2T(const std::vector<XLFormulaArg>& args)
{
    // T.DIST.2T(x, deg_freedom) — two-tail = 2 * T.DIST.RT
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double x  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (df < 1.0 || x < 0.0) return errNum();
    double betaVal = stat_xlMathIncBeta(df / 2.0, 0.5, df / (df + x * x));
    return XLCellValue(betaVal);
}

XLCellValue Formula::fnTInv(const std::vector<XLFormulaArg>& args)
{
    // T.INV(probability, deg_freedom) — left-tail (one-tail) inverse
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double p  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (p <= 0.0 || p >= 1.0 || df < 1.0) return errNum();
    return XLCellValue(stat_xlMathTInv1Tail(p, df));
}

XLCellValue Formula::fnTInv2T(const std::vector<XLFormulaArg>& args)
{
    // T.INV.2T(probability, deg_freedom) — two-tail inverse (returns positive t)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double p  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (p <= 0.0 || p > 1.0 || df < 1.0) return errNum();
    // two-tail: find t such that 2*RT(t,df)=p → RT(t,df)=p/2 → one-tail at 1-p/2
    return XLCellValue(std::abs(stat_xlMathTInv1Tail(1.0 - p / 2.0, df)));
}

XLCellValue Formula::fnChisqDist(const std::vector<XLFormulaArg>& args)
{
    // CHISQ.DIST(x, deg_freedom, cumulative)
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double x     = toDouble(args[0][0]);
    double df    = toDouble(args[1][0]);
    bool   cumul = (toDouble(args[2][0]) != 0.0);
    if (x < 0.0 || df < 1.0) return errNum();

    if (!cumul) {
        // PDF: x^(df/2-1) * e^(-x/2) / (2^(df/2) * Γ(df/2))
        double pdf = std::exp((df / 2.0 - 1.0) * std::log(x)
                             - x / 2.0
                             - (df / 2.0) * std::log(2.0)
                             - std::lgamma(df / 2.0));
        return XLCellValue(pdf);
    }
    return XLCellValue(stat_xlMathIncGamma(df / 2.0, x / 2.0));
}

XLCellValue Formula::fnChisqDistRT(const std::vector<XLFormulaArg>& args)
{
    // CHISQ.DIST.RT(x, deg_freedom) — right-tail
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double x  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (x < 0.0 || df < 1.0) return errNum();
    return XLCellValue(1.0 - stat_xlMathIncGamma(df / 2.0, x / 2.0));
}

XLCellValue Formula::fnChisqInv(const std::vector<XLFormulaArg>& args)
{
    // CHISQ.INV(probability, deg_freedom)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double p  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (p < 0.0 || p >= 1.0 || df < 1.0) return errNum();
    return XLCellValue(stat_xlMathChiSqInv(p, df));
}

XLCellValue Formula::fnChisqInvRT(const std::vector<XLFormulaArg>& args)
{
    // CHISQ.INV.RT(probability, deg_freedom) — right-tail: invert 1-CDF
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double p  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (p <= 0.0 || p > 1.0 || df < 1.0) return errNum();
    return XLCellValue(stat_xlMathChiSqInv(1.0 - p, df));
}

XLCellValue Formula::fnBinomDist(const std::vector<XLFormulaArg>& args)
{
    // BINOM.DIST(number_s, trials, probability_s, cumulative)
    if (args.size() < 4 || args[0].empty() || args[1].empty() || args[2].empty() || args[3].empty())
        return errValue();
    double k     = std::trunc(toDouble(args[0][0]));
    double n     = std::trunc(toDouble(args[1][0]));
    double p     = toDouble(args[2][0]);
    bool   cumul = (toDouble(args[3][0]) != 0.0);
    if (k < 0.0 || k > n || n < 0.0 || p < 0.0 || p > 1.0) return errNum();

    if (!cumul) {
        // PMF: C(n,k) * p^k * (1-p)^(n-k)
        // Compute via log to avoid overflow
        double logPmf = std::lgamma(n + 1.0) - std::lgamma(k + 1.0) - std::lgamma(n - k + 1.0)
                      + k * std::log(p == 0.0 ? 1.0 : p)
                      + (n - k) * std::log(p == 1.0 ? 1.0 : 1.0 - p);
        if (p == 0.0 && k > 0.0) return XLCellValue(0.0);
        if (p == 1.0 && k < n)   return XLCellValue(0.0);
        return XLCellValue(std::exp(logPmf));
    }
    // CDF: I_{1-p}(n-k, k+1) — regularised incomplete beta
    if (k >= n) return XLCellValue(1.0);
    return XLCellValue(stat_xlMathIncBeta(n - k, k + 1.0, 1.0 - p));
}

XLCellValue Formula::fnPoissonDist(const std::vector<XLFormulaArg>& args)
{
    // POISSON.DIST(x, mean, cumulative)
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double k     = std::trunc(toDouble(args[0][0]));
    double lam   = toDouble(args[1][0]);
    bool   cumul = (toDouble(args[2][0]) != 0.0);
    if (k < 0.0 || lam < 0.0) return errNum();

    if (!cumul) {
        // PMF: λ^k * e^-λ / k!
        return XLCellValue(std::exp(k * std::log(lam == 0.0 ? 1.0 : lam) - lam - std::lgamma(k + 1.0)));
    }
    // CDF: 1 - P(k+1, λ) = incomplete gamma
    return XLCellValue(1.0 - stat_xlMathIncGamma(k + 1.0, lam));
}

XLCellValue Formula::fnExponDist(const std::vector<XLFormulaArg>& args)
{
    // EXPON.DIST(x, lambda, cumulative)
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double x     = toDouble(args[0][0]);
    double lam   = toDouble(args[1][0]);
    bool   cumul = (toDouble(args[2][0]) != 0.0);
    if (x < 0.0 || lam <= 0.0) return errNum();

    if (!cumul) return XLCellValue(lam * std::exp(-lam * x));  // PDF
    return XLCellValue(1.0 - std::exp(-lam * x));               // CDF
}

XLCellValue Formula::fnSubtotal(const std::vector<XLFormulaArg>& args)
{
    // SUBTOTAL(function_num, ref1, ...)
    // function_num 1-11 (include hidden) or 101-111 (ignore hidden — we treat same for now)
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    int fn = static_cast<int>(toDouble(args[0][0]));
    if (fn > 100) fn -= 100;  // strip "ignore hidden" flag — treated same

    std::vector<XLFormulaArg> dataArgs(args.begin() + 1, args.end());

    switch (fn) {
        case 1:  return fnAverage(dataArgs);
        case 2:  return fnCount(dataArgs);
        case 3:  return fnCounta(dataArgs);
        case 4:  return fnMax(dataArgs);
        case 5:  return fnMin(dataArgs);
        case 6:  return fnProduct(dataArgs);
        case 7:  return fnStdev(dataArgs);
        case 8:  return fnStdevp(dataArgs);
        case 9:  return fnSum(dataArgs);
        case 10: return fnVar(dataArgs);
        case 11: return fnVarp(dataArgs);
        default: return errValue();
    }
}

void OpenXLSX::registerStatisticalFunctions(XLFormulaRegistry& registry)
{
    registry.registerFunction("COUNT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCount));
    registry.registerFunction("COUNTA", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCounta));
    registry.registerFunction("COUNTIF", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCountif));
    registry.registerFunction("SUMIF", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSumif));
    registry.registerFunction("COUNTIFS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCountifs));
    registry.registerFunction("SUMIFS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSumifs));
    registry.registerFunction("MAXIFS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMaxifs));
    registry.registerFunction("MINIFS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMinifs));
    registry.registerFunction("AVERAGEIF", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAverageif));
    registry.registerFunction("RANK", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRank));
    registry.registerFunction("LARGE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnLarge));
    registry.registerFunction("SMALL", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSmall));
    registry.registerFunction("STDEV", std::make_shared<XLSimpleFormulaFunction>(Formula::fnStdev));
    registry.registerFunction("VAR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnVar));
    registry.registerFunction("MEDIAN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMedian));
    registry.registerFunction("COUNTBLANK", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCountblank));
    registry.registerFunction("VAR.P", std::make_shared<XLSimpleFormulaFunction>(Formula::fnVarp));
    registry.registerFunction("STDEV.P", std::make_shared<XLSimpleFormulaFunction>(Formula::fnStdevp));
    registry.registerFunction("VARA", std::make_shared<XLSimpleFormulaFunction>(Formula::fnVara));
    registry.registerFunction("VARPA", std::make_shared<XLSimpleFormulaFunction>(Formula::fnVarpa));
    registry.registerFunction("STDEVA", std::make_shared<XLSimpleFormulaFunction>(Formula::fnStdeva));
    registry.registerFunction("STDEVPA", std::make_shared<XLSimpleFormulaFunction>(Formula::fnStdevpa));
    registry.registerFunction("PERMUT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnPermut));
    registry.registerFunction("PERMUTATIONA", std::make_shared<XLSimpleFormulaFunction>(Formula::fnPermutationa));
    registry.registerFunction("FISHER", std::make_shared<XLSimpleFormulaFunction>(Formula::fnFisher));
    registry.registerFunction("FISHERINV", std::make_shared<XLSimpleFormulaFunction>(Formula::fnFisherinv));
    registry.registerFunction("STANDARDIZE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnStandardize));
    registry.registerFunction("PEARSON", std::make_shared<XLSimpleFormulaFunction>(Formula::fnPearson));
    registry.registerFunction("AVERAGEIFS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAverageifs));
    registry.registerFunction("COVARIANCE.P", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCovarianceP));
    registry.registerFunction("COVARIANCE.S", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCovarianceS));
    registry.registerFunction("TRIMMEAN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTrimmean));
    registry.registerFunction("SLOPE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSlope));
    registry.registerFunction("INTERCEPT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIntercept));
    registry.registerFunction("PERCENTILE.INC", std::make_shared<XLSimpleFormulaFunction>(Formula::fnPercentileInc));
    registry.registerFunction("PERCENTILE.EXC", std::make_shared<XLSimpleFormulaFunction>(Formula::fnPercentileExc));
    registry.registerFunction("QUARTILE.INC", std::make_shared<XLSimpleFormulaFunction>(Formula::fnQuartileInc));
    registry.registerFunction("QUARTILE.EXC", std::make_shared<XLSimpleFormulaFunction>(Formula::fnQuartileExc));
    registry.registerFunction("RSQ", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRsq));
    registry.registerFunction("AVEDEV", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAvedev));
    registry.registerFunction("DEVSQ", std::make_shared<XLSimpleFormulaFunction>(Formula::fnDevsq));
    registry.registerFunction("AVERAGEA", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAveragea));
    registry.registerFunction("MODE.SNGL", std::make_shared<XLSimpleFormulaFunction>(Formula::fnModeSngl));
    registry.registerFunction("SKEW", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSkew));
    registry.registerFunction("KURT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnKurt));
    registry.registerFunction("FORECAST.LINEAR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnForecastLinear));
    registry.registerFunction("NORM.S.DIST", std::make_shared<XLSimpleFormulaFunction>(Formula::fnNormSDist));
    registry.registerFunction("NORM.DIST", std::make_shared<XLSimpleFormulaFunction>(Formula::fnNormDist));
    registry.registerFunction("NORM.S.INV", std::make_shared<XLSimpleFormulaFunction>(Formula::fnNormSInv));
    registry.registerFunction("NORM.INV", std::make_shared<XLSimpleFormulaFunction>(Formula::fnNormInv));
    registry.registerFunction("T.DIST", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTDist));
    registry.registerFunction("T.DIST.RT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTDistRT));
    registry.registerFunction("T.DIST.2T", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTDist2T));
    registry.registerFunction("T.INV", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTInv));
    registry.registerFunction("T.INV.2T", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTInv2T));
    registry.registerFunction("CHISQ.DIST", std::make_shared<XLSimpleFormulaFunction>(Formula::fnChisqDist));
    registry.registerFunction("CHISQ.DIST.RT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnChisqDistRT));
    registry.registerFunction("CHISQ.INV", std::make_shared<XLSimpleFormulaFunction>(Formula::fnChisqInv));
    registry.registerFunction("CHISQ.INV.RT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnChisqInvRT));
    registry.registerFunction("BINOM.DIST", std::make_shared<XLSimpleFormulaFunction>(Formula::fnBinomDist));
    registry.registerFunction("POISSON.DIST", std::make_shared<XLSimpleFormulaFunction>(Formula::fnPoissonDist));
    registry.registerFunction("EXPON.DIST", std::make_shared<XLSimpleFormulaFunction>(Formula::fnExponDist));
    registry.registerFunction("SUBTOTAL", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSubtotal));
}

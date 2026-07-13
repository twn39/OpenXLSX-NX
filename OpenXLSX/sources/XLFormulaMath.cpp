#include "XLFormulaFunctionsInternal.hpp"
#include "XLFormulaRegistry.hpp"
#include "XLDateTime.hpp"
#include "XLNumberFormatter.hpp"
#include "XLFormulaUtils.hpp"
#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>
#include <limits>
#include <numeric>
#include <random>
#include <regex>
#include <unordered_set>
#include <cctype>
#include <string>

using namespace OpenXLSX;

namespace
{
    inline std::mt19937_64& math_getThreadLocalRNG() { return formulaRng(); }

    constexpr double math_kPi = 3.14159265358979323846;

    std::vector<double> math_extractArrayStringFallback(const XLFormulaArg& arg)
    {
        std::vector<double> res;
        if (arg.empty()) return res;
        const auto& v = arg[0];
        if (v.type() == XLValueType::String) {
            std::string s = v.get<std::string>();
            if (!s.empty() && s.front() == '{' && s.back() == '}') {
                s                 = s.substr(1, s.length() - 2);
                std::size_t start = 0;
                while (start < s.length()) {
                    std::size_t pos = s.find_first_of(",;", start);
                    if (pos == std::string::npos) pos = s.length();
                    std::string token = s.substr(start, pos - start);
                    try {
                        res.push_back(std::stod(token));
                    }
                    catch (...) {
                    }
                    start = pos + 1;
                }
            }
        }
        else if (isNumeric(v)) {
            res.push_back(toDouble(v));
        }
        return res;
    }
} // namespace

XLCellValue Formula::fnSum(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    auto        nums = numericsOrError(args, err);
    if (isError(err)) return err;
    double total = 0.0;
    for (double v : nums) total += v;
    return XLCellValue(total);
}

XLCellValue Formula::fnAverage(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    auto        nums = numericsOrError(args, err);
    if (isError(err)) return err;
    if (nums.empty()) return errDiv0();
    double total = 0.0;
    for (double v : nums) total += v;
    return XLCellValue(total / static_cast<double>(nums.size()));
}

XLCellValue Formula::fnMin(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    auto        nums = numericsOrError(args, err);
    if (isError(err)) return err;
    if (nums.empty()) return XLCellValue(0.0);
    return XLCellValue(*std::min_element(nums.begin(), nums.end()));
}

XLCellValue Formula::fnMax(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    auto        nums = numericsOrError(args, err);
    if (isError(err)) return err;
    if (nums.empty()) return XLCellValue(0.0);
    return XLCellValue(*std::max_element(nums.begin(), nums.end()));
}

namespace
{
    /** Extract two coerced scalars for binary math; blank → 0. */
    bool mathTwoScalars(const std::vector<XLFormulaArg>& args, double& a, double& b, XLCellValue& errOut)
    {
        if (args.size() < 2) {
            errOut = errValue();
            return false;
        }
        auto ca = coerceArgScalar(args[0], true);
        if (isError(ca)) {
            errOut = ca;
            return false;
        }
        auto cb = coerceArgScalar(args[1], true);
        if (isError(cb)) {
            errOut = cb;
            return false;
        }
        a = toDouble(ca);
        b = toDouble(cb);
        return true;
    }

    bool mathOneScalar(const std::vector<XLFormulaArg>& args, double& a, XLCellValue& errOut, bool blankAsZero = true)
    {
        if (args.empty()) {
            errOut = errValue();
            return false;
        }
        auto c = coerceArgScalar(args[0], blankAsZero);
        if (isError(c)) {
            errOut = c;
            return false;
        }
        a = toDouble(c);
        return true;
    }

    /**
     * @brief Classic Excel CEILING: same-sign significance; round away from zero along significance.
     */
    XLCellValue mathCeilingClassic(double number, double significance)
    {
        if (significance == 0.0) return XLCellValue(0.0);
        if (number == 0.0) return XLCellValue(0.0);
        // Different signs → #NUM! (classic CEILING / FLOOR)
        if ((number > 0.0 && significance < 0.0) || (number < 0.0 && significance > 0.0)) return errNum();
        return XLCellValue(std::ceil(number / significance) * significance);
    }

    /**
     * @brief Classic Excel FLOOR: same-sign significance; round toward zero along significance axis.
     */
    XLCellValue mathFloorClassic(double number, double significance)
    {
        if (significance == 0.0) return XLCellValue(0.0);
        if (number == 0.0) return XLCellValue(0.0);
        if ((number > 0.0 && significance < 0.0) || (number < 0.0 && significance > 0.0)) return errNum();
        return XLCellValue(std::floor(number / significance) * significance);
    }

    /**
     * @brief CEILING.PRECISE / ISO.CEILING: ignore significance sign; always round toward +∞.
     */
    XLCellValue mathCeilingPrecise(double number, double significance)
    {
        const double sig = std::abs(significance);
        if (sig == 0.0) return XLCellValue(0.0);
        if (number == 0.0) return XLCellValue(0.0);
        return XLCellValue(std::ceil(number / sig) * sig);
    }

    /**
     * @brief FLOOR.PRECISE: ignore significance sign; always round toward −∞.
     */
    XLCellValue mathFloorPrecise(double number, double significance)
    {
        const double sig = std::abs(significance);
        if (sig == 0.0) return XLCellValue(0.0);
        if (number == 0.0) return XLCellValue(0.0);
        return XLCellValue(std::floor(number / sig) * sig);
    }
}    // namespace

XLCellValue Formula::fnAbs(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      x = 0.0;
    if (!mathOneScalar(args, x, err, /*blankAsZero=*/false)) return err;
    return XLCellValue(std::abs(x));
}

XLCellValue Formula::fnRound(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      num = 0.0, dig = 0.0;
    if (!mathTwoScalars(args, num, dig, err)) return err;
    int    digits = static_cast<int>(dig);
    double factor = std::pow(10.0, digits);
    return XLCellValue(std::round(num * factor) / factor);
}

XLCellValue Formula::fnRoundup(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      num = 0.0, dig = 0.0;
    if (!mathTwoScalars(args, num, dig, err)) return err;
    int    digits = static_cast<int>(dig);
    double factor = std::pow(10.0, digits);
    return XLCellValue(num >= 0 ? std::ceil(num * factor) / factor : std::floor(num * factor) / factor);
}

XLCellValue Formula::fnRounddown(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      num = 0.0, dig = 0.0;
    if (!mathTwoScalars(args, num, dig, err)) return err;
    int    digits = static_cast<int>(dig);
    double factor = std::pow(10.0, digits);
    return XLCellValue(num >= 0 ? std::floor(num * factor) / factor : std::ceil(num * factor) / factor);
}

XLCellValue Formula::fnSqrt(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      d = 0.0;
    if (!mathOneScalar(args, d, err, /*blankAsZero=*/false)) return err;
    if (d < 0.0) return errNum();
    return XLCellValue(std::sqrt(d));
}

XLCellValue Formula::fnPi(const std::vector<XLFormulaArg>& args)
{
    (void)args;
    return XLCellValue(3.14159265358979323846);
}

XLCellValue Formula::fnSin(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::sin(toDouble(args[0][0])));
}

XLCellValue Formula::fnCos(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::cos(toDouble(args[0][0])));
}

XLCellValue Formula::fnTan(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::tan(toDouble(args[0][0])));
}

XLCellValue Formula::fnAsin(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double v = toDouble(args[0][0]);
    if (v < -1.0 || v > 1.0) return errNum();
    return XLCellValue(std::asin(v));
}

XLCellValue Formula::fnAcos(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double v = toDouble(args[0][0]);
    if (v < -1.0 || v > 1.0) return errNum();
    return XLCellValue(std::acos(v));
}

XLCellValue Formula::fnDegrees(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(toDouble(args[0][0]) * 180.0 / 3.14159265358979323846);
}

XLCellValue Formula::fnRadians(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(toDouble(args[0][0]) * 3.14159265358979323846 / 180.0);
}

XLCellValue Formula::fnRand(const std::vector<XLFormulaArg>& args)
{
    (void)args;
    std::uniform_real_distribution<double> dist(0.0, 1.0);
    return XLCellValue(dist(math_getThreadLocalRNG()));
}

XLCellValue Formula::fnRandbetween(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty() || !isNumeric(args[0][0]) || !isNumeric(args[1][0])) return errValue();
    int64_t low  = static_cast<int64_t>(std::ceil(toDouble(args[0][0])));
    int64_t high = static_cast<int64_t>(std::floor(toDouble(args[1][0])));
    if (low > high) return errNum();

    std::uniform_int_distribution<int64_t> dist(low, high);
    return XLCellValue(static_cast<double>(dist(math_getThreadLocalRNG())));
}

XLCellValue Formula::fnInt(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      x = 0.0;
    if (!mathOneScalar(args, x, err, /*blankAsZero=*/false)) return err;
    // Excel INT = floor toward −∞
    return XLCellValue(static_cast<int64_t>(std::floor(x)));
}

XLCellValue Formula::fnMod(const std::vector<XLFormulaArg>& args)
{
    // Excel: MOD(n, d) = n - d * INT(n/d), INT = floor toward −∞ (not C fmod).
    XLCellValue err;
    double      n = 0.0, d = 0.0;
    if (!mathTwoScalars(args, n, d, err)) return err;
    if (d == 0.0) return errDiv0();
    return XLCellValue(n - d * std::floor(n / d));
}

XLCellValue Formula::fnPower(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      base = 0.0, expv = 0.0;
    if (!mathTwoScalars(args, base, expv, err)) return err;
    return XLCellValue(std::pow(base, expv));
}

XLCellValue Formula::fnSumproduct(const std::vector<XLFormulaArg>& args)
{
    // SUMPRODUCT(array1, array2, …): element-wise product then sum.
    // Arrays must share the same dimensions (#VALUE! if not). Non-numeric → 0; errors propagate.
    if (args.empty()) return XLCellValue(0.0);

    std::vector<XLFormulaArg> mats;
    mats.reserve(args.size());
    std::size_t rows = 0, cols = 0;

    for (std::size_t i = 0; i < args.size(); ++i) {
        auto m = args[i].materialize();
        if (m.size() == 0) {
            // Treat fully empty as 1×1 blank (product contributes 0)
            m = XLFormulaArg(XLCellValue{});
        }
        const std::size_t r = std::max<std::size_t>(m.rows(), 1);
        const std::size_t c = std::max<std::size_t>(m.cols(), 1);
        // Normalize scalar shape: size-1 array with rows/cols 0 edge cases
        if (m.rows() == 0 || m.cols() == 0) {
            // materialize of empty → size 0; already handled above
        }
        if (i == 0) {
            rows = r;
            cols = c;
        }
        else if (r != rows || c != cols) {
            return errValue();
        }
        mats.push_back(std::move(m));
    }

    double total = 0.0;
    const std::size_t n = rows * cols;
    for (std::size_t i = 0; i < n; ++i) {
        double prod = 1.0;
        for (const auto& m : mats) {
            // Index by linear order; for mismatched internal layout use at(r,c)
            const std::size_t rr = (cols > 0) ? (i / cols) : 0;
            const std::size_t cc = (cols > 0) ? (i % cols) : 0;
            const XLCellValue v  = (m.rows() > 0 && m.cols() > 0) ? m.at(rr, cc) : m[i];
            if (isError(v)) return v;
            if (isEmpty(v)) {
                prod = 0.0;
                continue;
            }
            if (isNumeric(v)) {
                prod *= toDouble(v);
                continue;
            }
            // Numeric strings coerced; other text → 0 (Excel SUMPRODUCT)
            if (v.type() == XLValueType::String) {
                double d = 0.0;
                if (tryParseNumericString(v.get<std::string>(), d))
                    prod *= d;
                else
                    prod = 0.0;
                continue;
            }
            prod = 0.0;
        }
        total += prod;
    }
    return XLCellValue(total);
}

XLCellValue Formula::fnCeil(const std::vector<XLFormulaArg>& args)
{
    // Classic CEILING(number, significance) — different signs → #NUM!
    XLCellValue err;
    double      num = 0.0, sig = 0.0;
    if (!mathTwoScalars(args, num, sig, err)) return err;
    return mathCeilingClassic(num, sig);
}

XLCellValue Formula::fnFloor(const std::vector<XLFormulaArg>& args)
{
    // Classic FLOOR(number, significance) — different signs → #NUM!
    XLCellValue err;
    double      num = 0.0, sig = 0.0;
    if (!mathTwoScalars(args, num, sig, err)) return err;
    return mathFloorClassic(num, sig);
}

XLCellValue Formula::fnCeilingPrecise(const std::vector<XLFormulaArg>& args)
{
    // CEILING.PRECISE(number, [significance]) — significance defaults to 1; sign ignored
    if (args.empty()) return errValue();
    XLCellValue err;
    double      num = 0.0;
    if (!mathOneScalar(args, num, err, true)) return err;
    double sig = 1.0;
    if (args.size() > 1 && args[1].size() > 0) {
        auto cs = coerceArgScalar(args[1], true);
        if (isError(cs)) return cs;
        sig = toDouble(cs);
    }
    return mathCeilingPrecise(num, sig);
}

XLCellValue Formula::fnFloorPrecise(const std::vector<XLFormulaArg>& args)
{
    if (args.empty()) return errValue();
    XLCellValue err;
    double      num = 0.0;
    if (!mathOneScalar(args, num, err, true)) return err;
    double sig = 1.0;
    if (args.size() > 1 && args[1].size() > 0) {
        auto cs = coerceArgScalar(args[1], true);
        if (isError(cs)) return cs;
        sig = toDouble(cs);
    }
    return mathFloorPrecise(num, sig);
}

XLCellValue Formula::fnIsoCeiling(const std::vector<XLFormulaArg>& args)
{
    // ISO.CEILING is specified as equivalent to CEILING.PRECISE
    return fnCeilingPrecise(args);
}

XLCellValue Formula::fnLog(const std::vector<XLFormulaArg>& args)
{
    if (args.empty()) return errValue();
    XLCellValue err;
    double      num = 0.0;
    if (!mathOneScalar(args, num, err, false)) return err;
    double base = 10.0;
    if (args.size() > 1 && args[1].size() > 0) {
        auto cb = coerceArgScalar(args[1], false);
        if (isError(cb)) return cb;
        base = toDouble(cb);
    }
    if (num <= 0.0 || base <= 0.0 || base == 1.0) return errNum();
    return XLCellValue(std::log(num) / std::log(base));
}

XLCellValue Formula::fnLog10(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      num = 0.0;
    if (!mathOneScalar(args, num, err, false)) return err;
    if (num <= 0.0) return errNum();
    return XLCellValue(std::log10(num));
}

XLCellValue Formula::fnExp(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      x = 0.0;
    if (!mathOneScalar(args, x, err, false)) return err;
    return XLCellValue(std::exp(x));
}

XLCellValue Formula::fnSign(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      v = 0.0;
    if (!mathOneScalar(args, v, err, false)) return err;
    int64_t s = 0;
    if (v > 0.0)
        s = 1;
    else if (v < 0.0)
        s = -1;
    return XLCellValue(s);
}

XLCellValue Formula::fnCeilingMath(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty() || nums.size() > 3) return errValue();
    double number       = nums[0];
    double significance = (nums.size() > 1) ? nums[1] : (number >= 0 ? 1.0 : -1.0);
    if (significance == 0) return XLCellValue(0.0);
    double mode = (nums.size() > 2) ? nums[2] : 0.0;

    double res;
    if (number >= 0) { res = std::ceil(number / std::abs(significance)) * std::abs(significance); }
    else {
        if (mode == 0) { res = std::ceil(number / std::abs(significance)) * std::abs(significance); }
        else {
            res = std::floor(number / std::abs(significance)) * std::abs(significance);
        }
    }
    return XLCellValue(res);
}

XLCellValue Formula::fnFloorMath(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty() || nums.size() > 3) return errValue();
    double number       = nums[0];
    double significance = (nums.size() > 1) ? nums[1] : (number >= 0 ? 1.0 : -1.0);
    if (significance == 0) return XLCellValue(0.0);
    double mode = (nums.size() > 2) ? nums[2] : 0.0;

    double res;
    if (number >= 0) { res = std::floor(number / std::abs(significance)) * std::abs(significance); }
    else {
        if (mode == 0) { res = std::floor(number / std::abs(significance)) * std::abs(significance); }
        else {
            res = std::ceil(number / std::abs(significance)) * std::abs(significance);
        }
    }
    return XLCellValue(res);
}

XLCellValue Formula::fnSumsq(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    auto        nums = numericsOrError(args, err);
    if (isError(err)) return err;
    double total = 0.0;
    for (const auto& v : nums) { total += v * v; }
    return XLCellValue(total);
}

XLCellValue Formula::fnSumx2my2(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2) return errValue();
    auto numsX = numerics(args[0]);
    auto numsY = numerics(args[1]);
    if (numsX.empty()) numsX = math_extractArrayStringFallback(args[0]);
    if (numsY.empty()) numsY = math_extractArrayStringFallback(args[1]);

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

XLCellValue Formula::fnSumx2py2(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2) return errValue();
    auto numsX = numerics(args[0]);
    auto numsY = numerics(args[1]);
    if (numsX.empty()) numsX = math_extractArrayStringFallback(args[0]);
    if (numsY.empty()) numsY = math_extractArrayStringFallback(args[1]);
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

XLCellValue Formula::fnSumxmy2(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2) return errValue();
    auto numsX = numerics(args[0]);
    auto numsY = numerics(args[1]);
    if (numsX.empty()) numsX = math_extractArrayStringFallback(args[0]);
    if (numsY.empty()) numsY = math_extractArrayStringFallback(args[1]);
    if (numsX.size() != numsY.size()) return errNA();
    std::size_t sz = numsX.size();
    if (sz == 0) return XLCellValue(0.0);
    double total = 0.0;
    for (std::size_t i = 0; i < sz; ++i) {
        double x    = numsX[i];
        double y    = numsY[i];
        double diff = x - y;
        total += diff * diff;
    }
    return XLCellValue(total);
}

XLCellValue Formula::fnTrunc(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double num = toDouble(args[0][0]);
    if (std::isnan(num)) return errValue();
    int digits = 0;
    if (args.size() > 1 && !args[1].empty()) { digits = static_cast<int>(toDouble(args[1][0])); }
    double factor = std::pow(10.0, digits);
    return XLCellValue(std::trunc(num * factor) / factor);
}

XLCellValue Formula::fnMround(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() != 2) return errValue();
    if (nums[1] == 0) return XLCellValue(0.0);
    if ((nums[0] * nums[1]) < 0) return errNum();
    double res = std::round(nums[0] / nums[1]) * nums[1];
    return XLCellValue(res);
}

XLCellValue Formula::fnLn(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double x = toDouble(args[0][0]);
    if (x <= 0.0) return errNum();
    return XLCellValue(std::log(x));
}

XLCellValue Formula::fnAtan(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::atan(toDouble(args[0][0])));
}

XLCellValue Formula::fnAtan2Fn(const std::vector<XLFormulaArg>& args)
{
    // Excel ATAN2(x_num, y_num) — note: Excel arg order is (x,y), std::atan2 is (y,x)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double x = toDouble(args[0][0]);
    double y = toDouble(args[1][0]);
    if (x == 0.0 && y == 0.0) return errDiv0();
    return XLCellValue(std::atan2(y, x));
}

XLCellValue Formula::fnSinh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::sinh(toDouble(args[0][0])));
}

XLCellValue Formula::fnCosh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::cosh(toDouble(args[0][0])));
}

XLCellValue Formula::fnTanh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::tanh(toDouble(args[0][0])));
}

XLCellValue Formula::fnAsinh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::asinh(toDouble(args[0][0])));
}

XLCellValue Formula::fnAcosh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double x = toDouble(args[0][0]);
    if (x < 1.0) return errNum();
    return XLCellValue(std::acosh(x));
}

XLCellValue Formula::fnAtanh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double x = toDouble(args[0][0]);
    if (x <= -1.0 || x >= 1.0) return errNum();
    return XLCellValue(std::atanh(x));
}

XLCellValue Formula::fnSqrtpi(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double x = toDouble(args[0][0]);
    if (x < 0.0) return errNum();
    return XLCellValue(std::sqrt(x * math_kPi));
}

XLCellValue Formula::fnFact(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double nd = std::trunc(toDouble(args[0][0]));
    if (nd < 0.0 || nd > 170.0) return errNum();
    double result = 1.0;
    for (double i = 2.0; i <= nd; i += 1.0) result *= i;
    return XLCellValue(result);
}

XLCellValue Formula::fnFactdouble(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double nd = std::trunc(toDouble(args[0][0]));
    if (nd < -1.0) return errNum();
    if (nd <= 0.0) return XLCellValue(1.0);  // FACTDOUBLE(0) = FACTDOUBLE(-1) = 1
    double result = 1.0;
    for (double i = nd; i >= 2.0; i -= 2.0) result *= i;
    return XLCellValue(result);
}

XLCellValue Formula::fnCombin(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double n = std::trunc(toDouble(args[0][0]));
    double k = std::trunc(toDouble(args[1][0]));
    if (n < 0.0 || k < 0.0 || k > n) return errNum();
    if (k == 0.0 || k == n) return XLCellValue(1.0);
    // Use the smaller of k and n-k for efficiency
    if (k > n - k) k = n - k;
    double result = 1.0;
    for (double i = 0.0; i < k; i += 1.0) {
        result = result * (n - i) / (i + 1.0);
    }
    return XLCellValue(std::round(result));
}

XLCellValue Formula::fnCombina(const std::vector<XLFormulaArg>& args)
{
    // COMBINA(n,k) = COMBIN(n+k-1, k)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double n = std::trunc(toDouble(args[0][0]));
    double k = std::trunc(toDouble(args[1][0]));
    if (n < 0.0 || k < 0.0) return errNum();
    if (n == 0.0 && k == 0.0) return XLCellValue(1.0);
    if (n == 0.0) return errNum();
    // Delegate to COMBIN logic
    std::vector<XLFormulaArg> combinArgs = {
        XLFormulaArg(XLCellValue(n + k - 1.0)),
        XLFormulaArg(XLCellValue(k))
    };
    return fnCombin(combinArgs);
}

XLCellValue Formula::fnProduct(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    auto        nums = numericsOrError(args, err);
    if (isError(err)) return err;
    if (nums.empty()) return XLCellValue(0.0);
    double result = 1.0;
    for (double v : nums) result *= v;
    return XLCellValue(result);
}

XLCellValue Formula::fnGcd(const std::vector<XLFormulaArg>& args)
{
    auto   nums = numerics(args);
    if (nums.empty()) return errValue();
    for (double v : nums) if (v < 0.0) return errNum();
    int64_t g = static_cast<int64_t>(std::trunc(nums[0]));
    for (std::size_t i = 1; i < nums.size(); ++i) {
        int64_t b = static_cast<int64_t>(std::trunc(nums[i]));
        g = std::gcd(g, b);
    }
    return XLCellValue(static_cast<double>(g));
}

XLCellValue Formula::fnLcm(const std::vector<XLFormulaArg>& args)
{
    auto   nums = numerics(args);
    if (nums.empty()) return errValue();
    for (double v : nums) if (v < 0.0) return errNum();
    int64_t l = static_cast<int64_t>(std::trunc(nums[0]));
    for (std::size_t i = 1; i < nums.size(); ++i) {
        int64_t b = static_cast<int64_t>(std::trunc(nums[i]));
        if (b == 0 || l == 0) { l = 0; break; }
        l = std::lcm(l, b);
    }
    return XLCellValue(static_cast<double>(l));
}

XLCellValue Formula::fnEven(const std::vector<XLFormulaArg>& args)
{
    // Round away from zero to nearest even integer
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double x = toDouble(args[0][0]);
    if (x == 0.0) return XLCellValue(0.0);
    double sign   = (x > 0.0) ? 1.0 : -1.0;
    double ceiled = std::ceil(std::abs(x));
    if (std::fmod(ceiled, 2.0) != 0.0) ceiled += 1.0;
    return XLCellValue(sign * ceiled);
}

XLCellValue Formula::fnOdd(const std::vector<XLFormulaArg>& args)
{
    // Round away from zero to nearest odd integer
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double x = toDouble(args[0][0]);
    if (x == 0.0) return XLCellValue(1.0);
    double sign   = (x > 0.0) ? 1.0 : -1.0;
    double ceiled = std::ceil(std::abs(x));
    if (std::fmod(ceiled, 2.0) == 0.0) ceiled += 1.0;
    return XLCellValue(sign * ceiled);
}

XLCellValue Formula::fnQuotient(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double n = toDouble(args[0][0]);
    double d = toDouble(args[1][0]);
    if (d == 0.0) return errDiv0();
    return XLCellValue(static_cast<int64_t>(std::trunc(n / d)));
}

// =============================================================================
// Phase B – BASE / DECIMAL / MULTINOMIAL / SERIESSUM / extra trig
// =============================================================================

namespace
{
    char mathDigitToChar(int d)
    {
        if (d < 10) return static_cast<char>('0' + d);
        return static_cast<char>('A' + (d - 10));
    }

    int mathCharToDigit(char c)
    {
        if (c >= '0' && c <= '9') return c - '0';
        if (c >= 'A' && c <= 'Z') return 10 + (c - 'A');
        if (c >= 'a' && c <= 'z') return 10 + (c - 'a');
        return -1;
    }

    double mathFactDouble(double nd)
    {
        if (nd < 0.0 || nd > 170.0) return std::numeric_limits<double>::quiet_NaN();
        double result = 1.0;
        for (double i = 2.0; i <= nd; i += 1.0) result *= i;
        return result;
    }
}    // namespace

XLCellValue Formula::fnBase(const std::vector<XLFormulaArg>& args)
{
    // BASE(number, radix, [min_length]) → string
    if (args.size() < 2) return errValue();
    XLCellValue err;
    double      number = 0.0, radixD = 0.0;
    if (!mathTwoScalars(args, number, radixD, err)) return err;
    if (number < 0.0 || number >= 0x1p53) return errNum();    // Excel: 0 … 2^53
    int radix = static_cast<int>(std::floor(radixD));
    if (radix < 2 || radix > 36) return errNum();
    int minLen = 0;
    if (args.size() > 2 && args[2].size() > 0) {
        auto cl = coerceArgScalar(args[2], true);
        if (isError(cl)) return cl;
        minLen = static_cast<int>(std::floor(toDouble(cl)));
        if (minLen < 0 || minLen > 255) return errNum();
    }
    uint64_t n = static_cast<uint64_t>(std::floor(number));
    std::string out;
    if (n == 0) out = "0";
    else {
        while (n > 0) {
            out.push_back(mathDigitToChar(static_cast<int>(n % static_cast<uint64_t>(radix))));
            n /= static_cast<uint64_t>(radix);
        }
        std::reverse(out.begin(), out.end());
    }
    if (static_cast<int>(out.size()) < minLen) out.insert(out.begin(), static_cast<std::size_t>(minLen - static_cast<int>(out.size())), '0');
    return XLCellValue(out);
}

XLCellValue Formula::fnDecimal(const std::vector<XLFormulaArg>& args)
{
    // DECIMAL(text, radix) → number
    if (args.size() < 2) return errValue();
    if (args[0].size() == 0 || args[1].size() == 0) return errValue();
    auto t = args[0].asScalar();
    if (isError(t)) return t;
    auto r = coerceArgScalar(args[1], false);
    if (isError(r)) return r;
    int radix = static_cast<int>(std::floor(toDouble(r)));
    if (radix < 2 || radix > 36) return errNum();
    std::string s = toString(t);
    s = strTrim(std::move(s));
    if (s.empty()) return errNum();
    // Optional leading +
    std::size_t i = 0;
    if (s[0] == '+') ++i;
    if (i >= s.size()) return errNum();
    double value = 0.0;
    for (; i < s.size(); ++i) {
        int dig = mathCharToDigit(s[i]);
        if (dig < 0 || dig >= radix) return errNum();
        value = value * static_cast<double>(radix) + static_cast<double>(dig);
        if (value >= 0x1p53) return errNum();    // Excel range guard
    }
    return XLCellValue(value);
}

XLCellValue Formula::fnMultinomial(const std::vector<XLFormulaArg>& args)
{
    // MULTINOMIAL(n1, n2, …) = FACT(sum ni) / Π FACT(ni)
    if (args.empty()) return errValue();
    XLCellValue err;
    auto        nums = numericsOrError(args, err);
    if (isError(err)) return err;
    if (nums.empty()) return errValue();
    double sum = 0.0;
    for (double v : nums) {
        if (v < 0.0) return errNum();
        double t = std::trunc(v);
        if (t != v && std::abs(t - v) > 1e-9) {
            // non-integer → truncate like Excel INT toward zero for multinomial inputs
        }
        sum += std::trunc(v);
    }
    if (sum > 170.0) return errNum();
    double num = mathFactDouble(sum);
    if (!std::isfinite(num)) return errNum();
    double den = 1.0;
    for (double v : nums) {
        double f = mathFactDouble(std::trunc(v));
        if (!std::isfinite(f) || f == 0.0) return errNum();
        den *= f;
    }
    return XLCellValue(std::round(num / den));
}

XLCellValue Formula::fnSeriessum(const std::vector<XLFormulaArg>& args)
{
    // SERIESSUM(x, n, m, coefficients)
    if (args.size() < 4) return errValue();
    double x = 0.0, n = 0.0, m = 0.0;
    auto   cx = coerceArgScalar(args[0], true);
    if (isError(cx)) return cx;
    auto cn = coerceArgScalar(args[1], true);
    if (isError(cn)) return cn;
    auto cm = coerceArgScalar(args[2], true);
    if (isError(cm)) return cm;
    x = toDouble(cx);
    n = toDouble(cn);
    m = toDouble(cm);

    auto coeffs = args[3].materialize();
    if (coeffs.size() == 0) return errValue();
    double total = 0.0;
    for (std::size_t i = 0; i < coeffs.size(); ++i) {
        const auto v = coeffs[i];
        if (isError(v)) return v;
        if (!isNumeric(v) && !isEmpty(v)) {
            // try string number
            auto c = coerceToNumber(v, true);
            if (isError(c)) return errValue();
            double a = toDouble(c);
            double power = n + static_cast<double>(i) * m;
            total += a * std::pow(x, power);
            continue;
        }
        double a = isEmpty(v) ? 0.0 : toDouble(v);
        double power = n + static_cast<double>(i) * m;
        total += a * std::pow(x, power);
    }
    return XLCellValue(total);
}

XLCellValue Formula::fnCsc(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      x = 0.0;
    if (!mathOneScalar(args, x, err, false)) return err;
    double s = std::sin(x);
    if (s == 0.0) return errDiv0();
    return XLCellValue(1.0 / s);
}

XLCellValue Formula::fnSec(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      x = 0.0;
    if (!mathOneScalar(args, x, err, false)) return err;
    double c = std::cos(x);
    if (c == 0.0) return errDiv0();
    return XLCellValue(1.0 / c);
}

XLCellValue Formula::fnCot(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      x = 0.0;
    if (!mathOneScalar(args, x, err, false)) return err;
    double s = std::sin(x);
    if (s == 0.0) return errDiv0();
    return XLCellValue(std::cos(x) / s);
}

XLCellValue Formula::fnAcot(const std::vector<XLFormulaArg>& args)
{
    // Excel ACOT: range (0, π]; use atan2(1, x)
    XLCellValue err;
    double      x = 0.0;
    if (!mathOneScalar(args, x, err, false)) return err;
    return XLCellValue(std::atan2(1.0, x));
}

XLCellValue Formula::fnAcoth(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      x = 0.0;
    if (!mathOneScalar(args, x, err, false)) return err;
    if (std::abs(x) <= 1.0) return errNum();
    // artanh(1/x) = 0.5 * ln((x+1)/(x-1))
    return XLCellValue(0.5 * std::log((x + 1.0) / (x - 1.0)));
}

XLCellValue Formula::fnCsch(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      x = 0.0;
    if (!mathOneScalar(args, x, err, false)) return err;
    if (x == 0.0) return errDiv0();
    double sh = std::sinh(x);
    if (sh == 0.0) return errDiv0();
    return XLCellValue(1.0 / sh);
}

XLCellValue Formula::fnSech(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      x = 0.0;
    if (!mathOneScalar(args, x, err, false)) return err;
    return XLCellValue(1.0 / std::cosh(x));
}

XLCellValue Formula::fnCoth(const std::vector<XLFormulaArg>& args)
{
    XLCellValue err;
    double      x = 0.0;
    if (!mathOneScalar(args, x, err, false)) return err;
    if (x == 0.0) return errDiv0();
    double sh = std::sinh(x);
    if (sh == 0.0) return errDiv0();
    return XLCellValue(std::cosh(x) / sh);
}

namespace
{
    XLFormulaArg mathErrArg(std::string_view token)
    {
        XLCellValue e;
        e.setError(std::string(token));
        return XLFormulaArg(std::move(e));
    }

    // Phase D: contiguous row-major matrices + dimension cap (protects against O(n³) blow-ups)
    constexpr size_t kMathMaxMatrixDim = 256;

    bool mathToMatrixFlat(const XLFormulaArg& arg, std::vector<double>& mat, size_t& rows, size_t& cols)
    {
        auto a = arg.materialize();
        rows   = a.rows();
        cols   = a.cols();
        if (rows == 0 || cols == 0) return false;
        if (rows > kMathMaxMatrixDim || cols > kMathMaxMatrixDim) return false;
        mat.resize(rows * cols);
        for (size_t r = 0; r < rows; ++r) {
            for (size_t c = 0; c < cols; ++c) {
                const auto& v = a.at(r, c);
                if (isError(v)) return false;
                if (isEmpty(v)) {
                    mat[r * cols + c] = 0.0;
                    continue;
                }
                if (!isNumeric(v)) return false;
                mat[r * cols + c] = toDouble(v);
            }
        }
        return true;
    }

    // MMULT(array1, array2) → m×p matrix (contiguous triple-loop)
    XLFormulaArg fnMmult(const std::vector<XLFormulaArg>& args, XLEvalSession&)
    {
        if (args.size() < 2) return mathErrArg("#VALUE!");
        std::vector<double> A, B;
        size_t m = 0, n = 0, n2 = 0, p = 0;
        if (!mathToMatrixFlat(args[0], A, m, n) || !mathToMatrixFlat(args[1], B, n2, p)) {
            // Distinguish dim overflow vs type error: oversize → #NUM!, else #VALUE!
            auto a0 = args[0].materialize();
            auto a1 = args[1].materialize();
            if (a0.rows() > kMathMaxMatrixDim || a0.cols() > kMathMaxMatrixDim || a1.rows() > kMathMaxMatrixDim ||
                a1.cols() > kMathMaxMatrixDim)
                return mathErrArg("#NUM!");
            return mathErrArg("#VALUE!");
        }
        if (n != n2) return mathErrArg("#VALUE!");
        std::vector<XLCellValue> out;
        out.reserve(m * p);
        for (size_t i = 0; i < m; ++i) {
            for (size_t j = 0; j < p; ++j) {
                double s = 0.0;
                for (size_t k = 0; k < n; ++k) s += A[i * n + k] * B[k * p + j];
                out.emplace_back(s);
            }
        }
        return XLFormulaArg(std::move(out), m, p);
    }

    // Gauss-Jordan inverse on contiguous n×n matrix (row-major).
    bool mathInvertFlat(std::vector<double>& a, std::vector<double>& inv, size_t n)
    {
        if (n == 0) return false;
        inv.assign(n * n, 0.0);
        for (size_t i = 0; i < n; ++i) inv[i * n + i] = 1.0;

        for (size_t col = 0; col < n; ++col) {
            size_t pivot = col;
            double best  = std::abs(a[col * n + col]);
            for (size_t r = col + 1; r < n; ++r) {
                double v = std::abs(a[r * n + col]);
                if (v > best) {
                    best  = v;
                    pivot = r;
                }
            }
            if (best < 1e-14) return false;
            if (pivot != col) {
                for (size_t c = 0; c < n; ++c) {
                    std::swap(a[pivot * n + c], a[col * n + c]);
                    std::swap(inv[pivot * n + c], inv[col * n + c]);
                }
            }
            const double diag = a[col * n + col];
            for (size_t c = 0; c < n; ++c) {
                a[col * n + c] /= diag;
                inv[col * n + c] /= diag;
            }
            for (size_t r = 0; r < n; ++r) {
                if (r == col) continue;
                const double f = a[r * n + col];
                for (size_t c = 0; c < n; ++c) {
                    a[r * n + c] -= f * a[col * n + c];
                    inv[r * n + c] -= f * inv[col * n + c];
                }
            }
        }
        return true;
    }

    XLFormulaArg fnMinverse(const std::vector<XLFormulaArg>& args, XLEvalSession&)
    {
        if (args.empty()) return mathErrArg("#VALUE!");
        std::vector<double> A;
        size_t              n = 0, c = 0;
        if (!mathToMatrixFlat(args[0], A, n, c) || n != c || n == 0) {
            auto a0 = args[0].materialize();
            if (a0.rows() > kMathMaxMatrixDim || a0.cols() > kMathMaxMatrixDim) return mathErrArg("#NUM!");
            return mathErrArg("#VALUE!");
        }
        std::vector<double> inv;
        if (!mathInvertFlat(A, inv, n)) return mathErrArg("#NUM!");
        std::vector<XLCellValue> out;
        out.reserve(n * n);
        for (size_t i = 0; i < n * n; ++i) out.emplace_back(inv[i]);
        return XLFormulaArg(std::move(out), n, n);
    }

    double mathDetFlat(std::vector<double> a, size_t n)
    {
        double det = 1.0;
        for (size_t col = 0; col < n; ++col) {
            size_t pivot = col;
            double best  = std::abs(a[col * n + col]);
            for (size_t r = col + 1; r < n; ++r) {
                double v = std::abs(a[r * n + col]);
                if (v > best) {
                    best  = v;
                    pivot = r;
                }
            }
            if (best < 1e-15) return 0.0;
            if (pivot != col) {
                for (size_t c = 0; c < n; ++c) std::swap(a[pivot * n + c], a[col * n + c]);
                det = -det;
            }
            det *= a[col * n + col];
            const double diag = a[col * n + col];
            for (size_t r = col + 1; r < n; ++r) {
                const double f = a[r * n + col] / diag;
                for (size_t c = col; c < n; ++c) a[r * n + c] -= f * a[col * n + c];
            }
        }
        return det;
    }

    // MDETERM returns scalar
    XLFormulaArg fnMdeterm(const std::vector<XLFormulaArg>& args, XLEvalSession&)
    {
        if (args.empty()) return mathErrArg("#VALUE!");
        std::vector<double> A;
        size_t              n = 0, c = 0;
        if (!mathToMatrixFlat(args[0], A, n, c) || n != c || n == 0) {
            auto a0 = args[0].materialize();
            if (a0.rows() > kMathMaxMatrixDim || a0.cols() > kMathMaxMatrixDim) return mathErrArg("#NUM!");
            return mathErrArg("#VALUE!");
        }
        return XLFormulaArg(XLCellValue(mathDetFlat(std::move(A), n)));
    }
}    // namespace

void OpenXLSX::registerMathFunctions(XLFormulaRegistry& registry)
{
    registry.registerFunction("SUM", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSum));
    registry.registerFunction("AVERAGE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAverage));
    registry.registerFunction("MIN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMin));
    registry.registerFunction("MAX", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMax));
    registry.registerFunction("ABS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAbs));
    registry.registerFunction("ROUND", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRound));
    registry.registerFunction("ROUNDUP", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRoundup));
    registry.registerFunction("ROUNDDOWN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRounddown));
    registry.registerFunction("SQRT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSqrt));
    registry.registerFunction("PI", std::make_shared<XLSimpleFormulaFunction>(Formula::fnPi));
    registry.registerFunction("SIN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSin));
    registry.registerFunction("COS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCos));
    registry.registerFunction("TAN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTan));
    registry.registerFunction("ASIN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAsin));
    registry.registerFunction("ACOS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAcos));
    registry.registerFunction("DEGREES", std::make_shared<XLSimpleFormulaFunction>(Formula::fnDegrees));
    registry.registerFunction("RADIANS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRadians));
    registry.registerFunction("RAND", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRand));
    registry.registerFunction("RANDBETWEEN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRandbetween));
    registry.registerFunction("INT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnInt));
    registry.registerFunction("MOD", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMod));
    registry.registerFunction("POWER", std::make_shared<XLSimpleFormulaFunction>(Formula::fnPower));
    registry.registerFunction("SUMPRODUCT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSumproduct));
    registry.registerFunction("CEILING", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCeil));
    registry.registerFunction("FLOOR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnFloor));
    registry.registerFunction("CEILING.PRECISE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCeilingPrecise));
    registry.registerFunction("FLOOR.PRECISE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnFloorPrecise));
    registry.registerFunction("ISO.CEILING", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIsoCeiling));
    registry.registerFunction("LOG", std::make_shared<XLSimpleFormulaFunction>(Formula::fnLog));
    registry.registerFunction("LOG10", std::make_shared<XLSimpleFormulaFunction>(Formula::fnLog10));
    registry.registerFunction("EXP", std::make_shared<XLSimpleFormulaFunction>(Formula::fnExp));
    registry.registerFunction("SIGN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSign));
    registry.registerFunction("CEILING.MATH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCeilingMath));
    registry.registerFunction("FLOOR.MATH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnFloorMath));
    registry.registerFunction("SUMSQ", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSumsq));
    registry.registerFunction("SUMX2MY2", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSumx2my2));
    registry.registerFunction("SUMX2PY2", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSumx2py2));
    registry.registerFunction("SUMXMY2", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSumxmy2));
    registry.registerFunction("TRUNC", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTrunc));
    registry.registerFunction("MROUND", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMround));
    registry.registerFunction("LN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnLn));
    registry.registerFunction("ATAN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAtan));
    registry.registerFunction("ATAN2", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAtan2Fn));
    registry.registerFunction("SINH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSinh));
    registry.registerFunction("COSH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCosh));
    registry.registerFunction("TANH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTanh));
    registry.registerFunction("ASINH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAsinh));
    registry.registerFunction("ACOSH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAcosh));
    registry.registerFunction("ATANH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAtanh));
    registry.registerFunction("SQRTPI", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSqrtpi));
    registry.registerFunction("FACT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnFact));
    registry.registerFunction("FACTDOUBLE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnFactdouble));
    registry.registerFunction("COMBIN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCombin));
    registry.registerFunction("COMBINA", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCombina));
    registry.registerFunction("PRODUCT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnProduct));
    registry.registerFunction("GCD", std::make_shared<XLSimpleFormulaFunction>(Formula::fnGcd));
    registry.registerFunction("LCM", std::make_shared<XLSimpleFormulaFunction>(Formula::fnLcm));
    registry.registerFunction("EVEN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnEven));
    registry.registerFunction("ODD", std::make_shared<XLSimpleFormulaFunction>(Formula::fnOdd));
    registry.registerFunction("QUOTIENT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnQuotient));
    registry.registerFunction("BASE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnBase));
    registry.registerFunction("DECIMAL", std::make_shared<XLSimpleFormulaFunction>(Formula::fnDecimal));
    registry.registerFunction("MULTINOMIAL", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMultinomial));
    registry.registerFunction("SERIESSUM", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSeriessum));
    registry.registerFunction("CSC", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCsc));
    registry.registerFunction("SEC", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSec));
    registry.registerFunction("COT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCot));
    registry.registerFunction("ACOT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAcot));
    registry.registerFunction("ACOTH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnAcoth));
    registry.registerFunction("CSCH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCsch));
    registry.registerFunction("SECH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSech));
    registry.registerFunction("COTH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCoth));
    registry.registerFunction("MMULT", std::make_shared<XLArrayFormulaFunction>(fnMmult));
    registry.registerFunction("MINVERSE", std::make_shared<XLArrayFormulaFunction>(fnMinverse));
    registry.registerFunction("MDETERM", std::make_shared<XLArrayFormulaFunction>(fnMdeterm));
}

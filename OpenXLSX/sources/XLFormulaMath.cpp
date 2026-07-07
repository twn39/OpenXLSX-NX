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
    inline std::mt19937_64& math_getThreadLocalRNG()
    {
        thread_local std::mt19937_64 engine(std::random_device{}());
        return engine;
    }

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
    double total = 0.0;
    for (const auto& v : numerics(args)) total += v;
    return XLCellValue(total);
}

XLCellValue Formula::fnAverage(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errDiv0();
    double total = 0.0;
    for (double v : nums) total += v;
    return XLCellValue(total / static_cast<double>(nums.size()));
}

XLCellValue Formula::fnMin(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return XLCellValue(0.0);
    return XLCellValue(*std::min_element(nums.begin(), nums.end()));
}

XLCellValue Formula::fnMax(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return XLCellValue(0.0);
    return XLCellValue(*std::max_element(nums.begin(), nums.end()));
}

XLCellValue Formula::fnAbs(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    const auto& v = args[0][0];
    if (!isNumeric(v)) return errValue();
    return XLCellValue(std::abs(toDouble(v)));
}

XLCellValue Formula::fnRound(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num    = toDouble(args[0][0]);
    int    digits = static_cast<int>(toDouble(args[1][0]));
    double factor = std::pow(10.0, digits);
    return XLCellValue(std::round(num * factor) / factor);
}

XLCellValue Formula::fnRoundup(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num    = toDouble(args[0][0]);
    int    digits = static_cast<int>(toDouble(args[1][0]));
    double factor = std::pow(10.0, digits);
    return XLCellValue(num >= 0 ? std::ceil(num * factor) / factor : std::floor(num * factor) / factor);
}

XLCellValue Formula::fnRounddown(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num    = toDouble(args[0][0]);
    int    digits = static_cast<int>(toDouble(args[1][0]));
    double factor = std::pow(10.0, digits);
    return XLCellValue(num >= 0 ? std::floor(num * factor) / factor : std::ceil(num * factor) / factor);
}

XLCellValue Formula::fnSqrt(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double d = toDouble(args[0][0]);
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
    if (args.empty() || args[0].empty()) return errValue();
    if (!isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(std::floor(toDouble(args[0][0]))));
}

XLCellValue Formula::fnMod(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double n = toDouble(args[0][0]), d = toDouble(args[1][0]);
    if (d == 0.0) return errDiv0();
    return XLCellValue(std::fmod(n, d));
}

XLCellValue Formula::fnPower(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    return XLCellValue(std::pow(toDouble(args[0][0]), toDouble(args[1][0])));
}

XLCellValue Formula::fnSumproduct(const std::vector<XLFormulaArg>& args)
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

XLCellValue Formula::fnCeil(const std::vector<XLFormulaArg>& args)
{
    // CEILING(number, significance)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num = toDouble(args[0][0]);
    double sig = toDouble(args[1][0]);
    if (sig == 0.0) return XLCellValue(0.0);
    return XLCellValue(std::ceil(num / sig) * sig);
}

XLCellValue Formula::fnFloor(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double num = toDouble(args[0][0]);
    double sig = toDouble(args[1][0]);
    if (sig == 0.0) return XLCellValue(0.0);
    return XLCellValue(std::floor(num / sig) * sig);
}

XLCellValue Formula::fnLog(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double num  = toDouble(args[0][0]);
    double base = (args.size() > 1 && !args[1].empty()) ? toDouble(args[1][0]) : 10.0;
    if (num <= 0.0 || base <= 0.0 || base == 1.0) return errNum();
    return XLCellValue(std::log(num) / std::log(base));
}

XLCellValue Formula::fnLog10(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double num = toDouble(args[0][0]);
    if (num <= 0.0) return errNum();
    return XLCellValue(std::log10(num));
}

XLCellValue Formula::fnExp(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    return XLCellValue(std::exp(toDouble(args[0][0])));
}

XLCellValue Formula::fnSign(const std::vector<XLFormulaArg>& args)
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
    double total = 0.0;
    for (const auto& v : numerics(args)) { total += v * v; }
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
    auto   nums   = numerics(args);
    double result = 1.0;
    if (nums.empty()) return XLCellValue(0.0);
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
}

/**
 * @file XLFormulaFunctions2.cpp
 * @brief Second batch of built-in Excel formula function implementations.
 *
 * This file adds three groups of functions that were absent from XLFormulaFunctions.cpp:
 *
 *  Batch 1 – Math / Trig (pure <cmath> wrappers):
 *      LN, ATAN, ATAN2, SINH, COSH, TANH, ASINH, ACOSH, ATANH, SQRTPI,
 *      FACT, FACTDOUBLE, COMBIN, COMBINA, PRODUCT, GCD, LCM, EVEN, ODD, QUOTIENT
 *
 *  Batch 2 – High-frequency utility functions:
 *      SUBTOTAL, CHOOSE, ROW, COLUMN, DATEDIF,
 *      IRR, MIRR, RATE, IPMT, PPMT,
 *      MODE.SNGL, SKEW, KURT, FORECAST.LINEAR
 *
 *  Batch 3 – Statistical distributions:
 *      Shared helpers: xlMathNormSInv, xlMathIncGamma, xlMathIncBeta
 *      NORM.S.DIST / NORMSDIST
 *      NORM.DIST   / NORMDIST
 *      NORM.S.INV  / NORMSINV
 *      NORM.INV    / NORMINV
 *      T.DIST / T.DIST.2T / T.DIST.RT / TDIST
 *      T.INV  / T.INV.2T  / TINV
 *      CHISQ.DIST / CHISQ.DIST.RT / CHIDIST
 *      CHISQ.INV  / CHISQ.INV.RT  / CHIINV
 *      BINOM.DIST / BINOMDIST
 *      POISSON.DIST / POISSON
 *      EXPON.DIST / EXPONDIST
 *
 * All implementations follow the project conventions:
 *   - Use numerics(), toDouble(), isNumeric(), isEmpty(), toString()
 *   - Return errValue() / errNum() / errDiv0() / errNA() as appropriate
 *   - Functions are static members of XLFormulaEngine
 */

#include "XLDateTime.hpp"
#include "XLFormulaEngine.hpp"
#include "XLFormulaUtils.hpp"
#include "XLCellReference.hpp"
#include <algorithm>
#include <array>
#include <cmath>
#include <numeric>   // std::gcd, std::lcm (C++17)
#include <string>
#include <vector>

using namespace OpenXLSX;

// ============================================================================
// Internal namespace: shared math helpers (not exposed in headers)
// ============================================================================

namespace
{
    // ---- Constants -----------------------------------------------------------

    constexpr double kPi      = 3.14159265358979323846;
    constexpr double kSqrt2Pi = 2.50662827463100050242; // sqrt(2*pi)

    // ---- Date helpers (mirrors the ones in XLFormulaFunctions.cpp anon ns) ---

    /// Excel serial → std::tm (returns epoch tm on failure)
    static std::tm f2_serialToTm(double serial)
    {
        try { return XLDateTime(serial).tm(); }
        catch (...) { return {}; }
    }

    /// std::tm → Excel serial (returns 0 on failure)
    static double f2_tmToSerial(std::tm t)
    {
        try { return XLDateTime(t).serial(); }
        catch (...) { return 0.0; }
    }

    /// Days in a given month/year
    static int f2_daysInMonth(int year, int month)
    {
        constexpr std::array<int, 12> days = {31,28,31,30,31,30,31,31,30,31,30,31};
        bool leap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
        int  d    = days.at(static_cast<std::size_t>(month - 1));
        if (month == 2 && leap) ++d;
        return d;
    }

    // =========================================================================
    // xlMathNormSInv  –  inverse standard normal CDF
    // Uses the rational approximation from Peter J. Acklam (2002),
    // max error < 1.15e-9 for p in (0,1).
    // =========================================================================
    static double xlMathNormSInv(double p)
    {
        if (p <= 0.0 || p >= 1.0) return std::numeric_limits<double>::quiet_NaN();

        // Coefficients for rational approximation
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

        constexpr double pLow  = 0.02425;
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

    // =========================================================================
    // xlMathIncGamma  –  regularised lower incomplete gamma P(a,x)
    // Uses series expansion for x < a+1, continued fraction for x >= a+1.
    // =========================================================================
    static double xlMathIncGamma(double a, double x)
    {
        if (x < 0.0 || a <= 0.0) return std::numeric_limits<double>::quiet_NaN();
        if (x == 0.0) return 0.0;

        const double logGammaA = std::lgamma(a);

        if (x < a + 1.0) {
            // Series expansion
            double ap  = a;
            double del = 1.0 / a;
            double sum = del;
            for (int i = 0; i < 200; ++i) {
                ap  += 1.0;
                del *= x / ap;
                sum += del;
                if (std::abs(del) < std::abs(sum) * 1e-12) break;
            }
            return sum * std::exp(-x + a * std::log(x) - logGammaA);
        } else {
            // Continued fraction (Lentz's method)
            double fpmin = 1e-300;
            double b     = x + 1.0 - a;
            double c     = 1.0 / fpmin;
            double d     = 1.0 / b;
            double h     = d;
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
            // Upper gamma: 1 - lower
            return 1.0 - std::exp(-x + a * std::log(x) - logGammaA) * h;
        }
    }

    // =========================================================================
    // xlMathIncBeta  –  regularised incomplete beta I_x(a,b)
    // Uses Lentz's continued fraction algorithm.
    // =========================================================================
    static double xlMathIncBetaCF(double a, double b, double x)
    {
        // Continued fraction representation
        constexpr double fpmin = 1e-300;
        double qab  = a + b;
        double qap  = a + 1.0;
        double qam  = a - 1.0;
        double c    = 1.0;
        double d    = 1.0 - qab * x / qap;
        if (std::abs(d) < fpmin) d = fpmin;
        d = 1.0 / d;
        double h = d;

        for (int m = 1; m <= 200; ++m) {
            int    m2 = 2 * m;
            double aa = m * (b - m) * x / ((qam + m2) * (a + m2));
            d = 1.0 + aa * d;
            if (std::abs(d) < fpmin) d = fpmin;
            c = 1.0 + aa / c;
            if (std::abs(c) < fpmin) c = fpmin;
            d = 1.0 / d;
            h *= d * c;
            aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
            d = 1.0 + aa * d;
            if (std::abs(d) < fpmin) d = fpmin;
            c = 1.0 + aa / c;
            if (std::abs(c) < fpmin) c = fpmin;
            d = 1.0 / d;
            double del = d * c;
            h *= del;
            if (std::abs(del - 1.0) < 1e-12) break;
        }
        return h;
    }

    static double xlMathIncBeta(double a, double b, double x)
    {
        if (x < 0.0 || x > 1.0) return std::numeric_limits<double>::quiet_NaN();
        if (x == 0.0) return 0.0;
        if (x == 1.0) return 1.0;

        double logBeta = std::lgamma(a) + std::lgamma(b) - std::lgamma(a + b);
        double front   = std::exp(std::log(x) * a + std::log(1.0 - x) * b - logBeta) / a;

        // Use symmetry to ensure good convergence
        if (x < (a + 1.0) / (a + b + 2.0))
            return front * xlMathIncBetaCF(a, b, x);
        else
            return 1.0 - (std::exp(std::log(1.0-x) * b + std::log(x) * a - logBeta) / b)
                         * xlMathIncBetaCF(b, a, 1.0 - x);
    }

    // =========================================================================
    // xlMathTDistInv  –  inverse of the T distribution CDF (two-tail)
    // Uses binary search + xlMathIncBeta.
    // =========================================================================
    static double xlMathTCDF(double t, double df)
    {
        // CDF of |t| using incomplete beta: P(T ≤ t) for t ≥ 0
        double x = df / (df + t * t);
        return 1.0 - 0.5 * xlMathIncBeta(df / 2.0, 0.5, x);
    }

    static double xlMathTInv1Tail(double p, double df)
    {
        // Solve for t such that CDF(t, df) = p, using bisection
        if (p <= 0.0 || p >= 1.0 || df < 1.0)
            return std::numeric_limits<double>::quiet_NaN();

        double lo = -1e9, hi = 1e9;
        for (int i = 0; i < 200; ++i) {
            double mid = 0.5 * (lo + hi);
            double val = xlMathTCDF(mid, df);
            if (std::abs(val - p) < 1e-12) return mid;
            if (val < p) lo = mid;
            else         hi = mid;
        }
        return 0.5 * (lo + hi);
    }

    // =========================================================================
    // xlMathChiSqInv  –  inverse Chi-squared CDF by bisection
    // =========================================================================
    static double xlMathChiSqInv(double p, double df)
    {
        if (p <= 0.0 || p >= 1.0 || df <= 0.0)
            return std::numeric_limits<double>::quiet_NaN();

        double lo = 0.0, hi = df * 10.0 + 100.0;
        for (int i = 0; i < 300; ++i) {
            double mid = 0.5 * (lo + hi);
            double cdf = xlMathIncGamma(df / 2.0, mid / 2.0);
            if (std::abs(cdf - p) < 1e-12) return mid;
            if (cdf < p) lo = mid;
            else         hi = mid;
        }
        return 0.5 * (lo + hi);
    }

} // anonymous namespace


// ============================================================================
// BATCH 1 — Math / Trig (pure cmath wrappers)
// ============================================================================

XLCellValue XLFormulaEngine::fnLn(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double x = toDouble(args[0][0]);
    if (x <= 0.0) return errNum();
    return XLCellValue(std::log(x));
}

XLCellValue XLFormulaEngine::fnAtan(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::atan(toDouble(args[0][0])));
}

XLCellValue XLFormulaEngine::fnAtan2Fn(const std::vector<XLFormulaArg>& args)
{
    // Excel ATAN2(x_num, y_num) — note: Excel arg order is (x,y), std::atan2 is (y,x)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double x = toDouble(args[0][0]);
    double y = toDouble(args[1][0]);
    if (x == 0.0 && y == 0.0) return errDiv0();
    return XLCellValue(std::atan2(y, x));
}

XLCellValue XLFormulaEngine::fnSinh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::sinh(toDouble(args[0][0])));
}

XLCellValue XLFormulaEngine::fnCosh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::cosh(toDouble(args[0][0])));
}

XLCellValue XLFormulaEngine::fnTanh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::tanh(toDouble(args[0][0])));
}

XLCellValue XLFormulaEngine::fnAsinh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(std::asinh(toDouble(args[0][0])));
}

XLCellValue XLFormulaEngine::fnAcosh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double x = toDouble(args[0][0]);
    if (x < 1.0) return errNum();
    return XLCellValue(std::acosh(x));
}

XLCellValue XLFormulaEngine::fnAtanh(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double x = toDouble(args[0][0]);
    if (x <= -1.0 || x >= 1.0) return errNum();
    return XLCellValue(std::atanh(x));
}

XLCellValue XLFormulaEngine::fnSqrtpi(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double x = toDouble(args[0][0]);
    if (x < 0.0) return errNum();
    return XLCellValue(std::sqrt(x * kPi));
}

XLCellValue XLFormulaEngine::fnFact(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double nd = std::trunc(toDouble(args[0][0]));
    if (nd < 0.0 || nd > 170.0) return errNum();
    double result = 1.0;
    for (double i = 2.0; i <= nd; i += 1.0) result *= i;
    return XLCellValue(result);
}

XLCellValue XLFormulaEngine::fnFactdouble(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double nd = std::trunc(toDouble(args[0][0]));
    if (nd < -1.0) return errNum();
    if (nd <= 0.0) return XLCellValue(1.0);  // FACTDOUBLE(0) = FACTDOUBLE(-1) = 1
    double result = 1.0;
    for (double i = nd; i >= 2.0; i -= 2.0) result *= i;
    return XLCellValue(result);
}

XLCellValue XLFormulaEngine::fnCombin(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnCombina(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnProduct(const std::vector<XLFormulaArg>& args)
{
    auto   nums   = numerics(args);
    double result = 1.0;
    if (nums.empty()) return XLCellValue(0.0);
    for (double v : nums) result *= v;
    return XLCellValue(result);
}

XLCellValue XLFormulaEngine::fnGcd(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnLcm(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnEven(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnOdd(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnQuotient(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double n = toDouble(args[0][0]);
    double d = toDouble(args[1][0]);
    if (d == 0.0) return errDiv0();
    return XLCellValue(static_cast<int64_t>(std::trunc(n / d)));
}


// ============================================================================
// BATCH 2 — High-frequency utility functions
// ============================================================================

XLCellValue XLFormulaEngine::fnSubtotal(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnChoose(const std::vector<XLFormulaArg>& args)
{
    // CHOOSE(index_num, value1, value2, ...)
    if (args.size() < 2 || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    auto   idx = static_cast<std::size_t>(toDouble(args[0][0]));
    if (idx < 1 || idx >= args.size()) return errValue();
    return args[idx].empty() ? XLCellValue{} : args[idx][0];
}

XLCellValue XLFormulaEngine::fnRow(const std::vector<XLFormulaArg>& args)
{
    // ROW([reference]) — return the 1-based row of the top-left cell of the argument
    if (args.empty() || args[0].empty()) return errValue();
    const auto& arg = args[0];
    
    // Best case: LazyRange carries exact row/col metadata
    if (arg.type() == XLFormulaArg::Type::LazyRange) {
        uint32_t r = arg.firstRow();
        if (r > 0) return XLCellValue(static_cast<int64_t>(r));
    }
    
    // Fallback: value may be a string like "A5" or numeric address — try parsing
    const auto& v = arg[0];
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
    } catch (...) {
        return errValue();
    }
}

XLCellValue XLFormulaEngine::fnColumn(const std::vector<XLFormulaArg>& args)
{
    // COLUMN([reference]) — return the 1-based column of the top-left cell
    if (args.empty() || args[0].empty()) return errValue();
    const auto& arg = args[0];
    
    if (arg.type() == XLFormulaArg::Type::LazyRange) {
        uint16_t c = arg.firstCol();
        if (c > 0) return XLCellValue(static_cast<int64_t>(c));
    }
    
    const auto& v = arg[0];
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
    } catch (...) {
        return errValue();
    }
}

XLCellValue XLFormulaEngine::fnDatedif(const std::vector<XLFormulaArg>& args)
{
    // DATEDIF(start_date, end_date, unit)
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    if (!isNumeric(args[0][0]) || !isNumeric(args[1][0])) return errValue();

    double d1Serial = toDouble(args[0][0]);
    double d2Serial = toDouble(args[1][0]);
    if (d1Serial > d2Serial) return errNum();  // start must be <= end

    auto tm1 = f2_serialToTm(d1Serial);
    auto tm2 = f2_serialToTm(d2Serial);
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
        if (days < 0) days += f2_daysInMonth(y2, m2 == 1 ? 12 : m2 - 1);
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
        double annSerial = f2_tmToSerial(tmAnn);
        if (annSerial > d2Serial) {
            tmAnn.tm_year -= 1;
            annSerial = f2_tmToSerial(tmAnn);
        }
        return XLCellValue(static_cast<int64_t>(std::floor(d2Serial) - std::floor(annSerial)));
    }
    return errValue();
}

XLCellValue XLFormulaEngine::fnIrr(const std::vector<XLFormulaArg>& args)
{
    // IRR(values, [guess=0.1])
    if (args.empty() || args[0].empty()) return errValue();
    auto cashflows = numerics(args[0]);
    if (cashflows.empty()) return errNum();

    double rate = (args.size() > 1 && !args[1].empty()) ? toDouble(args[1][0]) : 0.1;

    // Newton-Raphson iteration
    for (int iter = 0; iter < 128; ++iter) {
        double npv  = 0.0;
        double dnpv = 0.0;
        for (std::size_t i = 0; i < cashflows.size(); ++i) {
            double denom = std::pow(1.0 + rate, static_cast<double>(i));
            npv  += cashflows[i] / denom;
            dnpv -= static_cast<double>(i) * cashflows[i] / (denom * (1.0 + rate));
        }
        if (std::abs(dnpv) < 1e-15) return errNum();
        double newRate = rate - npv / dnpv;
        if (std::abs(newRate - rate) < 1e-10) {
            if (newRate <= -1.0) return errNum();
            return XLCellValue(newRate);
        }
        rate = newRate;
        if (rate <= -1.0) rate = -0.9999;  // clamp to valid domain
    }
    return errNum();  // Did not converge
}

XLCellValue XLFormulaEngine::fnMirr(const std::vector<XLFormulaArg>& args)
{
    // MIRR(values, finance_rate, reinvest_rate)
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    auto   cashflows     = numerics(args[0]);
    double financeRate   = toDouble(args[1][0]);
    double reinvestRate  = toDouble(args[2][0]);
    int    n             = static_cast<int>(cashflows.size());
    if (n < 2) return errValue();

    // PV of negative cashflows at finance_rate
    double pvNeg = 0.0;
    // FV of positive cashflows at reinvest_rate
    double fvPos = 0.0;

    for (int i = 0; i < n; ++i) {
        if (cashflows[i] < 0.0)
            pvNeg += cashflows[i] / std::pow(1.0 + financeRate, static_cast<double>(i));
        else if (cashflows[i] > 0.0)
            fvPos += cashflows[i] * std::pow(1.0 + reinvestRate, static_cast<double>(n - 1 - i));
    }

    if (pvNeg == 0.0 || fvPos == 0.0) return errDiv0();
    return XLCellValue(std::pow(-fvPos / pvNeg, 1.0 / static_cast<double>(n - 1)) - 1.0);
}

XLCellValue XLFormulaEngine::fnRate(const std::vector<XLFormulaArg>& args)
{
    // RATE(nper, pmt, pv, [fv=0], [type=0], [guess=0.1])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double nper  = toDouble(args[0][0]);
    double pmt   = toDouble(args[1][0]);
    double pv    = toDouble(args[2][0]);
    double fv    = (args.size() > 3 && !args[3].empty()) ? toDouble(args[3][0]) : 0.0;
    int    type  = (args.size() > 4 && !args[4].empty()) ? static_cast<int>(toDouble(args[4][0])) : 0;
    double rate  = (args.size() > 5 && !args[5].empty()) ? toDouble(args[5][0]) : 0.1;

    // Newton-Raphson: f(r) = pv*(1+r)^n + pmt*(1+r*type)*((1+r)^n - 1)/r + fv = 0
    for (int iter = 0; iter < 300; ++iter) {
        if (std::abs(rate) < 1e-15) {
            // r ≈ 0 linearisation
            double f  = pv + pmt * nper + fv;
            double df = pmt * nper * (1.0 + type * 0) - pmt;
            if (std::abs(df) < 1e-15) return errNum();
            rate -= f / df;
            continue;
        }
        double q   = std::pow(1.0 + rate, nper);
        double f   = pv * q + pmt * (1.0 + rate * type) * (q - 1.0) / rate + fv;
        double dq  = nper * std::pow(1.0 + rate, nper - 1.0);
        double df  = pv * dq +
                     pmt * type * (q - 1.0) / rate +
                     pmt * (1.0 + rate * type) * (dq * rate - (q - 1.0)) / (rate * rate);
        if (std::abs(df) < 1e-15) return errNum();
        double newRate = rate - f / df;
        if (std::abs(newRate - rate) < 1e-10) return XLCellValue(newRate);
        rate = newRate;
    }
    return errNum();
}

XLCellValue XLFormulaEngine::fnIpmt(const std::vector<XLFormulaArg>& args)
{
    // IPMT(rate, per, nper, pv, [fv=0], [type=0])
    if (args.size() < 4 || args[0].empty() || args[1].empty() || args[2].empty() || args[3].empty()) return errValue();
    double rate = toDouble(args[0][0]);
    double per  = toDouble(args[1][0]);
    double nper = toDouble(args[2][0]);
    double pv   = toDouble(args[3][0]);
    double fv   = (args.size() > 4 && !args[4].empty()) ? toDouble(args[4][0]) : 0.0;
    int    type = (args.size() > 5 && !args[5].empty()) ? static_cast<int>(toDouble(args[5][0])) : 0;

    if (per < 1.0 || per > nper) return errNum();

    // PMT using the same formula as fnPmt
    double pmt = 0.0;
    if (rate == 0.0) {
        pmt = -(pv + fv) / nper;
    } else {
        double q = std::pow(1.0 + rate, nper);
        pmt = -(pv * q + fv) * rate / ((1.0 + rate * type) * (q - 1.0));
    }

    // IPMT(per) = -PV(per-1) * rate
    double pvAtPer = pv * std::pow(1.0 + rate, per - 1.0) +
                     pmt * (1.0 + rate * type) *
                     (std::pow(1.0 + rate, per - 1.0) - 1.0) / (rate == 0.0 ? 1.0 : rate);

    if (rate == 0.0) return XLCellValue(0.0);
    double ipmt = -pvAtPer * rate;
    if (type == 1 && per == 1) ipmt = 0.0;
    else if (type == 1) ipmt = ipmt / (1.0 + rate);
    return XLCellValue(ipmt);
}

XLCellValue XLFormulaEngine::fnPpmt(const std::vector<XLFormulaArg>& args)
{
    // PPMT = PMT - IPMT
    if (args.size() < 4) return errValue();

    // Compute PMT
    std::vector<XLFormulaArg> pmtArgs = {args[0], args[2], args[3]};
    if (args.size() > 4) pmtArgs.push_back(args[4]);
    if (args.size() > 5) pmtArgs.push_back(args[5]);
    XLCellValue pmtVal = fnPmt(pmtArgs);
    if (pmtVal.type() == XLValueType::Error) return pmtVal;

    XLCellValue ipmtVal = fnIpmt(args);
    if (ipmtVal.type() == XLValueType::Error) return ipmtVal;

    double pmt  = pmtVal.type()  == XLValueType::Float ? pmtVal.get<double>()
                                                        : static_cast<double>(pmtVal.get<int64_t>());
    double ipmt = ipmtVal.type() == XLValueType::Float ? ipmtVal.get<double>()
                                                       : static_cast<double>(ipmtVal.get<int64_t>());
    return XLCellValue(pmt - ipmt);
}

XLCellValue XLFormulaEngine::fnModeSngl(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnSkew(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnKurt(const std::vector<XLFormulaArg>& args)
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

XLCellValue XLFormulaEngine::fnForecastLinear(const std::vector<XLFormulaArg>& args)
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


// ============================================================================
// BATCH 3 — Statistical Distribution Functions
// ============================================================================

// ----- Normal Distribution --------------------------------------------------

XLCellValue XLFormulaEngine::fnNormSDist(const std::vector<XLFormulaArg>& args)
{
    // NORM.S.DIST(z, cumulative)
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double z      = toDouble(args[0][0]);
    bool   cumul  = (args.size() < 2 || args[1].empty()) ? true
                  : (toDouble(args[1][0]) != 0.0);
    if (!cumul) {
        // PDF: φ(z) = e^(-z²/2) / √(2π)
        return XLCellValue(std::exp(-0.5 * z * z) / kSqrt2Pi);
    }
    // CDF: Φ(z) = 0.5 * erfc(-z / √2)
    return XLCellValue(0.5 * std::erfc(-z / std::sqrt(2.0)));
}

XLCellValue XLFormulaEngine::fnNormDist(const std::vector<XLFormulaArg>& args)
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
        return XLCellValue(std::exp(-0.5 * z * z) / (sigma * kSqrt2Pi));
    }
    return XLCellValue(0.5 * std::erfc(-z / std::sqrt(2.0)));
}

XLCellValue XLFormulaEngine::fnNormSInv(const std::vector<XLFormulaArg>& args)
{
    // NORM.S.INV(probability)
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    double p = toDouble(args[0][0]);
    if (p <= 0.0 || p >= 1.0) return errNum();
    double result = xlMathNormSInv(p);
    return std::isnan(result) ? errNum() : XLCellValue(result);
}

XLCellValue XLFormulaEngine::fnNormInv(const std::vector<XLFormulaArg>& args)
{
    // NORM.INV(probability, mean, standard_dev)
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double p     = toDouble(args[0][0]);
    double mu    = toDouble(args[1][0]);
    double sigma = toDouble(args[2][0]);
    if (p <= 0.0 || p >= 1.0) return errNum();
    if (sigma <= 0.0) return errNum();
    double z = xlMathNormSInv(p);
    return std::isnan(z) ? errNum() : XLCellValue(mu + sigma * z);
}

// ----- T Distribution -------------------------------------------------------

XLCellValue XLFormulaEngine::fnTDist(const std::vector<XLFormulaArg>& args)
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
                              - 0.5 * std::log(df * kPi));
        return XLCellValue(coeff * std::pow(1.0 + x * x / df, -(df + 1.0) / 2.0));
    }
    // CDF using incomplete beta
    double betaVal = xlMathIncBeta(df / 2.0, 0.5, df / (df + x * x));
    double cdf     = (x >= 0.0) ? 1.0 - 0.5 * betaVal : 0.5 * betaVal;
    return XLCellValue(cdf);
}

XLCellValue XLFormulaEngine::fnTDistRT(const std::vector<XLFormulaArg>& args)
{
    // T.DIST.RT(x, deg_freedom) — right-tail = 1 - CDF
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double x  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (df < 1.0 || x < 0.0) return errNum();
    double betaVal = xlMathIncBeta(df / 2.0, 0.5, df / (df + x * x));
    return XLCellValue(0.5 * betaVal);
}

XLCellValue XLFormulaEngine::fnTDist2T(const std::vector<XLFormulaArg>& args)
{
    // T.DIST.2T(x, deg_freedom) — two-tail = 2 * T.DIST.RT
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double x  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (df < 1.0 || x < 0.0) return errNum();
    double betaVal = xlMathIncBeta(df / 2.0, 0.5, df / (df + x * x));
    return XLCellValue(betaVal);
}

XLCellValue XLFormulaEngine::fnTInv(const std::vector<XLFormulaArg>& args)
{
    // T.INV(probability, deg_freedom) — left-tail (one-tail) inverse
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double p  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (p <= 0.0 || p >= 1.0 || df < 1.0) return errNum();
    return XLCellValue(xlMathTInv1Tail(p, df));
}

XLCellValue XLFormulaEngine::fnTInv2T(const std::vector<XLFormulaArg>& args)
{
    // T.INV.2T(probability, deg_freedom) — two-tail inverse (returns positive t)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double p  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (p <= 0.0 || p > 1.0 || df < 1.0) return errNum();
    // two-tail: find t such that 2*RT(t,df)=p → RT(t,df)=p/2 → one-tail at 1-p/2
    return XLCellValue(std::abs(xlMathTInv1Tail(1.0 - p / 2.0, df)));
}

// ----- Chi-Square Distribution ----------------------------------------------

XLCellValue XLFormulaEngine::fnChisqDist(const std::vector<XLFormulaArg>& args)
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
    return XLCellValue(xlMathIncGamma(df / 2.0, x / 2.0));
}

XLCellValue XLFormulaEngine::fnChisqDistRT(const std::vector<XLFormulaArg>& args)
{
    // CHISQ.DIST.RT(x, deg_freedom) — right-tail
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double x  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (x < 0.0 || df < 1.0) return errNum();
    return XLCellValue(1.0 - xlMathIncGamma(df / 2.0, x / 2.0));
}

XLCellValue XLFormulaEngine::fnChisqInv(const std::vector<XLFormulaArg>& args)
{
    // CHISQ.INV(probability, deg_freedom)
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double p  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (p < 0.0 || p >= 1.0 || df < 1.0) return errNum();
    return XLCellValue(xlMathChiSqInv(p, df));
}

XLCellValue XLFormulaEngine::fnChisqInvRT(const std::vector<XLFormulaArg>& args)
{
    // CHISQ.INV.RT(probability, deg_freedom) — right-tail: invert 1-CDF
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double p  = toDouble(args[0][0]);
    double df = toDouble(args[1][0]);
    if (p <= 0.0 || p > 1.0 || df < 1.0) return errNum();
    return XLCellValue(xlMathChiSqInv(1.0 - p, df));
}

// ----- Binomial Distribution ------------------------------------------------

XLCellValue XLFormulaEngine::fnBinomDist(const std::vector<XLFormulaArg>& args)
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
    return XLCellValue(xlMathIncBeta(n - k, k + 1.0, 1.0 - p));
}

// ----- Poisson Distribution -------------------------------------------------

XLCellValue XLFormulaEngine::fnPoissonDist(const std::vector<XLFormulaArg>& args)
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
    return XLCellValue(1.0 - xlMathIncGamma(k + 1.0, lam));
}

// ----- Exponential Distribution ---------------------------------------------

XLCellValue XLFormulaEngine::fnExponDist(const std::vector<XLFormulaArg>& args)
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

/**
 * @file XLFormulaEngineering.cpp
 * @brief Phase C engineering / scientific formula functions (bit, base convert, complex, special).
 */

#include "XLFormulaFunctionsInternal.hpp"
#include "XLFormulaRegistry.hpp"
#include "XLFormulaUtils.hpp"

#include <fmt/format.h>
#include <algorithm>
#include <cmath>
#include <cstdint>
#include <limits>
#include <string>
#include <vector>

using namespace OpenXLSX;

namespace
{
    constexpr int64_t kBitMax = (1LL << 48) - 1;    // Excel BIT* domain: 0 … 2^48-1

    bool engOne(const std::vector<XLFormulaArg>& args, double& a, XLCellValue& err, bool blankAsZero = false)
    {
        if (args.empty()) {
            err = errValue();
            return false;
        }
        auto c = coerceArgScalar(args[0], blankAsZero);
        if (isError(c)) {
            err = c;
            return false;
        }
        a = toDouble(c);
        return true;
    }

    bool engTwo(const std::vector<XLFormulaArg>& args, double& a, double& b, XLCellValue& err, bool blankAsZero = false)
    {
        if (args.size() < 2) {
            err = errValue();
            return false;
        }
        auto ca = coerceArgScalar(args[0], blankAsZero);
        if (isError(ca)) {
            err = ca;
            return false;
        }
        auto cb = coerceArgScalar(args[1], blankAsZero);
        if (isError(cb)) {
            err = cb;
            return false;
        }
        a = toDouble(ca);
        b = toDouble(cb);
        return true;
    }

    bool engBitInt(double v, int64_t& out, XLCellValue& err)
    {
        if (v < 0.0 || v > static_cast<double>(kBitMax) || std::floor(v) != v) {
            err = errNum();
            return false;
        }
        out = static_cast<int64_t>(v);
        return true;
    }

    int engDigit(char c)
    {
        if (c >= '0' && c <= '9') return c - '0';
        if (c >= 'A' && c <= 'Z') return 10 + (c - 'A');
        if (c >= 'a' && c <= 'z') return 10 + (c - 'a');
        return -1;
    }

    char engDigitChar(int d)
    {
        if (d < 10) return static_cast<char>('0' + d);
        return static_cast<char>('A' + (d - 10));
    }

    XLCellValue engFromBase(const std::string& sIn, int radix, int maxDigits)
    {
        std::string s = strTrim(sIn);
        if (s.empty() || static_cast<int>(s.size()) > maxDigits) return errNum();
        // Optional sign for some Excel converters — HEX2DEC allows negative via two's complement for 10 digits
        bool neg = false;
        if (!s.empty() && (s[0] == '-' || s[0] == '+')) {
            neg = (s[0] == '-');
            s.erase(s.begin());
            if (s.empty()) return errNum();
        }
        // Two's complement detection for full-width hex/bin/oct (Excel)
        bool twos = false;
        if (!neg && radix == 16 && s.size() == 10) {
            int msd = engDigit(s[0]);
            if (msd >= 8) twos = true;
        }
        if (!neg && radix == 2 && s.size() == 10) {
            if (s[0] == '1') twos = true;
        }
        if (!neg && radix == 8 && s.size() == 10) {
            int msd = engDigit(s[0]);
            if (msd >= 4) twos = true;
        }

        int64_t value = 0;
        for (char c : s) {
            int d = engDigit(c);
            if (d < 0 || d >= radix) return errNum();
            value = value * radix + d;
        }
        if (twos) {
            // Interpret as 10-digit two's complement in given radix
            int64_t mod = 1;
            for (int i = 0; i < 10; ++i) mod *= radix;
            value -= mod;
        }
        if (neg) value = -value;
        return XLCellValue(static_cast<double>(value));
    }

    XLCellValue engToBase(double number, int radix, int places, int maxPlaces)
    {
        if (std::floor(number) != number) return errNum();
        int64_t n = static_cast<int64_t>(number);
        // Excel places optional padding; negative numbers via two's complement for fixed width
        if (places < 0 || places > maxPlaces) return errNum();

        const bool neg = n < 0;
        uint64_t   u   = 0;
        if (neg) {
            // Two's complement in 10 digits of radix when places==0 use minimal, else pad
            int64_t mod = 1;
            for (int i = 0; i < 10; ++i) mod *= radix;
            u = static_cast<uint64_t>(n + mod);
        }
        else {
            u = static_cast<uint64_t>(n);
        }

        std::string out;
        if (u == 0) out = "0";
        else {
            while (u > 0) {
                out.push_back(engDigitChar(static_cast<int>(u % static_cast<uint64_t>(radix))));
                u /= static_cast<uint64_t>(radix);
            }
            std::reverse(out.begin(), out.end());
        }
        if (places > 0) {
            if (static_cast<int>(out.size()) > places) return errNum();
            if (static_cast<int>(out.size()) < places)
                out.insert(out.begin(), static_cast<std::size_t>(places - static_cast<int>(out.size())), '0');
        }
        return XLCellValue(out);
    }

    // ---- Complex helpers (Excel: "x+yi" / "x+yj") ----
    struct EngComplex
    {
        double re{0.0};
        double im{0.0};
    };

    bool engParseComplex(const XLCellValue& v, EngComplex& out)
    {
        if (isError(v)) return false;
        if (isNumeric(v)) {
            out.re = toDouble(v);
            out.im = 0.0;
            return true;
        }
        if (isEmpty(v)) {
            out = {};
            return true;
        }
        if (v.type() != XLValueType::String) return false;
        std::string s = strTrim(v.get<std::string>());
        if (s.empty()) return false;
        // Normalize j → i
        for (char& c : s) {
            if (c == 'j' || c == 'J') c = 'i';
            if (c == 'I') c = 'i';
        }
        // Pure imaginary: "i", "-i", "3i", "-2.5i"
        if (s == "i") {
            out = {0, 1};
            return true;
        }
        if (s == "-i" || s == "+i") {
            out = {0, s[0] == '-' ? -1.0 : 1.0};
            return true;
        }
        if (!s.empty() && (s.back() == 'i')) {
            std::string coef = s.substr(0, s.size() - 1);
            if (coef.empty() || coef == "+") {
                out = {0, 1};
                return true;
            }
            if (coef == "-") {
                out = {0, -1};
                return true;
            }
            // Forms: "a+bi", "a-bi", "bi"
            auto plus  = coef.find_last_of('+');
            auto minus = coef.find_last_of('-');
            // Find split between re and im (last + or - not at start)
            std::size_t split = std::string::npos;
            if (plus != std::string::npos && plus > 0) split = plus;
            if (minus != std::string::npos && minus > 0 && (split == std::string::npos || minus > split)) split = minus;

            if (split == std::string::npos) {
                // Pure imag coefficient
                try {
                    out.re = 0.0;
                    out.im = std::stod(coef);
                    return true;
                }
                catch (...) {
                    return false;
                }
            }
            try {
                out.re = std::stod(coef.substr(0, split));
                std::string imPart = coef.substr(split);
                if (imPart == "+" || imPart == "-")
                    out.im = (imPart == "-") ? -1.0 : 1.0;
                else
                    out.im = std::stod(imPart);
                return true;
            }
            catch (...) {
                return false;
            }
        }
        // Real only string number
        double d = 0.0;
        if (tryParseNumericString(s, d)) {
            out = {d, 0};
            return true;
        }
        return false;
    }

    std::string engFormatComplex(const EngComplex& z, char suffix = 'i')
    {
        // Excel-like: "3+4i", "3-4i", "4i", "-4i", "3"
        if (z.im == 0.0) return fmt::format("{}", z.re);
        if (z.re == 0.0) {
            if (z.im == 1.0) return std::string(1, suffix);
            if (z.im == -1.0) return std::string("-") + suffix;
            return fmt::format("{}{}", z.im, suffix);
        }
        if (z.im == 1.0) return fmt::format("{}+{}", z.re, suffix);
        if (z.im == -1.0) return fmt::format("{}-{}", z.re, suffix);
        if (z.im < 0.0) return fmt::format("{}{}{}", z.re, z.im, suffix);
        return fmt::format("{}+{}{}", z.re, z.im, suffix);
    }

    constexpr double kPi = 3.14159265358979323846;

    double engLogGamma(double z)
    {
        // Lanczos coefficients (g=7)
        static const double c[] = {0.99999999999980993,  676.5203681218851,     -1259.1392167224028,
                                   771.32342877765313,   -176.61502916214059,   12.507343278686905,
                                   -0.13857109526572012, 9.9843696540789804e-6, 1.5056327351493116e-7};
        if (z < 0.5) {
            // reflection: Γ(z)Γ(1-z) = π / sin(πz)
            return std::log(kPi / std::sin(kPi * z)) - engLogGamma(1.0 - z);
        }
        z -= 1.0;
        double x = c[0];
        for (int i = 1; i < 9; ++i) x += c[i] / (z + static_cast<double>(i));
        const double t = z + 7.5;
        return 0.5 * std::log(2.0 * kPi) + (z + 0.5) * std::log(t) - t + std::log(x);
    }

    double engGammaImpl(double z)
    {
        if (z <= 0.0 && std::floor(z) == z) return std::numeric_limits<double>::quiet_NaN();    // poles
        if (z < 0.5) return kPi / (std::sin(kPi * z) * engGammaImpl(1.0 - z));
        return std::exp(engLogGamma(z));
    }

}    // namespace

// =============================================================================
// Bitwise
// =============================================================================

namespace OpenXLSX::Formula
{
    XLCellValue fnBitand(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      a = 0, b = 0;
        if (!engTwo(args, a, b, err)) return err;
        int64_t ia = 0, ib = 0;
        if (!engBitInt(a, ia, err) || !engBitInt(b, ib, err)) return err;
        return XLCellValue(static_cast<double>(ia & ib));
    }

    XLCellValue fnBitor(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      a = 0, b = 0;
        if (!engTwo(args, a, b, err)) return err;
        int64_t ia = 0, ib = 0;
        if (!engBitInt(a, ia, err) || !engBitInt(b, ib, err)) return err;
        return XLCellValue(static_cast<double>(ia | ib));
    }

    XLCellValue fnBitxor(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      a = 0, b = 0;
        if (!engTwo(args, a, b, err)) return err;
        int64_t ia = 0, ib = 0;
        if (!engBitInt(a, ia, err) || !engBitInt(b, ib, err)) return err;
        return XLCellValue(static_cast<double>(ia ^ ib));
    }

    XLCellValue fnBitlshift(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      a = 0, sh = 0;
        if (!engTwo(args, a, sh, err)) return err;
        int64_t ia = 0;
        if (!engBitInt(a, ia, err)) return err;
        int shift = static_cast<int>(std::trunc(sh));
        if (shift < 0) return errNum();
        if (shift > 53) return errNum();
        // Excel: result must stay within 48-bit non-negative domain
        if (shift >= 48) return (ia == 0) ? XLCellValue(0.0) : errNum();
        int64_t r = ia << shift;
        if (r < 0 || r > kBitMax) return errNum();
        return XLCellValue(static_cast<double>(r));
    }

    XLCellValue fnBitrshift(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      a = 0, sh = 0;
        if (!engTwo(args, a, sh, err)) return err;
        int64_t ia = 0;
        if (!engBitInt(a, ia, err)) return err;
        int shift = static_cast<int>(std::trunc(sh));
        if (shift < 0) return errNum();
        if (shift >= 48) return XLCellValue(0.0);
        return XLCellValue(static_cast<double>(static_cast<uint64_t>(ia) >> shift));
    }

    // =============================================================================
    // Base conversion
    // =============================================================================

    XLCellValue fnHex2dec(const std::vector<XLFormulaArg>& args)
    {
        if (args.empty() || args[0].size() == 0) return errValue();
        auto v = args[0].asScalar();
        if (isError(v)) return v;
        return engFromBase(toString(v), 16, 10);
    }

    XLCellValue fnDec2hex(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      n = 0;
        if (!engOne(args, n, err)) return err;
        int places = 0;
        if (args.size() > 1 && args[1].size() > 0) {
            auto p = coerceArgScalar(args[1], true);
            if (isError(p)) return p;
            places = static_cast<int>(std::trunc(toDouble(p)));
        }
        return engToBase(n, 16, places, 10);
    }

    XLCellValue fnBin2dec(const std::vector<XLFormulaArg>& args)
    {
        if (args.empty() || args[0].size() == 0) return errValue();
        auto v = args[0].asScalar();
        if (isError(v)) return v;
        return engFromBase(toString(v), 2, 10);
    }

    XLCellValue fnDec2bin(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      n = 0;
        if (!engOne(args, n, err)) return err;
        int places = 0;
        if (args.size() > 1 && args[1].size() > 0) {
            auto p = coerceArgScalar(args[1], true);
            if (isError(p)) return p;
            places = static_cast<int>(std::trunc(toDouble(p)));
        }
        return engToBase(n, 2, places, 10);
    }

    XLCellValue fnOct2dec(const std::vector<XLFormulaArg>& args)
    {
        if (args.empty() || args[0].size() == 0) return errValue();
        auto v = args[0].asScalar();
        if (isError(v)) return v;
        return engFromBase(toString(v), 8, 10);
    }

    XLCellValue fnDec2oct(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      n = 0;
        if (!engOne(args, n, err)) return err;
        int places = 0;
        if (args.size() > 1 && args[1].size() > 0) {
            auto p = coerceArgScalar(args[1], true);
            if (isError(p)) return p;
            places = static_cast<int>(std::trunc(toDouble(p)));
        }
        return engToBase(n, 8, places, 10);
    }

    XLCellValue fnHex2bin(const std::vector<XLFormulaArg>& args)
    {
        auto d = fnHex2dec(args);
        if (isError(d)) return d;
        std::vector<XLFormulaArg> a{XLFormulaArg(d)};
        if (args.size() > 1) a.push_back(args[1]);
        return fnDec2bin(a);
    }

    XLCellValue fnBin2hex(const std::vector<XLFormulaArg>& args)
    {
        auto d = fnBin2dec(args);
        if (isError(d)) return d;
        std::vector<XLFormulaArg> a{XLFormulaArg(d)};
        if (args.size() > 1) a.push_back(args[1]);
        return fnDec2hex(a);
    }

    // =============================================================================
    // Complex
    // =============================================================================

    XLCellValue fnComplex(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      re = 0, im = 0;
        if (!engTwo(args, re, im, err, true)) return err;
        char suffix = 'i';
        if (args.size() > 2 && args[2].size() > 0) {
            auto s = toString(args[2].asScalar());
            if (s == "j" || s == "J") suffix = 'j';
            else if (s == "i" || s == "I")
                suffix = 'i';
            else
                return errValue();
        }
        return XLCellValue(engFormatComplex({re, im}, suffix));
    }

    XLCellValue fnImreal(const std::vector<XLFormulaArg>& args)
    {
        if (args.empty() || args[0].size() == 0) return errValue();
        EngComplex z;
        auto       v = args[0].asScalar();
        if (isError(v)) return v;
        if (!engParseComplex(v, z)) return errNum();
        return XLCellValue(z.re);
    }

    XLCellValue fnImaginary(const std::vector<XLFormulaArg>& args)
    {
        if (args.empty() || args[0].size() == 0) return errValue();
        EngComplex z;
        auto       v = args[0].asScalar();
        if (isError(v)) return v;
        if (!engParseComplex(v, z)) return errNum();
        return XLCellValue(z.im);
    }

    XLCellValue fnImabs(const std::vector<XLFormulaArg>& args)
    {
        if (args.empty() || args[0].size() == 0) return errValue();
        EngComplex z;
        auto       v = args[0].asScalar();
        if (isError(v)) return v;
        if (!engParseComplex(v, z)) return errNum();
        return XLCellValue(std::hypot(z.re, z.im));
    }

    XLCellValue fnImargument(const std::vector<XLFormulaArg>& args)
    {
        if (args.empty() || args[0].size() == 0) return errValue();
        EngComplex z;
        auto       v = args[0].asScalar();
        if (isError(v)) return v;
        if (!engParseComplex(v, z)) return errNum();
        if (z.re == 0.0 && z.im == 0.0) return errDiv0();
        return XLCellValue(std::atan2(z.im, z.re));
    }

    XLCellValue fnImconjugate(const std::vector<XLFormulaArg>& args)
    {
        if (args.empty() || args[0].size() == 0) return errValue();
        EngComplex z;
        auto       v = args[0].asScalar();
        if (isError(v)) return v;
        if (!engParseComplex(v, z)) return errNum();
        return XLCellValue(engFormatComplex({z.re, -z.im}));
    }

    XLCellValue fnImsum(const std::vector<XLFormulaArg>& args)
    {
        EngComplex sum{};
        for (const auto& a : args) {
            auto mat = a.materialize();
            for (std::size_t i = 0; i < mat.size(); ++i) {
                auto v = mat[i];
                if (isError(v)) return v;
                if (isEmpty(v)) continue;
                EngComplex z;
                if (!engParseComplex(v, z)) return errNum();
                sum.re += z.re;
                sum.im += z.im;
            }
        }
        return XLCellValue(engFormatComplex(sum));
    }

    XLCellValue fnImsub(const std::vector<XLFormulaArg>& args)
    {
        if (args.size() < 2) return errValue();
        EngComplex a, b;
        auto       va = args[0].asScalar();
        auto       vb = args[1].asScalar();
        if (isError(va)) return va;
        if (isError(vb)) return vb;
        if (!engParseComplex(va, a) || !engParseComplex(vb, b)) return errNum();
        return XLCellValue(engFormatComplex({a.re - b.re, a.im - b.im}));
    }

    XLCellValue fnImproduct(const std::vector<XLFormulaArg>& args)
    {
        EngComplex prod{1.0, 0.0};
        bool       any = false;
        for (const auto& a : args) {
            auto mat = a.materialize();
            for (std::size_t i = 0; i < mat.size(); ++i) {
                auto v = mat[i];
                if (isError(v)) return v;
                if (isEmpty(v)) continue;
                EngComplex z;
                if (!engParseComplex(v, z)) return errNum();
                EngComplex n{prod.re * z.re - prod.im * z.im, prod.re * z.im + prod.im * z.re};
                prod = n;
                any  = true;
            }
        }
        if (!any) return errValue();
        return XLCellValue(engFormatComplex(prod));
    }

    XLCellValue fnImdiv(const std::vector<XLFormulaArg>& args)
    {
        if (args.size() < 2) return errValue();
        EngComplex a, b;
        auto       va = args[0].asScalar();
        auto       vb = args[1].asScalar();
        if (isError(va)) return va;
        if (isError(vb)) return vb;
        if (!engParseComplex(va, a) || !engParseComplex(vb, b)) return errNum();
        const double den = b.re * b.re + b.im * b.im;
        if (den == 0.0) return errNum();
        return XLCellValue(engFormatComplex({(a.re * b.re + a.im * b.im) / den, (a.im * b.re - a.re * b.im) / den}));
    }

    // =============================================================================
    // Special functions
    // =============================================================================

    XLCellValue fnErf(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      x = 0;
        if (!engOne(args, x, err)) return err;
        if (args.size() >= 2 && args[1].size() > 0) {
            auto y = coerceArgScalar(args[1], false);
            if (isError(y)) return y;
            // ERF(x,y) = erf(y)-erf(x)
            return XLCellValue(std::erf(toDouble(y)) - std::erf(x));
        }
        return XLCellValue(std::erf(x));
    }

    XLCellValue fnErfc(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      x = 0;
        if (!engOne(args, x, err)) return err;
        return XLCellValue(std::erfc(x));
    }

    XLCellValue fnGamma(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      x = 0;
        if (!engOne(args, x, err)) return err;
        double g = engGammaImpl(x);
        if (!std::isfinite(g)) return errNum();
        return XLCellValue(g);
    }

    XLCellValue fnGammaln(const std::vector<XLFormulaArg>& args)
    {
        XLCellValue err;
        double      x = 0;
        if (!engOne(args, x, err)) return err;
        if (x <= 0.0) return errNum();
        double g = engLogGamma(x);
        if (!std::isfinite(g)) return errNum();
        return XLCellValue(g);
    }

}    // namespace OpenXLSX::Formula

void OpenXLSX::registerEngineeringFunctions(XLFormulaRegistry& registry)
{
    using namespace OpenXLSX::Formula;
    registry.registerFunction("BITAND", std::make_shared<XLSimpleFormulaFunction>(fnBitand));
    registry.registerFunction("BITOR", std::make_shared<XLSimpleFormulaFunction>(fnBitor));
    registry.registerFunction("BITXOR", std::make_shared<XLSimpleFormulaFunction>(fnBitxor));
    registry.registerFunction("BITLSHIFT", std::make_shared<XLSimpleFormulaFunction>(fnBitlshift));
    registry.registerFunction("BITRSHIFT", std::make_shared<XLSimpleFormulaFunction>(fnBitrshift));

    registry.registerFunction("HEX2DEC", std::make_shared<XLSimpleFormulaFunction>(fnHex2dec));
    registry.registerFunction("DEC2HEX", std::make_shared<XLSimpleFormulaFunction>(fnDec2hex));
    registry.registerFunction("BIN2DEC", std::make_shared<XLSimpleFormulaFunction>(fnBin2dec));
    registry.registerFunction("DEC2BIN", std::make_shared<XLSimpleFormulaFunction>(fnDec2bin));
    registry.registerFunction("OCT2DEC", std::make_shared<XLSimpleFormulaFunction>(fnOct2dec));
    registry.registerFunction("DEC2OCT", std::make_shared<XLSimpleFormulaFunction>(fnDec2oct));
    registry.registerFunction("HEX2BIN", std::make_shared<XLSimpleFormulaFunction>(fnHex2bin));
    registry.registerFunction("BIN2HEX", std::make_shared<XLSimpleFormulaFunction>(fnBin2hex));

    registry.registerFunction("COMPLEX", std::make_shared<XLSimpleFormulaFunction>(fnComplex));
    registry.registerFunction("IMREAL", std::make_shared<XLSimpleFormulaFunction>(fnImreal));
    registry.registerFunction("IMAGINARY", std::make_shared<XLSimpleFormulaFunction>(fnImaginary));
    registry.registerFunction("IMABS", std::make_shared<XLSimpleFormulaFunction>(fnImabs));
    registry.registerFunction("IMARGUMENT", std::make_shared<XLSimpleFormulaFunction>(fnImargument));
    registry.registerFunction("IMCONJUGATE", std::make_shared<XLSimpleFormulaFunction>(fnImconjugate));
    registry.registerFunction("IMSUM", std::make_shared<XLSimpleFormulaFunction>(fnImsum));
    registry.registerFunction("IMSUB", std::make_shared<XLSimpleFormulaFunction>(fnImsub));
    registry.registerFunction("IMPRODUCT", std::make_shared<XLSimpleFormulaFunction>(fnImproduct));
    registry.registerFunction("IMDIV", std::make_shared<XLSimpleFormulaFunction>(fnImdiv));

    registry.registerFunction("ERF", std::make_shared<XLSimpleFormulaFunction>(fnErf));
    registry.registerFunction("ERFC", std::make_shared<XLSimpleFormulaFunction>(fnErfc));
    registry.registerFunction("GAMMA", std::make_shared<XLSimpleFormulaFunction>(fnGamma));
    registry.registerFunction("GAMMALN", std::make_shared<XLSimpleFormulaFunction>(fnGammaln));
}

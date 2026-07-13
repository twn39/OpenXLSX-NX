#include "XLFormulaUtils.hpp"

#include <fmt/format.h>
#include <atomic>
#include <cctype>
#include <cmath>
#include <limits>
#include <mutex>

namespace OpenXLSX::formula
{
    namespace
    {
        std::mutex              g_seedMutex;
        std::optional<uint64_t> g_formulaSeed;
        std::atomic<uint64_t>   g_seedGeneration{0};

        struct RngState
        {
            std::mt19937_64 engine;
            uint64_t        gen{std::numeric_limits<uint64_t>::max()};
            bool            seeded{false};
        };

        RngState& tlsRng()
        {
            thread_local RngState st;
            const uint64_t        gen = g_seedGeneration.load(std::memory_order_acquire);
            if (st.gen != gen) {
                std::optional<uint64_t> seed;
                {
                    std::lock_guard<std::mutex> lock(g_seedMutex);
                    seed = g_formulaSeed;
                }
                if (seed) {
                    st.engine.seed(*seed);
                    st.seeded = true;
                }
                else {
                    st.engine.seed(std::random_device{}());
                    st.seeded = false;
                }
                st.gen = gen;
            }
            return st;
        }
    }    // namespace

    void setFormulaRandomSeed(uint64_t seed)
    {
        {
            std::lock_guard<std::mutex> lock(g_seedMutex);
            g_formulaSeed = seed;
        }
        g_seedGeneration.fetch_add(1, std::memory_order_release);
    }

    void clearFormulaRandomSeed()
    {
        {
            std::lock_guard<std::mutex> lock(g_seedMutex);
            g_formulaSeed.reset();
        }
        g_seedGeneration.fetch_add(1, std::memory_order_release);
    }

    std::optional<uint64_t> formulaRandomSeed()
    {
        std::lock_guard<std::mutex> lock(g_seedMutex);
        return g_formulaSeed;
    }

    std::mt19937_64& formulaRng() { return tlsRng().engine; }

    double toDouble(const XLCellValue& v)
    {
        switch (v.type()) {
            case XLValueType::Integer:
                return static_cast<double>(v.get<int64_t>());
            case XLValueType::Float:
                return v.get<double>();
            case XLValueType::Boolean:
                return v.get<bool>() ? 1.0 : 0.0;
            default:
                return std::numeric_limits<double>::quiet_NaN();
        }
    }

    bool isNumeric(const XLCellValue& v)
    {
        return v.type() == XLValueType::Integer || v.type() == XLValueType::Float || v.type() == XLValueType::Boolean;
    }

    bool isEmpty(const XLCellValue& v) { return v.type() == XLValueType::Empty; }

    bool isError(const XLCellValue& v) { return v.type() == XLValueType::Error; }

    std::string toString(const XLCellValue& v)
    {
        switch (v.type()) {
            case XLValueType::String:
                return v.get<std::string>();
            case XLValueType::Integer:
                return fmt::format("{}", v.get<int64_t>());
            case XLValueType::Float:
                return fmt::format("{}", v.get<double>());
            case XLValueType::Boolean:
                return v.get<bool>() ? "TRUE" : "FALSE";
            default:
                return "";
        }
    }

    XLCellValue makeError(std::string_view token)
    {
        XLCellValue r;
        r.setError(std::string(token));
        return r;
    }

    std::string strTrim(std::string s)
    {
        const char* ws = " \t\r\n";
        const auto  b  = s.find_first_not_of(ws);
        if (b == std::string::npos) return {};
        const auto e = s.find_last_not_of(ws);
        return s.substr(b, e - b + 1);
    }

    bool tryParseNumericString(std::string_view s, double& out)
    {
        // Trim ASCII whitespace
        while (!s.empty() && std::isspace(static_cast<unsigned char>(s.front()))) s.remove_prefix(1);
        while (!s.empty() && std::isspace(static_cast<unsigned char>(s.back()))) s.remove_suffix(1);
        if (s.empty()) return false;

        // Reject pure boolean words (Excel does not coerce "TRUE" in arithmetic the same way as boolean type)
        if (s.size() == 4 || s.size() == 5) {
            std::string upper(s);
            for (char& c : upper) c = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
            if (upper == "TRUE" || upper == "FALSE") return false;
        }

        try {
            std::size_t idx = 0;
            std::string tmp(s);
            double      d = std::stod(tmp, &idx);
            // Allow trailing whitespace only (already trimmed) — require full consume
            if (idx != tmp.size()) return false;
            if (!std::isfinite(d)) return false;
            out = d;
            return true;
        }
        catch (...) {
            return false;
        }
    }

    XLCellValue coerceToNumber(const XLCellValue& v, bool blankAsZero)
    {
        if (isError(v)) return v;
        if (isEmpty(v)) return blankAsZero ? XLCellValue(0.0) : errValue();
        if (isNumeric(v)) return XLCellValue(toDouble(v));
        if (v.type() == XLValueType::String) {
            double d = 0.0;
            if (tryParseNumericString(v.get<std::string>(), d)) return XLCellValue(d);
        }
        return errValue();
    }

    XLCellValue coerceArgScalar(const XLFormulaArg& arg, bool blankAsZero)
    {
        // Missing argument (no cells) → #VALUE!. A blank cell still has size 1 and coerces via blankAsZero.
        if (arg.size() == 0) return errValue();
        return coerceToNumber(arg.asScalar(), blankAsZero);
    }

    XLCellValue firstError(const XLFormulaArg& arg)
    {
        const auto mat = arg.materialize();
        for (std::size_t i = 0; i < mat.size(); ++i) {
            const auto v = mat[i];
            if (isError(v)) return v;
        }
        return XLCellValue{};
    }

    XLCellValue firstError(const std::vector<XLFormulaArg>& args)
    {
        for (const auto& a : args) {
            auto e = firstError(a);
            if (isError(e)) return e;
        }
        return XLCellValue{};
    }

    std::vector<double> numerics(const std::vector<XLFormulaArg>& args)
    {
        std::vector<double> out;
        for (const auto& arg : args)
            for (const auto& v : arg)
                if (isNumeric(v)) out.push_back(toDouble(v));
        return out;
    }

    std::vector<double> numerics(const XLFormulaArg& arg)
    {
        std::vector<double> out;
        for (const auto& v : arg)
            if (isNumeric(v)) out.push_back(toDouble(v));
        return out;
    }

    std::vector<double> numericsOrError(const XLFormulaArg& arg, XLCellValue& outError)
    {
        outError = XLCellValue{};
        std::vector<double> out;
        const auto          mat = arg.materialize();
        for (std::size_t i = 0; i < mat.size(); ++i) {
            const auto v = mat[i];
            if (isError(v)) {
                outError = v;
                return {};
            }
            if (isNumeric(v)) out.push_back(toDouble(v));
        }
        return out;
    }

    std::vector<double> numericsOrError(const std::vector<XLFormulaArg>& args, XLCellValue& outError)
    {
        outError = XLCellValue{};
        std::vector<double> out;
        for (const auto& arg : args) {
            const auto mat = arg.materialize();
            for (std::size_t i = 0; i < mat.size(); ++i) {
                const auto v = mat[i];
                if (isError(v)) {
                    outError = v;
                    return {};
                }
                if (isNumeric(v)) out.push_back(toDouble(v));
            }
        }
        return out;
    }

}    // namespace OpenXLSX::formula

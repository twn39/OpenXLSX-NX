#include "XLFormulaUtils.hpp"

#include <fmt/format.h>
#include <limits>

namespace OpenXLSX::formula
{
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

}    // namespace OpenXLSX::formula

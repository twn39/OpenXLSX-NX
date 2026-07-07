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


XLCellValue Formula::fnConcatenate(const std::vector<XLFormulaArg>& args)
{
    std::string result;
    for (const auto& arg : args) {
        for (const auto& v : arg) { result += toString(v); }
    }
    return XLCellValue(result);
}

XLCellValue Formula::fnLen(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    return XLCellValue(static_cast<int64_t>(toString(args[0][0]).size()));
}

XLCellValue Formula::fnLeft(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    int64_t     n = (args.size() > 1 && !args[1].empty()) ? static_cast<int64_t>(toDouble(args[1][0])) : 1;
    if (n < 0) return errValue();
    return XLCellValue(s.substr(0, static_cast<std::size_t>(n)));
}

XLCellValue Formula::fnRight(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    int64_t     n = (args.size() > 1 && !args[1].empty()) ? static_cast<int64_t>(toDouble(args[1][0])) : 1;
    if (n < 0) return errValue();
    auto len   = static_cast<int64_t>(s.size());
    auto start = len - n;
    if (start < 0) start = 0;
    return XLCellValue(s.substr(static_cast<std::size_t>(start)));
}

XLCellValue Formula::fnMid(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    std::string s     = toString(args[0][0]);
    int64_t     start = static_cast<int64_t>(toDouble(args[1][0])) - 1;    // 1-based
    int64_t     count = static_cast<int64_t>(toDouble(args[2][0]));
    if (start < 0 || count < 0) return errValue();
    return XLCellValue(s.substr(static_cast<std::size_t>(start), static_cast<std::size_t>(count)));
}

XLCellValue Formula::fnUpper(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    std::transform(s.begin(), s.end(), s.begin(), ::toupper);
    return XLCellValue(s);
}

XLCellValue Formula::fnLower(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    std::transform(s.begin(), s.end(), s.begin(), ::tolower);
    return XLCellValue(s);
}

XLCellValue Formula::fnTrim(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    return XLCellValue(strTrim(toString(args[0][0])));
}

XLCellValue Formula::fnText(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();

    try {
        // 1. Get the value to format
        auto value = args[0][0];

        // 2. Get the format string - don't strip literal quotes if they are inside the string
        // Actually toString(args[1][0]) will return `#,##0.00;(#,##0.00);"-";*@*` because
        // string arguments in formulas don't have the outer quotes.
        // Wait, excel format strings can contain literal double quotes inside them, e.g. \"-\".
        // Let's ensure formatStr is correctly retrieved.
        std::string formatStr = toString(args[1][0]);

        // 3. Instantiate the format engine (parses format on construction)
        XLNumberFormatter formatter(formatStr);

        // 4. Execute formatting and return text cell value
        std::string formattedStr = formatter.format(value);
        return XLCellValue(formattedStr);
    }
    catch (...) {
        // Return #VALUE! if format string is highly invalid or other exceptions
        return errValue();
    }
}

XLCellValue Formula::fnFind(const std::vector<XLFormulaArg>& args)
{
    // FIND(find_text, within_text, [start_num=1]) – case-sensitive, no wildcards
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::string needle   = toString(args[0][0]);
    std::string haystack = toString(args[1][0]);
    int         startPos = (args.size() > 2 && !args[2].empty()) ? static_cast<int>(toDouble(args[2][0])) - 1 : 0;
    if (startPos < 0 || startPos >= static_cast<int>(haystack.size())) return errValue();
    auto pos = haystack.find(needle, static_cast<std::size_t>(startPos));
    if (pos == std::string::npos) return errValue();
    return XLCellValue(static_cast<int64_t>(pos + 1));
}

XLCellValue Formula::fnSearch(const std::vector<XLFormulaArg>& args)
{
    // SEARCH(find_text, within_text, [start_num=1]) – case-insensitive
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::string needle   = toString(args[0][0]);
    std::string haystack = toString(args[1][0]);
    std::transform(needle.begin(), needle.end(), needle.begin(), ::tolower);
    std::transform(haystack.begin(), haystack.end(), haystack.begin(), ::tolower);
    int startPos = (args.size() > 2 && !args[2].empty()) ? static_cast<int>(toDouble(args[2][0])) - 1 : 0;
    if (startPos < 0 || startPos >= static_cast<int>(haystack.size())) return errValue();
    auto pos = haystack.find(needle, static_cast<std::size_t>(startPos));
    if (pos == std::string::npos) return errValue();
    return XLCellValue(static_cast<int64_t>(pos + 1));
}

XLCellValue Formula::fnSubstitute(const std::vector<XLFormulaArg>& args)
{
    // SUBSTITUTE(text, old_text, new_text, [instance_num])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    std::string s       = toString(args[0][0]);
    std::string oldStr  = toString(args[1][0]);
    std::string newStr  = toString(args[2][0]);
    int         instNum = (args.size() > 3 && !args[3].empty()) ? static_cast<int>(toDouble(args[3][0])) : 0;

    if (oldStr.empty()) return XLCellValue(s);
    std::string result;
    result.reserve(s.size());
    std::size_t pos        = 0;
    int         occurrence = 0;
    while (true) {
        auto found = s.find(oldStr, pos);
        if (found == std::string::npos) {
            result.append(s, pos, std::string::npos);
            break;
        }
        ++occurrence;
        result.append(s, pos, found - pos);
        if (instNum == 0 || occurrence == instNum)
            result += newStr;
        else
            result.append(oldStr);
        pos = found + oldStr.size();
    }
    return XLCellValue(result);
}

XLCellValue Formula::fnReplace(const std::vector<XLFormulaArg>& args)
{
    // REPLACE(old_text, start_num, num_chars, new_text)
    if (args.size() < 4 || args[0].empty() || args[1].empty() || args[2].empty() || args[3].empty()) return errValue();
    std::string s      = toString(args[0][0]);
    int         start  = static_cast<int>(toDouble(args[1][0])) - 1;    // convert to 0-based
    int         nchars = static_cast<int>(toDouble(args[2][0]));
    std::string newStr = toString(args[3][0]);
    if (start < 0) start = 0;
    auto startSz  = static_cast<std::size_t>(start);
    auto ncharsSz = static_cast<std::size_t>(std::max(0, nchars));
    if (startSz > s.size()) startSz = s.size();
    s.replace(startSz, ncharsSz, newStr);
    return XLCellValue(s);
}

XLCellValue Formula::fnRept(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::string s = toString(args[0][0]);
    int         n = static_cast<int>(toDouble(args[1][0]));
    if (n < 0) return errValue();
    std::string result;
    result.reserve(s.size() * static_cast<std::size_t>(n));
    for (int i = 0; i < n; ++i) result += s;
    return XLCellValue(result);
}

XLCellValue Formula::fnExact(const std::vector<XLFormulaArg>& args)
{
    // EXACT(text1, text2) – case-sensitive comparison
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    return XLCellValue(toString(args[0][0]) == toString(args[1][0]));
}

XLCellValue Formula::fnT(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return XLCellValue(std::string(""));
    const auto& v = args[0][0];
    return v.type() == XLValueType::String ? v : XLCellValue(std::string(""));
}

XLCellValue Formula::fnValue(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    if (isNumeric(args[0][0])) return XLCellValue(toDouble(args[0][0]));
    std::string s = toString(args[0][0]);
    try {
        std::size_t idx = 0;
        double      v   = std::stod(s, &idx);
        if (idx == s.size()) return XLCellValue(v);
    }
    catch (...) {
    }
    return errValue();
}

XLCellValue Formula::fnTextjoin(const std::vector<XLFormulaArg>& args)
{
    // TEXTJOIN(delimiter, ignore_empty, text1, text2, ...)
    if (args.size() < 3 || args[0].empty() || args[1].empty()) return errValue();
    std::string delim       = toString(args[0][0]);
    bool        ignoreEmpty = static_cast<bool>(toDouble(args[1][0]));
    std::string result;
    bool        first = true;
    for (std::size_t i = 2; i < args.size(); ++i) {
        for (const auto& v : args[i]) {
            std::string s = toString(v);
            if (ignoreEmpty && s.empty()) continue;
            if (!first) result += delim;
            result += s;
            first = false;
        }
    }
    return XLCellValue(result);
}

XLCellValue Formula::fnClean(const std::vector<XLFormulaArg>& args)
{
    // CLEAN – remove non-printable characters (< 0x20)
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    s.erase(std::remove_if(s.begin(), s.end(), [](unsigned char c) { return c < 0x20; }), s.end());
    return XLCellValue(s);
}

XLCellValue Formula::fnProper(const std::vector<XLFormulaArg>& args)
{
    // PROPER – Title Case
    if (args.empty() || args[0].empty()) return errValue();
    std::string s       = toString(args[0][0]);
    bool        capNext = true;
    for (char& c : s) {
        if (std::isspace(static_cast<unsigned char>(c))) { capNext = true; }
        else if (capNext) {
            c       = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
            capNext = false;
        }
        else {
            c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
        }
    }
    return XLCellValue(s);
}

XLCellValue Formula::fnChar(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double dval = toDouble(args[0][0]);
    if (std::isnan(dval)) return errValue();
    int code = static_cast<int>(dval);
    if (code < 1 || code > 255) return errValue();
    std::string s(1, static_cast<char>(code));
    return XLCellValue(s);
}

XLCellValue Formula::fnUnichar(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    double dval = toDouble(args[0][0]);
    if (std::isnan(dval)) return errValue();
    int code = static_cast<int>(dval);
    if (code <= 0 || code > 0x10FFFF) return errValue();    // Valid Unicode range

    std::string s;
    if (code <= 0x7F) { s.push_back(static_cast<char>(code)); }
    else if (code <= 0x7FF) {
        s.push_back(static_cast<char>(0xC0 | (code >> 6)));
        s.push_back(static_cast<char>(0x80 | (code & 0x3F)));
    }
    else if (code <= 0xFFFF) {
        s.push_back(static_cast<char>(0xE0 | (code >> 12)));
        s.push_back(static_cast<char>(0x80 | ((code >> 6) & 0x3F)));
        s.push_back(static_cast<char>(0x80 | (code & 0x3F)));
    }
    else {
        s.push_back(static_cast<char>(0xF0 | (code >> 18)));
        s.push_back(static_cast<char>(0x80 | ((code >> 12) & 0x3F)));
        s.push_back(static_cast<char>(0x80 | ((code >> 6) & 0x3F)));
        s.push_back(static_cast<char>(0x80 | (code & 0x3F)));
    }
    return XLCellValue(s);
}

XLCellValue Formula::fnCode(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    if (s.empty()) return errValue();
    // Excel's CODE function typically just looks at the first single byte or system charset
    // Here we maintain basic behavior for first byte
    return XLCellValue(static_cast<int64_t>(static_cast<unsigned char>(s[0])));
}

XLCellValue Formula::fnUnicode(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty()) return errValue();
    std::string s = toString(args[0][0]);
    if (s.empty()) return errValue();

    const unsigned char* str = reinterpret_cast<const unsigned char*>(s.c_str());
    int                  cp  = 0;
    if (str[0] < 0x80) { cp = str[0]; }
    else if ((str[0] & 0xE0) == 0xC0) {
        if (s.length() < 2) return errValue();
        cp = ((str[0] & 0x1F) << 6) | (str[1] & 0x3F);
    }
    else if ((str[0] & 0xF0) == 0xE0) {
        if (s.length() < 3) return errValue();
        cp = ((str[0] & 0x0F) << 12) | ((str[1] & 0x3F) << 6) | (str[2] & 0x3F);
    }
    else if ((str[0] & 0xF8) == 0xF0) {
        if (s.length() < 4) return errValue();
        cp = ((str[0] & 0x07) << 18) | ((str[1] & 0x3F) << 12) | ((str[2] & 0x3F) << 6) | (str[3] & 0x3F);
    }
    else {
        return errValue();    // Invalid UTF-8
    }

    return XLCellValue(static_cast<int64_t>(cp));
}

void OpenXLSX::registerTextFunctions(XLFormulaRegistry& registry)
{
    registry.registerFunction("CONCATENATE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnConcatenate));
    registry.registerFunction("LEN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnLen));
    registry.registerFunction("LEFT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnLeft));
    registry.registerFunction("RIGHT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRight));
    registry.registerFunction("MID", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMid));
    registry.registerFunction("UPPER", std::make_shared<XLSimpleFormulaFunction>(Formula::fnUpper));
    registry.registerFunction("LOWER", std::make_shared<XLSimpleFormulaFunction>(Formula::fnLower));
    registry.registerFunction("TRIM", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTrim));
    registry.registerFunction("TEXT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnText));
    registry.registerFunction("FIND", std::make_shared<XLSimpleFormulaFunction>(Formula::fnFind));
    registry.registerFunction("SEARCH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSearch));
    registry.registerFunction("SUBSTITUTE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSubstitute));
    registry.registerFunction("REPLACE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnReplace));
    registry.registerFunction("REPT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRept));
    registry.registerFunction("EXACT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnExact));
    registry.registerFunction("T", std::make_shared<XLSimpleFormulaFunction>(Formula::fnT));
    registry.registerFunction("VALUE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnValue));
    registry.registerFunction("TEXTJOIN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTextjoin));
    registry.registerFunction("CLEAN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnClean));
    registry.registerFunction("PROPER", std::make_shared<XLSimpleFormulaFunction>(Formula::fnProper));
    registry.registerFunction("CHAR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnChar));
    registry.registerFunction("UNICHAR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnUnichar));
    registry.registerFunction("CODE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnCode));
    registry.registerFunction("UNICODE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnUnicode));
}

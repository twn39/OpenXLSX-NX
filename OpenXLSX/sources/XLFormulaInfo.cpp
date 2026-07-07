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

XLCellValue Formula::fnVlookup(const std::vector<XLFormulaArg>& args)
{
    // VLOOKUP(lookup_value, table_array, col_index, [exact_match])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();

    const XLCellValue& lookupVal = args[0][0];
    const auto&        table     = args[1];
    int                colIdx    = static_cast<int>(toDouble(args[2][0]));    // 1-based
    bool               exact     = (args.size() < 4 || args[3].empty()) ? true : (toDouble(args[3][0]) == 0.0);

    if (colIdx < 1) return errValue();

    int nCols = static_cast<int>(table.cols());
    if (colIdx > nCols) return errRef();

    int nRows = static_cast<int>(table.rows());

    // Linear scan column 0 of each row
    for (int r = 0; r < nRows; ++r) {
        int cellIdx = r * nCols;
        if (cellIdx >= static_cast<int>(table.size())) break;
        const XLCellValue& key = table[static_cast<std::size_t>(cellIdx)];

        bool match = false;
        if (exact) {
            if (isNumeric(key) && isNumeric(lookupVal))
                match = (toDouble(key) == toDouble(lookupVal));
            else
                match = (toString(key) == toString(lookupVal));
        }
        else {
            // Approximate: find largest key <= lookupVal
            if (isNumeric(key) && isNumeric(lookupVal)) match = (toDouble(key) <= toDouble(lookupVal));
        }

        if (match) {
            int resultIdx = r * nCols + (colIdx - 1);
            if (resultIdx < static_cast<int>(table.size())) return table[static_cast<std::size_t>(resultIdx)];
            return errRef();
        }
    }
    return errNA();
}

XLCellValue Formula::fnHlookup(const std::vector<XLFormulaArg>& args)
{
    // HLOOKUP(lookup_value, table_array, row_index, [exact_match])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();

    const XLCellValue& lookupVal = args[0][0];
    const auto&        table     = args[1];
    int                rowIdx    = static_cast<int>(toDouble(args[2][0]));    // 1-based row to return
    bool               exact     = (args.size() < 4 || args[3].empty()) ? true : (toDouble(args[3][0]) == 0.0);

    if (rowIdx < 1) return errValue();

    int nRows = static_cast<int>(table.rows());
    if (rowIdx > nRows) return errRef();

    int nCols = static_cast<int>(table.cols());

    // HLOOKUP searches the first row.
    for (int c = 0; c < nCols; ++c) {
        if (c >= static_cast<int>(table.size())) break;
        const XLCellValue& key   = table[static_cast<std::size_t>(c)];
        bool               match = false;
        if (exact) {
            if (isNumeric(key) && isNumeric(lookupVal))
                match = (toDouble(key) == toDouble(lookupVal));
            else
                match = (toString(key) == toString(lookupVal));
        }
        else {
            if (isNumeric(key) && isNumeric(lookupVal)) match = (toDouble(key) <= toDouble(lookupVal));
        }
        if (match) {
            std::size_t resultIdx = static_cast<std::size_t>(c + (rowIdx - 1) * nCols);
            if (resultIdx < table.size()) return table[resultIdx];
            return errRef();
        }
    }
    return errNA();
}

XLCellValue Formula::fnXlookup(const std::vector<XLFormulaArg>& args)
{
    // XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();

    const auto& lookupVal = args[0][0];
    const auto& lookupArr = args[1];
    const auto& returnArr = args[2];

    int matchMode  = (args.size() > 4 && !args[4].empty()) ? static_cast<int>(toDouble(args[4][0])) : 0;
    int searchMode = (args.size() > 5 && !args[5].empty()) ? static_cast<int>(toDouble(args[5][0])) : 1;

    int bestMatchIdx = -1;

    auto compareValues = [matchMode](const XLCellValue& key, const XLCellValue& val) -> double {
        if (isNumeric(key) && isNumeric(val)) { return toDouble(key) - toDouble(val); }
        else {
            // For exact matches, if types mismatch entirely, don't loosely convert them unless they're both strings
            if (matchMode == 0 && (isNumeric(key) != isNumeric(val))) {
                return std::numeric_limits<double>::infinity();    // Force failure
            }

            std::string sKey = toString(key);
            std::string sVal = toString(val);
            std::transform(sKey.begin(), sKey.end(), sKey.begin(), ::tolower);
            std::transform(sVal.begin(), sVal.end(), sVal.begin(), ::tolower);
            int cmp = sKey.compare(sVal);
            return cmp < 0 ? -1.0 : (cmp > 0 ? 1.0 : 0.0);
        }
    };

    if (searchMode == 2 || searchMode == -2) {
        // Binary search (2 = ascending, -2 = descending)
        int low  = 0;
        int high = static_cast<int>(lookupArr.size()) - 1;

        while (low <= high) {
            int         mid  = low + (high - low) / 2;
            const auto& key  = lookupArr[static_cast<std::size_t>(mid)];
            double      diff = compareValues(key, lookupVal);

            if (diff == 0.0) {
                bestMatchIdx = mid;
                break;    // Exact match found
            }

            if (searchMode == 2) {                              // Ascending array
                if (diff < 0) {                                 // key < lookupVal
                    if (matchMode == -1) bestMatchIdx = mid;    // Candidate for exact or next smaller
                    low = mid + 1;
                }
                else {                                         // key > lookupVal
                    if (matchMode == 1) bestMatchIdx = mid;    // Candidate for exact or next larger
                    high = mid - 1;
                }
            }
            else {                                              // Descending array (searchMode == -2)
                if (diff < 0) {                                 // key < lookupVal
                    if (matchMode == -1) bestMatchIdx = mid;    // Candidate for exact or next smaller
                    high = mid - 1;                             // smaller values are to the left in descending
                }
                else {                                         // key > lookupVal
                    if (matchMode == 1) bestMatchIdx = mid;    // Candidate for exact or next larger
                    low = mid + 1;                             // larger values are to the left, so smaller-than-current are right
                }
            }
        }
    }
    else {
        // Linear search (1 = first-to-last, -1 = last-to-first)
        int startIdx = (searchMode == -1) ? static_cast<int>(lookupArr.size()) - 1 : 0;
        int endIdx   = (searchMode == -1) ? -1 : static_cast<int>(lookupArr.size());
        int step     = (searchMode == -1) ? -1 : 1;

        double bestDiff = std::numeric_limits<double>::infinity();

        for (int i = startIdx; i != endIdx; i += step) {
            const auto& key = lookupArr[static_cast<std::size_t>(i)];

            if (matchMode == 0) {    // Exact match
                if (compareValues(key, lookupVal) == 0.0) {
                    bestMatchIdx = i;
                    break;
                }
            }
            else if (matchMode == -1) {                         // Exact match or next smaller item
                double diff = compareValues(lookupVal, key);    // lookupVal - key
                if (diff == 0.0) {
                    bestMatchIdx = i;
                    break;
                }
                else if (diff > 0 && diff < bestDiff) {
                    bestDiff     = diff;
                    bestMatchIdx = i;
                }
            }
            else if (matchMode == 1) {                          // Exact match or next larger item
                double diff = compareValues(key, lookupVal);    // key - lookupVal
                if (diff == 0.0) {
                    bestMatchIdx = i;
                    break;
                }
                else if (diff > 0 && diff < bestDiff) {
                    bestDiff     = diff;
                    bestMatchIdx = i;
                }
            }
            // wildcard match_mode (2) is not fully implemented here yet
        }
    }

    if (bestMatchIdx != -1) {
        if (static_cast<std::size_t>(bestMatchIdx) < returnArr.size()) { return returnArr[static_cast<std::size_t>(bestMatchIdx)]; }
        return errRef();    // Return array smaller than lookup array
    }

    // Not found
    if (args.size() > 3 && !args[3].empty()) { return args[3][0]; }
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
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    const XLCellValue& lookupVal = args[0][0];
    const auto&        arr       = args[1];
    int                matchType = (args.size() > 2 && !args[2].empty()) ? static_cast<int>(toDouble(args[2][0])) : 1;

    if (matchType == 0) {
        // Exact match
        for (std::size_t i = 0; i < arr.size(); ++i) {
            bool match = false;
            if (isNumeric(arr[i]) && isNumeric(lookupVal))
                match = (toDouble(arr[i]) == toDouble(lookupVal));
            else
                match = (toString(arr[i]) == toString(lookupVal));
            if (match) return XLCellValue(static_cast<int64_t>(i + 1));
        }
        return errNA();
    }
    // Approximate match (matchType == 1: largest <= lookup)
    int64_t bestIdx = -1;
    double  bestVal = std::numeric_limits<double>::lowest();
    for (std::size_t i = 0; i < arr.size(); ++i) {
        if (isNumeric(arr[i]) && isNumeric(lookupVal)) {
            double d = toDouble(arr[i]);
            if (d <= toDouble(lookupVal) && d >= bestVal) {
                bestVal = d;
                bestIdx = static_cast<int64_t>(i + 1);
            }
        }
    }
    if (bestIdx < 0) return errNA();
    return XLCellValue(bestIdx);
}

XLCellValue Formula::fnRow(const std::vector<XLFormulaArg>& args)
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

XLCellValue Formula::fnColumn(const std::vector<XLFormulaArg>& args)
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
    registry.registerFunction("ROW", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRow));
    registry.registerFunction("COLUMN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnColumn));
    registry.registerFunction("DATEDIF", std::make_shared<XLSimpleFormulaFunction>(Formula::fnDatedif));
    registry.registerFunction("ISEVEN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIseven));
    registry.registerFunction("ISODD", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIsodd));
}

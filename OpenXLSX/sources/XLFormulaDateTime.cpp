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
    std::tm date_serialToTm(double serial)
    {
        try {
            return XLDateTime(serial).tm();
        }
        catch (...) {
            return {};
        }
    }

    double date_tmToSerial(std::tm t)
    {
        try {
            return XLDateTime(t).serial();
        }
        catch (...) {
            return 0.0;
        }
    }

    int date_daysInMonth(int year, int month)
    {
        constexpr std::array<int, 12> days = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
        bool leap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
        int d = days.at(static_cast<std::size_t>(month - 1));
        if (month == 2 && leap) ++d;
        return d;
    }

    bool date_isWeekend(double serial)
    {
        std::tm t = date_serialToTm(serial);
        return t.tm_wday == 0 || t.tm_wday == 6; // 0=Sun, 6=Sat
    }
} // namespace

XLCellValue Formula::fnToday(const std::vector<XLFormulaArg>&)
{ return XLCellValue(XLDateTime::now().serial() - std::fmod(XLDateTime::now().serial(), 1.0)); }

XLCellValue Formula::fnNow(const std::vector<XLFormulaArg>&) { return XLCellValue(XLDateTime::now().serial()); }

XLCellValue Formula::fnDate(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    std::tm t{};
    t.tm_year = static_cast<int>(toDouble(args[0][0])) - 1900;
    t.tm_mon  = static_cast<int>(toDouble(args[1][0])) - 1;
    t.tm_mday = static_cast<int>(toDouble(args[2][0]));
    return XLCellValue(date_tmToSerial(t));
}

XLCellValue Formula::fnTime(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    if (!isNumeric(args[0][0]) || !isNumeric(args[1][0]) || !isNumeric(args[2][0])) return errValue();

    int64_t h = static_cast<int64_t>(toDouble(args[0][0]));
    int64_t m = static_cast<int64_t>(toDouble(args[1][0]));
    int64_t s = static_cast<int64_t>(toDouble(args[2][0]));

    double fraction = (h * 3600.0 + m * 60.0 + s) / 86400.0;

    // Normalize to 0-1
    fraction = fraction - std::floor(fraction);
    if (fraction < 0) fraction += 1.0;

    return XLCellValue(fraction);
}

XLCellValue Formula::fnYear(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(date_serialToTm(toDouble(args[0][0])).tm_year + 1900));
}

XLCellValue Formula::fnMonth(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(date_serialToTm(toDouble(args[0][0])).tm_mon + 1));
}

XLCellValue Formula::fnDay(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(date_serialToTm(toDouble(args[0][0])).tm_mday));
}

XLCellValue Formula::fnHour(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(date_serialToTm(toDouble(args[0][0])).tm_hour));
}

XLCellValue Formula::fnMinute(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(date_serialToTm(toDouble(args[0][0])).tm_min));
}

XLCellValue Formula::fnSecond(const std::vector<XLFormulaArg>& args)
{
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    return XLCellValue(static_cast<int64_t>(date_serialToTm(toDouble(args[0][0])).tm_sec));
}

XLCellValue Formula::fnDays(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    if (!isNumeric(args[0][0]) || !isNumeric(args[1][0])) return errValue();

    double end_date   = toDouble(args[0][0]);
    double start_date = toDouble(args[1][0]);

    return XLCellValue(end_date - start_date);
}

XLCellValue Formula::fnWeekday(const std::vector<XLFormulaArg>& args)
{
    // WEEKDAY(serial, [return_type])
    // return_type 1 (default): 1=Sunday … 7=Saturday
    // return_type 2: 1=Monday … 7=Sunday
    // return_type 3: 0=Monday … 6=Sunday
    if (args.empty() || args[0].empty() || !isNumeric(args[0][0])) return errValue();
    int mode = (args.size() > 1 && !args[1].empty()) ? static_cast<int>(toDouble(args[1][0])) : 1;
    int wday = date_serialToTm(toDouble(args[0][0])).tm_wday;    // 0=Sun … 6=Sat
    switch (mode) {
        case 2:
            return XLCellValue(static_cast<int64_t>(wday == 0 ? 7 : wday));
        case 3:
            return XLCellValue(static_cast<int64_t>(wday == 0 ? 6 : wday - 1));
        default:
            return XLCellValue(static_cast<int64_t>(wday + 1));
    }
}

XLCellValue Formula::fnEdate(const std::vector<XLFormulaArg>& args)
{
    // EDATE(start_date, months) – same day, N months later
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::tm t      = date_serialToTm(toDouble(args[0][0]));
    int     months = static_cast<int>(toDouble(args[1][0]));
    t.tm_mon += months;
    // Normalise overflow
    t.tm_year += t.tm_mon / 12;
    t.tm_mon = t.tm_mon % 12;
    if (t.tm_mon < 0) {
        t.tm_year--;
        t.tm_mon += 12;
    }
    // Clamp day to month end
    int maxDay = date_daysInMonth(t.tm_year + 1900, t.tm_mon + 1);
    if (t.tm_mday > maxDay) t.tm_mday = maxDay;
    return XLCellValue(date_tmToSerial(t));
}

XLCellValue Formula::fnEomonth(const std::vector<XLFormulaArg>& args)
{
    // EOMONTH(start_date, months) – last day of month, N months later
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    std::tm t      = date_serialToTm(toDouble(args[0][0]));
    int     months = static_cast<int>(toDouble(args[1][0]));
    t.tm_mon += months;
    t.tm_year += t.tm_mon / 12;
    t.tm_mon = t.tm_mon % 12;
    if (t.tm_mon < 0) {
        t.tm_year--;
        t.tm_mon += 12;
    }
    t.tm_mday = date_daysInMonth(t.tm_year + 1900, t.tm_mon + 1);
    return XLCellValue(date_tmToSerial(t));
}

XLCellValue Formula::fnWorkday(const std::vector<XLFormulaArg>& args)
{
    // WORKDAY(start_date, days, [holidays])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double current = std::floor(toDouble(args[0][0]));
    int    days    = static_cast<int>(toDouble(args[1][0]));
    int    step    = (days > 0) ? 1 : -1;

    // Collect holidays (arg[2] if present)
    std::vector<double> holidays;
    if (args.size() > 2) {
        for (const auto& v : args[2])
            if (isNumeric(v)) holidays.push_back(std::floor(toDouble(v)));
    }

    int remaining = std::abs(days);
    while (remaining > 0) {
        current += step;
        if (date_isWeekend(current)) continue;
        bool isHoliday = std::find(holidays.begin(), holidays.end(), current) != holidays.end();
        if (!isHoliday) --remaining;
    }
    return XLCellValue(current);
}

XLCellValue Formula::fnNetworkdays(const std::vector<XLFormulaArg>& args)
{
    // NETWORKDAYS(start_date, end_date, [holidays])
    if (args.size() < 2 || args[0].empty() || args[1].empty()) return errValue();
    double start = std::floor(toDouble(args[0][0]));
    double end   = std::floor(toDouble(args[1][0]));

    std::vector<double> holidays;
    if (args.size() > 2)
        for (const auto& v : args[2])
            if (isNumeric(v)) holidays.push_back(std::floor(toDouble(v)));

    int    count = 0;
    int    sign  = (end >= start) ? 1 : -1;
    double d     = start;
    while ((sign > 0 && d <= end) || (sign < 0 && d >= end)) {
        if (!date_isWeekend(d)) {
            bool isHoliday = std::find(holidays.begin(), holidays.end(), d) != holidays.end();
            if (!isHoliday) ++count;
        }
        d += sign;
    }
    return XLCellValue(static_cast<int64_t>(sign * count));
}

XLCellValue Formula::fnIsoweeknum(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errValue();
    double serial = nums[0];
    if (serial < 0) return errNum();

    auto dayOfWeek = [](double s) {
        const auto day = static_cast<int64_t>(std::floor(s)) % 7;
        return (day == 0 ? 6 : static_cast<int>(day - 1));
    };

    int    wday_mon    = (dayOfWeek(serial) + 6) % 7;    // 0=Mon .. 6=Sun
    double nearest_thu = std::floor(serial) + 3 - wday_mon;

    auto tm_thu = static_cast<std::tm>(XLDateTime(nearest_thu));

    std::tm tm_jan4{};
    tm_jan4.tm_year    = tm_thu.tm_year;
    tm_jan4.tm_mon     = 0;
    tm_jan4.tm_mday    = 4;
    tm_jan4.tm_isdst   = -1;
    double jan4_serial = std::floor(XLDateTime(tm_jan4).serial());

    int    jan4_wday_mon = (dayOfWeek(jan4_serial) + 6) % 7;
    double first_thu     = jan4_serial + 3 - jan4_wday_mon;

    int iso_week = 1 + static_cast<int>(std::round((nearest_thu - first_thu) / 7.0));
    return XLCellValue(static_cast<double>(iso_week));
}

XLCellValue Formula::fnWeeknum(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.empty()) return errValue();
    double serial = nums[0];
    if (serial < 0) return errNum();
    int return_type = (nums.size() > 1) ? static_cast<int>(std::trunc(nums[1])) : 1;

    if (return_type == 21) {
        return fnIsoweeknum(args);    // System 2 (ISO)
    }

    auto dayOfWeek = [](double s) {
        const auto day = static_cast<int64_t>(std::floor(s)) % 7;
        return (day == 0 ? 6 : static_cast<int>(day - 1));
    };

    auto    tm = static_cast<std::tm>(XLDateTime(serial));
    std::tm tm_jan1{};
    tm_jan1.tm_year    = tm.tm_year;
    tm_jan1.tm_mon     = 0;
    tm_jan1.tm_mday    = 1;
    tm_jan1.tm_isdst   = -1;
    double jan1_serial = std::floor(XLDateTime(tm_jan1).serial());

    int days_elapsed = static_cast<int>(std::floor(serial) - jan1_serial);

    if (return_type == 1 || return_type == 17) {     // Week begins on Sunday
        int jan1_wday   = dayOfWeek(jan1_serial);    // 0=Sun..6=Sat
        int days_in_wk1 = 7 - jan1_wday;
        if (days_elapsed < days_in_wk1) return XLCellValue(1.0);
        return XLCellValue(2.0 + std::floor((days_elapsed - days_in_wk1) / 7.0));
    }
    else {                                                       // Week begins on Monday
        int jan1_wday_mon = (dayOfWeek(jan1_serial) + 6) % 7;    // 0=Mon..6=Sun
        int days_in_wk1   = 7 - jan1_wday_mon;
        if (days_elapsed < days_in_wk1) return XLCellValue(1.0);
        return XLCellValue(2.0 + std::floor((days_elapsed - days_in_wk1) / 7.0));
    }
}

XLCellValue Formula::fnDays360(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() < 2) return errValue();
    double d1     = nums[0];
    double d2     = nums[1];
    bool   method = (nums.size() > 2) && nums[2] != 0.0;

    auto tm1 = static_cast<std::tm>(XLDateTime(d1));
    auto tm2 = static_cast<std::tm>(XLDateTime(d2));

    int y1 = tm1.tm_year + 1900, m1 = tm1.tm_mon + 1, day1 = tm1.tm_mday;
    int y2 = tm2.tm_year + 1900, m2 = tm2.tm_mon + 1, day2 = tm2.tm_mday;

    if (method) {    // European
        if (day1 == 31) day1 = 30;
        if (day2 == 31) day2 = 30;
    }
    else {                                       // US (NASD)
        if (m1 == 2 && day1 >= 28) day1 = 30;    // approx for end of Feb
        if (day1 == 31) day1 = 30;
        if (day2 == 31 && day1 >= 30) day2 = 30;
    }

    double res = (y2 - y1) * 360.0 + (m2 - m1) * 30.0 + (day2 - day1);
    return XLCellValue(res);
}

void OpenXLSX::registerDateTimeFunctions(XLFormulaRegistry& registry)
{
    registry.registerFunction("TODAY", std::make_shared<XLSimpleFormulaFunction>(Formula::fnToday));
    registry.registerFunction("NOW", std::make_shared<XLSimpleFormulaFunction>(Formula::fnNow));
    registry.registerFunction("DATE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnDate));
    registry.registerFunction("TIME", std::make_shared<XLSimpleFormulaFunction>(Formula::fnTime));
    registry.registerFunction("YEAR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnYear));
    registry.registerFunction("MONTH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMonth));
    registry.registerFunction("DAY", std::make_shared<XLSimpleFormulaFunction>(Formula::fnDay));
    registry.registerFunction("HOUR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnHour));
    registry.registerFunction("MINUTE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMinute));
    registry.registerFunction("SECOND", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSecond));
    registry.registerFunction("DAYS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnDays));
    registry.registerFunction("WEEKDAY", std::make_shared<XLSimpleFormulaFunction>(Formula::fnWeekday));
    registry.registerFunction("EDATE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnEdate));
    registry.registerFunction("EOMONTH", std::make_shared<XLSimpleFormulaFunction>(Formula::fnEomonth));
    registry.registerFunction("WORKDAY", std::make_shared<XLSimpleFormulaFunction>(Formula::fnWorkday));
    registry.registerFunction("NETWORKDAYS", std::make_shared<XLSimpleFormulaFunction>(Formula::fnNetworkdays));
    registry.registerFunction("ISOWEEKNUM", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIsoweeknum));
    registry.registerFunction("WEEKNUM", std::make_shared<XLSimpleFormulaFunction>(Formula::fnWeeknum));
    registry.registerFunction("DAYS360", std::make_shared<XLSimpleFormulaFunction>(Formula::fnDays360));
}

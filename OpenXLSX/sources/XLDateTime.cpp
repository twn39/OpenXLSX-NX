#include "XLDateTime.hpp"
#include "XLException.hpp"

#include <array>
#include <cmath>
#include <cstdint>
#include <gsl/gsl>
#include <iomanip>
#include <sstream>

namespace
{
    /**
     * @brief Determines if a year should be treated as a leap year.
     * @details This function implements the standard Gregorian rules, but explicitly
     * includes 1900 to ensure compatibility with Excel's historical bug.
     */
    [[nodiscard]] constexpr bool isLeapYear(int year) noexcept
    {
        if (year == 1900) return true;
        return (year % 400 == 0) || (year % 4 == 0 && year % 100 != 0);
    }

    /**
     * @brief Provides the day count for a specific month.
     * @details Uses a lookup table for efficiency, with a special case for February
     * in leap years.
     */
    [[nodiscard]] constexpr int daysInMonth(int month, int year) noexcept
    {
        constexpr std::array<int, 12> monthDays = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
        if (month < 1 || month > 12) return 0;
        if (month == 2 && isLeapYear(year)) return 29;
        return monthDays[gsl::narrow_cast<size_t>(month - 1)];
    }

    /**
     * @brief Map Excel serial numbers to days of the week.
     * @details Excel uses a 1-7 range (Sunday-Saturday) and inherits the 1900 leap bug
     * for its epoch starting point.
     */
    [[nodiscard]] int dayOfWeek(double serial) noexcept
    {
        const auto day = gsl::narrow_cast<int32_t>(serial) % 7;
        return (day == 0 ? 6 : day - 1);
    }
}    // namespace

namespace OpenXLSX
{
    XLDateTime::XLDateTime() = default;

    XLDateTime::XLDateTime(double serial) : m_serial(serial)
    {
        if (serial < 0.0) { throw XLDateTimeError("Excel date/time serial number is invalid (must be >= 0.0)"); }
    }

    /**
     * @details Converts a standard tm struct to an Excel serial number.
     * The conversion process starts from the 1900 epoch and accounts for every
     * full year and month before adding the current day and fractional time.
     */
    XLDateTime::XLDateTime(const std::tm& timepoint)
    {
        if (timepoint.tm_year < 0) throw XLDateTimeError("Invalid year. Must be >= 0.");
        if (timepoint.tm_mon < 0 || timepoint.tm_mon > 11) throw XLDateTimeError("Invalid month. Must be 0-11.");

        int days = daysInMonth(timepoint.tm_mon + 1, timepoint.tm_year + 1900);
        if (timepoint.tm_mday <= 0 || timepoint.tm_mday > days) { throw XLDateTimeError("Invalid day for the given month."); }

        m_serial = 1.0;

        for (int i = 0; i < timepoint.tm_year; ++i) { m_serial += (isLeapYear(1900 + i) ? 366 : 365); }

        for (int i = 0; i < timepoint.tm_mon; ++i) { m_serial += daysInMonth(i + 1, timepoint.tm_year + 1900); }

        m_serial += timepoint.tm_mday - 1;

        const int32_t seconds = timepoint.tm_hour * 3600 + timepoint.tm_min * 60 + timepoint.tm_sec;
        m_serial += seconds / 86400.0;
    }

    /**
     * @details Unix-to-Excel conversion uses a fixed offset of 25569 days
     * (the difference between the two epochs).
     */
    XLDateTime::XLDateTime(time_t unixtime) { m_serial = (static_cast<double>(unixtime) / 86400.0) + 25569.0; }

    XLDateTime::XLDateTime(const std::chrono::system_clock::time_point& timepoint)
        : XLDateTime(std::chrono::system_clock::to_time_t(timepoint))
    {}

    XLDateTime XLDateTime::now() { return XLDateTime(std::chrono::system_clock::now()); }

    XLDateTime XLDateTime::fromString(const std::string& dateString, const std::string& format)
    {
        std::tm            tm = {};
        std::istringstream ss(dateString);
        ss >> std::get_time(&tm, format.c_str());
        if (ss.fail()) throw XLDateTimeError("Failed to parse date string: " + dateString);
        return XLDateTime(tm);
    }

    std::string XLDateTime::toString(const std::string& format) const
    {
        std::tm            timepoint = tm();
        std::ostringstream ss;
        ss << std::put_time(&timepoint, format.c_str());
        return ss.str();
    }

    XLDateTime& XLDateTime::operator=(double serial)
    {
        if (serial < 1.0) throw XLDateTimeError("Invalid serial number");
        m_serial = serial;
        return *this;
    }

    XLDateTime& XLDateTime::operator=(const std::tm& timepoint)
    {
        *this = XLDateTime(timepoint);
        return *this;
    }

    XLDateTime::operator std::tm() const { return tm(); }

    std::chrono::system_clock::time_point XLDateTime::chrono() const
    {
        auto unixtime = gsl::narrow_cast<time_t>(std::round((m_serial - 25569.0) * 86400.0));
        return std::chrono::system_clock::from_time_t(unixtime);
    }

    double XLDateTime::serial() const { return m_serial; }

    /**
     * @details Decomposition of the Excel serial number.
     * The process handles the 1900 bug by treating 1900 as a leap year, ensuring
     * that all subsequent dates are correctly aligned with Excel's calendar.
     * Time components are extracted through fractional remainders, with an
     * overflow check to handle precision rounding issues that could push
     * a timestamp into the next second or day.
     */
    std::tm XLDateTime::tm() const
    {
        std::tm result{};
        result.tm_isdst = -1;
        double serial   = m_serial;

        int year = 0;
        while (serial > 1.0) {
            const int days = (isLeapYear(year + 1900) ? 366 : 365);
            if (static_cast<double>(days) + 1.0 > serial) break;
            serial -= days;
            ++year;
        }
        result.tm_year = year;
        result.tm_wday = dayOfWeek(m_serial);

        int month = 0;
        while (month < 11) {
            int days = daysInMonth(month + 1, 1900 + year);
            if (static_cast<double>(days) + 1.0 > serial) break;
            serial -= days;
            ++month;
        }
        result.tm_mon = month;

        result.tm_mday = gsl::narrow_cast<int>(serial);
        serial -= result.tm_mday;

        int yday = result.tm_mday - 1;
        for (int i = 0; i < month; ++i) { yday += daysInMonth(i + 1, 1900 + year); }
        result.tm_yday = yday;

        result.tm_hour = gsl::narrow_cast<int>(serial * 24);
        serial -= (result.tm_hour / 24.0);

        result.tm_min = gsl::narrow_cast<int>(serial * 1440);
        serial -= (result.tm_min / 1440.0);

        result.tm_sec = gsl::narrow_cast<int>(std::lround(serial * 86400));

        // Pass rounded overflows back up the date hierarchy to prevent
        // "60 seconds" or "24 hours" timestamps.
        if (result.tm_sec >= 60) {
            result.tm_sec -= 60;
            if (++result.tm_min >= 60) {
                result.tm_min -= 60;
                if (++result.tm_hour >= 24) {
                    result.tm_hour -= 24;
                    if (++result.tm_mday > daysInMonth(result.tm_mon + 1, result.tm_year + 1900)) {
                        result.tm_mday = 1;
                        if (++result.tm_mon > 11) {
                            result.tm_mon = 0;
                            ++result.tm_year;
                        }
                    }
                }
            }
        }

        return result;
    }
}    // namespace OpenXLSX

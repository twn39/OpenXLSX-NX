/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLDATETIME_HPP
#define OPENXLSX_XLDATETIME_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

// ===== External Includes ===== //
#include <chrono>
#include <ctime>
#include <string>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLException.hpp"

namespace OpenXLSX
{
    /**
     * @brief Manages date and time values according to the Excel 1900 date system.
     * @details Excel represents dates as floating-point numbers where 1.0 is 1900-01-01.
     * This class handles the historical "1900 Leap Year Bug" (Excel incorrectly treats 1900
     * as a leap year) to ensure absolute compatibility with Microsoft Excel files.
     * All string inputs and outputs are in the ISO 8601-like format (YYYY-MM-DD HH:MM:SS)
     * unless a custom format is specified.
     */
    class OPENXLSX_EXPORT XLDateTime
    {
    public:
        /**
         * @brief Constructs an object representing the Excel epoch (1900-01-01).
         */
        XLDateTime();

        /**
         * @param serial Excel's floating-point date representation. Must be >= 1.0.
         * @throws XLDateTimeError if serial is less than 1.0 (Excel doesn't support earlier dates).
         */
        explicit XLDateTime(double serial);

        /**
         * @param timepoint A standard C tm struct. Only date and time fields are used.
         * @throws XLDateTimeError if the date is invalid or before 1900.
         */
        explicit XLDateTime(const std::tm& timepoint);

        /**
         * @param unixtime Seconds since the Unix epoch (1970-01-01 00:00:00 UTC).
         */
        explicit XLDateTime(time_t unixtime);

        /**
         * @param timepoint A high-precision system clock time point.
         */
        explicit XLDateTime(const std::chrono::system_clock::time_point& timepoint);

        // Rule of Zero: All copy/move operations are handled by the compiler
        // as the class only contains a single primitive double (m_serial).
        XLDateTime(const XLDateTime& other)                = default;
        XLDateTime(XLDateTime&& other) noexcept            = default;
        ~XLDateTime()                                      = default;
        XLDateTime& operator=(const XLDateTime& other)     = default;
        XLDateTime& operator=(XLDateTime&& other) noexcept = default;

        /**
         * @return An XLDateTime object initialized to the current system time.
         */
        [[nodiscard]] static XLDateTime now();

        /**
         * @brief Parses a string into an Excel date.
         * @param dateString The source string (e.g., "2024-03-16 10:00:00").
         * @param format Standard strptime-compatible format string.
         * @throws XLDateTimeError if parsing fails.
         */
        [[nodiscard]] static XLDateTime fromString(const std::string& dateString, const std::string& format = "%Y-%m-%d %H:%M:%S");

        /**
         * @brief Formats the date into a human-readable string.
         * @param format Standard strftime-compatible format string.
         */
        [[nodiscard]] std::string toString(const std::string& format = "%Y-%m-%d %H:%M:%S") const;

        /**
         * @brief Updates the serial number and validates its range.
         */
        XLDateTime& operator=(double serial);

        /**
         * @brief Updates the date from a tm struct.
         */
        XLDateTime& operator=(const std::tm& timepoint);

        /**
         * @brief Implicitly converts the date to a floating-point serial number.
         * @details This allows seamless integration with worksheet cell values.
         */
        template<typename T, typename = std::enable_if_t<std::is_floating_point_v<T>>>
        operator T() const    // NOLINT
        { return static_cast<T>(serial()); }

        /**
         * @brief Explicitly converts the Excel date to a standard C tm struct.
         */
        operator std::tm() const;    // NOLINT

        /**
         * @brief Converts the Excel date to a C++ system_clock::time_point.
         */
        [[nodiscard]] std::chrono::system_clock::time_point chrono() const;

        /**
         * @return The raw Excel serial number (days since 1899-12-30).
         */
        [[nodiscard]] double serial() const;

        /**
         * @brief Decomposes the Excel serial number into its date and time components.
         * @details This method handles the 1900 leap year bug and rounding overflows
         * to provide an accurate representation of the time.
         */
        [[nodiscard]] std::tm tm() const;

    private:
        double m_serial{1.0}; /**< Excel's internal representation. 1.0 = 1900-01-01. */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLDATETIME_HPP

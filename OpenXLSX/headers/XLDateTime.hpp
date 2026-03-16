#ifndef OPENXLSX_XLDATETIME_HPP
#define OPENXLSX_XLDATETIME_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <chrono>
#include <ctime>
#include <string>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLException.hpp"

// ========== CLASS AND ENUM TYPE DEFINITIONS ========== //
namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLDateTime
    {
    public:
        /**
         * @brief Constructor.
         */
        XLDateTime();

        /**
         * @brief Constructor taking an Excel time point serial number as an argument.
         * @param serial Excel time point serial number.
         */
        explicit XLDateTime(double serial);

        /**
         * @brief Constructor taking a std::tm struct as an argument.
         * @param timepoint A std::tm struct.
         */
        explicit XLDateTime(const std::tm& timepoint);

        /**
         * @brief Constructor taking a unixtime format (seconds since 1/1/1970) as an argument.
         * @param unixtime A time_t number.
         */
        explicit XLDateTime(time_t unixtime);

        /**
         * @brief Constructor taking a std::chrono::system_clock::time_point as an argument.
         * @param timepoint A chrono time_point.
         */
        explicit XLDateTime(const std::chrono::system_clock::time_point& timepoint);

        /**
         * @brief Default copy constructor
         */
        XLDateTime(const XLDateTime& other) = default;

        /**
         * @brief Default move constructor
         */
        XLDateTime(XLDateTime&& other) noexcept = default;

        /**
         * @brief Default destructor
         */
        ~XLDateTime() = default;

        /**
         * @brief Default copy assignment operator
         */
        XLDateTime& operator=(const XLDateTime& other) = default;

        /**
         * @brief Default move assignment operator
         */
        XLDateTime& operator=(XLDateTime&& other) noexcept = default;

        /**
         * @brief Get the current time as an XLDateTime object.
         * @return An XLDateTime object representing the current time.
         */
        [[nodiscard]] static XLDateTime now();

        /**
         * @brief Create an XLDateTime object from a string.
         * @param dateString The string representation of the date/time.
         * @param format The format string (default: ISO 8601 "%Y-%m-%d %H:%M:%S").
         * @return An XLDateTime object.
         */
        [[nodiscard]] static XLDateTime fromString(const std::string& dateString, const std::string& format = "%Y-%m-%d %H:%M:%S");

        /**
         * @brief Convert the XLDateTime object to a string.
         * @param format The format string (default: ISO 8601 "%Y-%m-%d %H:%M:%S").
         * @return A string representation of the date/time.
         */
        [[nodiscard]] std::string toString(const std::string& format = "%Y-%m-%d %H:%M:%S") const;

        /**
         * @brief Assignment operator taking an Excel date/time serial number as an argument.
         * @param serial A floating point value with the serial number.
         * @return Reference to the copied-to object.
         */
        XLDateTime& operator=(double serial);

        /**
         * @brief Assignment operator taking a std::tm object as an argument.
         * @param timepoint std::tm object with the time point
         * @return Reference to the copied-to object.
         */
        XLDateTime& operator=(const std::tm& timepoint);

        /**
         * @brief Implicit conversion to Excel date/time serial number (any floating point type).
         * @tparam T Type to convert to (any floating point type).
         * @return Excel date/time serial number.
         */
        template<typename T, typename = std::enable_if_t<std::is_floating_point_v<T>>>
        operator T() const    // NOLINT
        { return static_cast<T>(serial()); }

        /**
         * @brief Implicit conversion to std::tm object.
         * @return std::tm object.
         */
        operator std::tm() const;    // NOLINT

        /**
         * @brief Conversion to std::chrono::system_clock::time_point.
         * @return A chrono time_point object.
         */
        [[nodiscard]] std::chrono::system_clock::time_point chrono() const;

        /**
         * @brief Get the date/time in the form of an Excel date/time serial number.
         * @return A double with the serial number.
         */
        [[nodiscard]] double serial() const;

        /**
         * @brief Get the date/time in the form of a std::tm struct.
         * @return A std::tm struct with the time point.
         */
        [[nodiscard]] std::tm tm() const;

    private:
        double m_serial{1.0}; /**<  */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLDATETIME_HPP

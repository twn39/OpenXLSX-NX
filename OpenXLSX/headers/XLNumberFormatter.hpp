/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#pragma once

#include <string>
#include <vector>
#include "XLCellValue.hpp"

namespace OpenXLSX {

    // TokenType for formatting instructions
    enum class XLFormatTokenType {
        Literal,      // Literal string/character (e.g., "-", "/", " ", text)
        Year,         // yy or yyyy
        Month,        // m, mm, mmm, mmmm, mmmmm
        Day,          // d, dd, ddd, dddd
        Hour,         // h, hh
        Minute,       // m, mm (resolved via context)
        Second,       // s, ss
        AMPM,         // AM/PM, A/P
        DigitZero,    // 0 (forced digit, zero padded)
        DigitOpt,     // # (optional digit)
        Decimal,      // . (decimal point)
        Thousands,    // , (thousands separator)
        Percent,      // % (multiply value by 100)
        TextPlaceholder // @ (text placeholder)
    };

    // FormatToken represents a parsed chunk of the format string
    struct XLFormatToken {
        XLFormatTokenType type;
        std::string value;    // Holds the original placeholder string or literal text
        int count;            // Number of characters in the placeholder (e.g., "yyyy" = 4)
    };

    // FormatSection represents a single section of a multi-section format string
    struct FormatSection {
        std::vector<XLFormatToken> tokens;
        bool isDateTime{false};
        bool isNumeric{false};
        bool hasPercent{false};
        int decimalPlaces{0};
    };

    // XLNumberFormatter parses an Excel number format string and applies it to an XLCellValue
    class XLNumberFormatter {
    public:
        /**
         * @brief Constructor that pre-compiles (parses) the format string
         * @param formatString Excel format string (e.g., "yyyy-mm-dd hh:mm:ss", "#,##0.00;[Red](#,##0.00)")
         */
        explicit XLNumberFormatter(const std::string& formatString);

        /**
         * @brief Applies the format to a given cell value
         * @param value The value to format
         * @return The formatted string representation
         */
        std::string format(const XLCellValue& value) const;

    private:
        std::string m_formatString;
        std::vector<FormatSection> m_sections;

        // Core parsing phase
        void parse();
        void parseSection(const std::string& sectionStr);

        // Formatting subroutines
        std::string formatDateTime(double excelDate, const FormatSection& section) const;
        std::string formatNumeric(double number, const FormatSection& section, bool addMinusSign) const;
        std::string formatText(const std::string& text, const FormatSection& section) const;
    };

} // namespace OpenXLSX
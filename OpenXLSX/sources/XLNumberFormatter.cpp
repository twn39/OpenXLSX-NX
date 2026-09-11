/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#include "XLNumberFormatter.hpp"

#include <algorithm>
#include <cctype>
#include <cmath>
#include <iomanip>
#include <sstream>
#include <fmt/format.h>
#include "XLDateTime.hpp"
#include "XLUtilities.hpp"

namespace OpenXLSX {

    XLNumberFormatter::XLNumberFormatter(const std::string& formatString) : m_formatString(formatString) {
        parse();
    }

    std::string XLNumberFormatter::format(const XLCellValue& value) const {
        if (value.type() == XLValueType::Empty) {
            return "";
        }
        
        if (value.type() == XLValueType::Boolean) {
            return value.get<bool>() ? "TRUE" : "FALSE";
        }

        if (value.type() == XLValueType::String) {
            std::string strVal = value.get<std::string>();
            if (m_sections.size() == 4) {
                return formatText(strVal, m_sections[3]);
            }
            if (m_sections.size() > 0) {
                // If there's no specific text section, we just output the string without formatting
                // Actually, if a format string has < 4 sections but contains an '@', it applies to text too.
                // Let's check the FIRST section that contains a TextPlaceholder.
                for (const auto& section : m_sections) {
                    for (const auto& token : section.tokens) {
                        if (token.type == XLFormatTokenType::TextPlaceholder) {
                            return formatText(strVal, section);
                        }
                    }
                }
            }
            return strVal;
        }

        if (value.type() == XLValueType::Integer || value.type() == XLValueType::Float) {
            double number = 0.0;
            if (value.type() == XLValueType::Integer) number = static_cast<double>(value.get<int64_t>());
            else number = value.get<double>();

            int sectionIndex = 0;
            bool addMinusSign = false;

            if (number < 0.0 && m_sections.size() >= 2) {
                sectionIndex = 1;
            } else if (number == 0.0 && m_sections.size() >= 3) {
                sectionIndex = 2;
            } else if (number < 0.0 && m_sections.size() == 1) {
                // If only 1 section is provided, Excel automatically prepends a minus sign for negative numbers.
                addMinusSign = true;
            }

            if (sectionIndex >= static_cast<int>(m_sections.size())) {
                sectionIndex = 0;
            }

            if (m_sections.empty()) {
                 // Should theoretically never happen as we always insert at least one section
                 if (addMinusSign) return "-" + fmt::format("{}", number);
                 return fmt::format("{}", number);
            }

            const auto& section = m_sections[sectionIndex];
            double evalNumber = (number < 0.0 && m_sections.size() >= 2) ? std::abs(number) : number;
            if (addMinusSign) {
                evalNumber = std::abs(number);
            }

            if (section.isDateTime) {
                return formatDateTime(evalNumber, section);
            } else if (section.isNumeric) {
                return formatNumeric(evalNumber, section, addMinusSign);
            } else if (!section.tokens.empty()) {
                // If it has tokens but isn't explicitly numeric (e.g. general mixed with text)
                // Just use numeric format. Text-only strings with @ get handled if string cell value
                return formatNumeric(evalNumber, section, addMinusSign);
            } else {
                // General/Fallback formatting
                if (addMinusSign) return "-" + fmt::format("{}", evalNumber);
                return fmt::format("{}", evalNumber);
            }
        }
        
        if (value.type() == XLValueType::Error) {
            return value.get<std::string>();
        }

        return "";
    }

    void XLNumberFormatter::parse() {
        if (m_formatString.empty() || m_formatString == "General") {
            return;
        }

        std::vector<std::string> sectionStrings;
        std::string currentSection;
        bool inQuotes = false;
        bool escaped = false;

        for (size_t i = 0; i < m_formatString.length(); ++i) {
            char c = m_formatString[i];

            if (escaped) {
                currentSection += c;
                escaped = false;
                continue;
            }

            if (c == '\\') {
                // Keep the backslash so parseSection can process it
                currentSection += c;
                escaped = true;
                continue;
            }

            if (c == '"') {
                currentSection += c;
                inQuotes = !inQuotes;
                continue;
            }

            if (c == ';' && !inQuotes) {
                sectionStrings.push_back(currentSection);
                currentSection.clear();
            } else {
                currentSection += c;
            }
        }
        sectionStrings.push_back(currentSection);

        // When evaluating the text function, if a format string contains multiple sections,
        // we parse all of them.
        for (const auto& sectionStr : sectionStrings) {
            parseSection(sectionStr);
        }
    }

    void XLNumberFormatter::parseSection(const std::string& sectionStr) {
        FormatSection section;
        
        std::string lowerFmt = sectionStr;
        std::transform(lowerFmt.begin(), lowerFmt.end(), lowerFmt.begin(), ::tolower);

        size_t len = lowerFmt.length();
        size_t i = 0;

        while (i < len) {
            char c = lowerFmt[i];
            char origC = sectionStr[i];

            // 1. Handle Quotes (Literal strings)
            if (c == '"') {
                std::string literal;
                i++;
                while (i < len && lowerFmt[i] != '"') {
                    literal += sectionStr[i];
                    i++;
                }
                if (i < len) i++; // Skip closing quote
                section.tokens.push_back({XLFormatTokenType::Literal, literal, static_cast<int>(literal.length())});
                continue;
            }

            // 2. Handle Escapes
            if (c == '\\') {
                if (i + 1 < len) {
                    i++;
                    std::string literal(1, sectionStr[i]);
                    section.tokens.push_back({XLFormatTokenType::Literal, literal, 1});
                }
                i++;
                continue;
            }
            
            // Handle asterisks (fill character). We'll just treat it as a literal for MVP.
            if (c == '*') {
                if (i + 1 < len) {
                    // Excel syntax: *x means fill cell with x. For MVP, we just output the char itself once.
                    // Actually, if we just eat the next character, we might eat the @ placeholder.
                    // Better to just push the literal of the next character, but what if the next character IS @?
                    // To be safe for MVP, let's just insert the literal and NOT advance i++ so the next char is parsed as whatever it is? No, if we do *x, the x shouldn't be parsed as a command if it is one. 
                    // Let's just do what excel does: output the next char literally, then skip it.
                    i++;
                    std::string literal(1, sectionStr[i]);
                    // If it was *@, we should NOT treat it as literal if we want @ to be replaced.
                    // Actually, *x means fill with x. *@ means fill with @. So it IS a literal @.
                    // But wait, if the format is just *@*, the first * means fill with @, and the second * means what?
                    section.tokens.push_back({XLFormatTokenType::Literal, literal, 1});
                }
                i++;
                continue;
            }
            
            // Handle underscore (skip width of next char). We'll just treat as literal space for MVP.
            if (c == '_') {
                if (i + 1 < len) {
                    i++;
                    section.tokens.push_back({XLFormatTokenType::Literal, " ", 1});
                }
                i++;
                continue;
            }

            // 3. Handle Brackets (Skip conditions/colors for TEXT function)
            if (c == '[') {
                i++;
                while (i < len && lowerFmt[i] != ']') {
                    i++;
                }
                if (i < len) i++;
                continue;
            }

            // 4. Handle Date/Time place holders
            if (c == 'y' || c == 'm' || c == 'd' || c == 'h' || c == 's') {
                section.isDateTime = true;
                int count = 0;
                while (i < len && lowerFmt[i] == c) {
                    count++;
                    i++;
                }
                
                XLFormatTokenType type = XLFormatTokenType::Literal;
                if (c == 'y') type = XLFormatTokenType::Year;
                else if (c == 'm') type = XLFormatTokenType::Month; // Ambiguous, will resolve later
                else if (c == 'd') type = XLFormatTokenType::Day;
                else if (c == 'h') type = XLFormatTokenType::Hour;
                else if (c == 's') type = XLFormatTokenType::Second;

                section.tokens.push_back({type, std::string(count, c), count});
                continue;
            }

            // 5. Handle AM/PM
            if (c == 'a' && i + 5 <= len && lowerFmt.substr(i, 5) == "am/pm") {
                section.isDateTime = true;
                section.tokens.push_back({XLFormatTokenType::AMPM, sectionStr.substr(i, 5), 5});
                i += 5;
                continue;
            }
            if (c == 'a' && i + 3 <= len && lowerFmt.substr(i, 3) == "a/p") {
                section.isDateTime = true;
                section.tokens.push_back({XLFormatTokenType::AMPM, sectionStr.substr(i, 3), 3});
                i += 3;
                continue;
            }

            // 6. Handle Numerics
            if (c == '0' || c == '#' || c == '.' || c == ',' || c == '%') {
                section.isNumeric = true;
                if (c == '0') section.tokens.push_back({XLFormatTokenType::DigitZero, "0", 1});
                else if (c == '#') section.tokens.push_back({XLFormatTokenType::DigitOpt, "#", 1});
                else if (c == '.') section.tokens.push_back({XLFormatTokenType::Decimal, ".", 1});
                else if (c == ',') section.tokens.push_back({XLFormatTokenType::Thousands, ",", 1});
                else if (c == '%') {
                    section.hasPercent = true;
                    section.tokens.push_back({XLFormatTokenType::Percent, "%", 1});
                }
                
                if (c == '.') {
                    // Count decimal places
                    size_t j = i + 1;
                    while (j < len && (lowerFmt[j] == '0' || lowerFmt[j] == '#')) {
                        section.decimalPlaces++;
                        j++;
                    }
                }
                i++;
                continue;
            }

            // 7. Handle Text place holders
            if (c == '@') {
                section.tokens.push_back({XLFormatTokenType::TextPlaceholder, "@", 1});
                i++;
                continue;
            }

            // Treat anything else as literal
            std::string literal(1, origC);
            section.tokens.push_back({XLFormatTokenType::Literal, literal, 1});
            i++;
        }

        // Pass 2: Resolve 'm' ambiguity (Month vs Minute)
        for (size_t k = 0; k < section.tokens.size(); ++k) {
            if (section.tokens[k].type == XLFormatTokenType::Month) {
                bool isMinute = false;
                
                // Look back
                for (int b = static_cast<int>(k) - 1; b >= 0; --b) {
                    if (section.tokens[b].type == XLFormatTokenType::Hour) {
                        isMinute = true;
                        break;
                    }
                    if (section.tokens[b].type != XLFormatTokenType::Literal) {
                        break;
                    }
                }
                
                // Look ahead
                if (!isMinute) {
                    for (size_t f = k + 1; f < section.tokens.size(); ++f) {
                        if (section.tokens[f].type == XLFormatTokenType::Second) {
                            isMinute = true;
                            break;
                        }
                        if (section.tokens[f].type != XLFormatTokenType::Literal) {
                            break;
                        }
                    }
                }

                if (isMinute) {
                    section.tokens[k].type = XLFormatTokenType::Minute;
                }
            }
        }
        
        m_sections.push_back(section);
    }

    std::string XLNumberFormatter::formatDateTime(double excelDate, const FormatSection& section) const {
        if (excelDate < 0.0) return "################################################################################################################################################################################################################################################################"; 
        
        try {
            XLDateTime dt(excelDate);
            std::tm t = dt.tm();
            
            bool isAMPM = false;
            for (const auto& token : section.tokens) {
                if (token.type == XLFormatTokenType::AMPM) {
                    isAMPM = true;
                    break;
                }
            }

            int hour12 = t.tm_hour % 12;
            if (hour12 == 0) hour12 = 12;

            std::ostringstream oss;
            
            for (const auto& token : section.tokens) {
                switch (token.type) {
                    case XLFormatTokenType::Literal:
                        oss << token.value;
                        break;
                    case XLFormatTokenType::Year:
                        if (token.count <= 2) oss << fmt::format("{:02d}", (t.tm_year + 1900) % 100);
                        else oss << fmt::format("{:04d}", t.tm_year + 1900);
                        break;
                    case XLFormatTokenType::Month:
                        if (token.count == 1) oss << fmt::format("{}", t.tm_mon + 1);
                        else if (token.count == 2) oss << fmt::format("{:02d}", t.tm_mon + 1);
                        else if (token.count == 3) {
                            const char* months[] = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
                            oss << months[t.tm_mon];
                        }
                        else {
                            const char* months[] = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"};
                            oss << months[t.tm_mon];
                        }
                        break;
                    case XLFormatTokenType::Day:
                        if (token.count == 1) oss << fmt::format("{}", t.tm_mday);
                        else if (token.count == 2) oss << fmt::format("{:02d}", t.tm_mday);
                        else if (token.count == 3) {
                            const char* days[] = {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"};
                            oss << days[t.tm_wday];
                        }
                        else {
                            const char* days[] = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};
                            oss << days[t.tm_wday];
                        }
                        break;
                    case XLFormatTokenType::Hour:
                        if (isAMPM) {
                            if (token.count == 1) oss << fmt::format("{}", hour12);
                            else oss << fmt::format("{:02d}", hour12);
                        } else {
                            if (token.count == 1) oss << fmt::format("{}", t.tm_hour);
                            else oss << fmt::format("{:02d}", t.tm_hour);
                        }
                        break;
                    case XLFormatTokenType::Minute:
                        if (token.count == 1) oss << fmt::format("{}", t.tm_min);
                        else oss << fmt::format("{:02d}", t.tm_min);
                        break;
                    case XLFormatTokenType::Second:
                        if (token.count == 1) oss << fmt::format("{}", t.tm_sec);
                        else oss << fmt::format("{:02d}", t.tm_sec);
                        break;
                    case XLFormatTokenType::AMPM:
                        if (token.count == 5) { // AM/PM
                            oss << (t.tm_hour < 12 ? "AM" : "PM");
                        } else { // A/P
                            oss << (t.tm_hour < 12 ? "A" : "P");
                        }
                        break;
                    case XLFormatTokenType::TextPlaceholder:
                        oss << "@";
                        break;
                    default:
                        oss << token.value;
                        break;
                }
            }
            return oss.str();
        } catch (...) {
            return fmt::format("{}", excelDate); // Fallback
        }
    }

    std::string XLNumberFormatter::formatNumeric(double number, const FormatSection& section, bool addMinusSign) const {
        double evalNumber = number;
        if (section.hasPercent) {
            evalNumber *= 100.0;
        }

        bool useThousands = false;
        for (const auto& token : section.tokens) {
            if (token.type == XLFormatTokenType::Thousands) {
                useThousands = true;
                break;
            }
        }
        
        std::string result;
        if (useThousands) {
            std::string raw = fmt::format("{:.{}f}", evalNumber, section.decimalPlaces);
            
            size_t decPos = raw.find('.');
            if (decPos == std::string::npos) decPos = raw.length();
            
            int intStart = (raw.length() > 0 && raw[0] == '-') ? 1 : 0;
            
            for (int k = static_cast<int>(decPos) - 3; k > intStart; k -= 3) {
                raw.insert(k, ",");
            }
            result = raw;
        } else {
            result = fmt::format("{:.{}f}", evalNumber, section.decimalPlaces);
        }

        // If format is entirely literal, just output literals (unless it has numerics)
        bool hasNumericToken = false;
        for (const auto& t : section.tokens) {
            if (t.type == XLFormatTokenType::DigitZero || t.type == XLFormatTokenType::DigitOpt || t.type == XLFormatTokenType::Decimal || t.type == XLFormatTokenType::Thousands) {
                hasNumericToken = true;
                break;
            }
        }
        if (!hasNumericToken && section.tokens.size() > 0) {
            // It might just be text formatting with no # or 0. e.g. "-"
            std::ostringstream oss;
            for (const auto& token : section.tokens) {
                if (token.type == XLFormatTokenType::Literal) oss << token.value;
                else if (token.type == XLFormatTokenType::TextPlaceholder) oss << "@";
            }
            return oss.str();
        }
        
        // Simple approach: combine literals.
        std::ostringstream oss;
        if (addMinusSign) {
            oss << "-";
        }
        
        bool numAdded = false;
        for (const auto& token : section.tokens) {
            if (token.type == XLFormatTokenType::Literal) {
                oss << token.value;
            } else if (token.type == XLFormatTokenType::TextPlaceholder) {
                oss << "@";
            } else if (token.type == XLFormatTokenType::Percent) {
                oss << "%";
            } else if (token.type == XLFormatTokenType::DigitZero || 
                       token.type == XLFormatTokenType::DigitOpt ||
                       token.type == XLFormatTokenType::Decimal ||
                       token.type == XLFormatTokenType::Thousands) {
                if (!numAdded) {
                    oss << result;
                    numAdded = true;
                }
            } else {
                // Ignore unexpected or duplicate numeric tokens
                if (!numAdded) {
                    oss << result;
                    numAdded = true;
                }
            }
        }

        // If no numeric tokens were in the format string (e.g. only text), we still output literals
        return oss.str();
    }
    
    std::string XLNumberFormatter::formatText(const std::string& text, const FormatSection& section) const {
        std::ostringstream oss;
        for (const auto& token : section.tokens) {
            if (token.type == XLFormatTokenType::Literal) {
                oss << token.value;
            } else if (token.type == XLFormatTokenType::TextPlaceholder) {
                oss << text;
            } else {
                oss << token.value;
            }
        }
        return oss.str();
    }

} // namespace OpenXLSX
/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLRICHTEXT_HPP
#define OPENXLSX_XLRICHTEXT_HPP

// ===== External Includes ===== //
#include <optional>
#include <string>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLColor.hpp"
#include "XLStyles.hpp"

namespace OpenXLSX
{
    /**
     * @brief A class representing a single run of rich text.
     */
    class OPENXLSX_EXPORT XLRichTextRun
    {
    public:
        XLRichTextRun() = default;
        explicit XLRichTextRun(const std::string& text) : m_text(text) {}

        const std::string& text() const { return m_text; }
        XLRichTextRun&     setText(const std::string& text)
        {
            m_text = text;
            return *this;
        }

        // Font properties
        std::optional<std::string> fontName() const { return m_fontName; }
        XLRichTextRun&             setFontName(const std::string& name)
        {
            m_fontName = name;
            return *this;
        }

        std::optional<size_t> fontSize() const { return m_fontSize; }
        XLRichTextRun&        setFontSize(size_t size)
        {
            m_fontSize = size;
            return *this;
        }

        std::optional<XLColor> fontColor() const { return m_fontColor; }
        XLRichTextRun&         setFontColor(const XLColor& color)
        {
            m_fontColor = color;
            return *this;
        }

        std::optional<bool> bold() const { return m_bold; }
        XLRichTextRun&      setBold(bool bold = true)
        {
            m_bold = bold;
            return *this;
        }

        std::optional<bool> italic() const { return m_italic; }
        XLRichTextRun&      setItalic(bool italic = true)
        {
            m_italic = italic;
            return *this;
        }

        // Returns true if any underline style is set (for backward compatibility)
        std::optional<bool> underline() const { return m_underlineStyle.has_value() ? m_underlineStyle.value() != XLUnderlineNone : std::optional<bool>{}; }
        
        XLRichTextRun&      setUnderline(bool underline = true)
        {
            m_underlineStyle = underline ? XLUnderlineSingle : XLUnderlineNone;
            return *this;
        }

        std::optional<XLUnderlineStyle> underlineStyle() const { return m_underlineStyle; }
        XLRichTextRun&                  setUnderlineStyle(XLUnderlineStyle style)
        {
            m_underlineStyle = style;
            return *this;
        }

        std::optional<bool> strikethrough() const { return m_strikethrough; }
        XLRichTextRun&      setStrikethrough(bool strikethrough = true)
        {
            m_strikethrough = strikethrough;
            return *this;
        }

        std::optional<XLVerticalAlignRunStyle> vertAlign() const { return m_vertAlign; }
        XLRichTextRun&                         setVertAlign(XLVerticalAlignRunStyle align)
        {
            m_vertAlign = align;
            return *this;
        }

        // Convenience methods for Superscript and Subscript
        XLRichTextRun& setSuperscript(bool enable = true)
        {
            if (enable) m_vertAlign = XLSuperscript;
            else if (m_vertAlign == XLSuperscript) m_vertAlign.reset();
            return *this;
        }

        XLRichTextRun& setSubscript(bool enable = true)
        {
            if (enable) m_vertAlign = XLSubscript;
            else if (m_vertAlign == XLSubscript) m_vertAlign.reset();
            return *this;
        }

    private:
        std::string                            m_text;
        std::optional<std::string>             m_fontName;
        std::optional<size_t>                  m_fontSize;
        std::optional<XLColor>                 m_fontColor;
        std::optional<bool>                    m_bold;
        std::optional<bool>                    m_italic;
        std::optional<XLUnderlineStyle>        m_underlineStyle;
        std::optional<bool>                    m_strikethrough;
        std::optional<XLVerticalAlignRunStyle> m_vertAlign;
    };

    /**
     * @brief A class representing rich text in a cell.
     */
    class OPENXLSX_EXPORT XLRichText
    {
    public:
        XLRichText() = default;
        explicit XLRichText(const std::string& plainText) { m_runs.emplace_back(plainText); }

        XLRichText& addRun(const XLRichTextRun& run)
        {
            m_runs.push_back(run);
            return *this;
        }
        XLRichTextRun& addRun(const std::string& text)
        {
            m_runs.emplace_back(text);
            return m_runs.back();
        }
        const std::vector<XLRichTextRun>& runs() const { return m_runs; }
        std::vector<XLRichTextRun>&       runs() { return m_runs; }

        /**
         * @brief Get the plain text representation of the rich text.
         */
        std::string plainText() const
        {
            std::string result;
            for (const auto& run : m_runs) { result += run.text(); }
            return result;
        }

        bool empty() const { return m_runs.empty(); }
        void clear() { m_runs.clear(); }

        XLRichText& operator+=(const XLRichText& other)
        {
            for (const auto& run : other.runs()) {
                m_runs.push_back(run);
            }
            return *this;
        }

        XLRichText& operator+=(const std::string& text)
        {
            m_runs.emplace_back(text);
            return *this;
        }

        XLRichText& operator+=(const XLRichTextRun& run)
        {
            m_runs.push_back(run);
            return *this;
        }

    private:
        std::vector<XLRichTextRun> m_runs;
    };

    inline XLRichText operator+(XLRichText lhs, const XLRichText& rhs) {
        lhs += rhs;
        return lhs;
    }
    
    inline XLRichText operator+(XLRichText lhs, const std::string& rhs) {
        lhs += rhs;
        return lhs;
    }

    inline XLRichText operator+(XLRichText lhs, const XLRichTextRun& rhs) {
        lhs += rhs;
        return lhs;
    }

    // Equality operators
    inline bool operator==(const XLRichTextRun& lhs, const XLRichTextRun& rhs)
    {
        return lhs.text() == rhs.text() && lhs.fontName() == rhs.fontName() && lhs.fontSize() == rhs.fontSize() &&
               lhs.fontColor() == rhs.fontColor() && lhs.bold() == rhs.bold() && lhs.italic() == rhs.italic() &&
               lhs.underlineStyle() == rhs.underlineStyle() && lhs.strikethrough() == rhs.strikethrough() &&
               lhs.vertAlign() == rhs.vertAlign();
    }

    inline bool operator!=(const XLRichTextRun& lhs, const XLRichTextRun& rhs) { return !(lhs == rhs); }

    inline bool operator==(const XLRichText& lhs, const XLRichText& rhs) { return lhs.runs() == rhs.runs(); }

    inline bool operator!=(const XLRichText& lhs, const XLRichText& rhs) { return !(lhs == rhs); }

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLRICHTEXT_HPP

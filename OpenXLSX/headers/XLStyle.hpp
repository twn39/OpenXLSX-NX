#ifndef OPENXLSX_XLSTYLE_HPP
#define OPENXLSX_XLSTYLE_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLColor.hpp"
#include "XLConstants.hpp"
#include "XLStyles.hpp"
#include <cstdint>
#include <optional>
#include <string>

namespace OpenXLSX
{
    /**
     * @brief A high-level, human-ergonomic structure representing the styling of a cell or range.
     * This acts as a builder and facade over the complex underlying OpenXLSX XLStyles system.
     */
    struct XLStyle
    {
        struct Font
        {
            std::optional<std::string> name;
            std::optional<uint32_t>    size;
            std::optional<XLColor>     color;
            std::optional<bool>        bold;
            std::optional<bool>        italic;
            std::optional<bool>        underline;
            std::optional<bool>        strikethrough;
        } font;

        struct Fill
        {
            std::optional<XLPatternType> pattern;
            std::optional<XLColor>       fgColor;
            std::optional<XLColor>       bgColor;
        } fill;

        struct BorderElement
        {
            std::optional<XLLineStyle> style;
            std::optional<XLColor>     color;
        };

        struct Border
        {
            BorderElement left;
            BorderElement right;
            BorderElement top;
            BorderElement bottom;
            BorderElement diagonal;
            std::optional<bool> diagonalUp;
            std::optional<bool> diagonalDown;
        } border;

        struct Alignment
        {
            std::optional<XLAlignmentStyle> horizontal;
            std::optional<XLAlignmentStyle> vertical;
            std::optional<bool>             wrapText;
            std::optional<uint16_t>         textRotation;
            std::optional<uint32_t>         indent;
        } alignment;

        std::optional<std::string> numberFormat;
    };

    // ===== Equality Operators ===== //

    inline bool operator==(const XLStyle::Font& lhs, const XLStyle::Font& rhs) noexcept
    {
        return lhs.name == rhs.name && lhs.size == rhs.size && lhs.color == rhs.color &&
               lhs.bold == rhs.bold && lhs.italic == rhs.italic && lhs.underline == rhs.underline &&
               lhs.strikethrough == rhs.strikethrough;
    }

    inline bool operator!=(const XLStyle::Font& lhs, const XLStyle::Font& rhs) noexcept { return !(lhs == rhs); }

    inline bool operator==(const XLStyle::Fill& lhs, const XLStyle::Fill& rhs) noexcept
    {
        return lhs.pattern == rhs.pattern && lhs.fgColor == rhs.fgColor && lhs.bgColor == rhs.bgColor;
    }

    inline bool operator!=(const XLStyle::Fill& lhs, const XLStyle::Fill& rhs) noexcept { return !(lhs == rhs); }

    inline bool operator==(const XLStyle::BorderElement& lhs, const XLStyle::BorderElement& rhs) noexcept
    {
        return lhs.style == rhs.style && lhs.color == rhs.color;
    }

    inline bool operator!=(const XLStyle::BorderElement& lhs, const XLStyle::BorderElement& rhs) noexcept { return !(lhs == rhs); }

    inline bool operator==(const XLStyle::Border& lhs, const XLStyle::Border& rhs) noexcept
    {
        return lhs.left == rhs.left && lhs.right == rhs.right && lhs.top == rhs.top &&
               lhs.bottom == rhs.bottom && lhs.diagonal == rhs.diagonal &&
               lhs.diagonalUp == rhs.diagonalUp && lhs.diagonalDown == rhs.diagonalDown;
    }

    inline bool operator!=(const XLStyle::Border& lhs, const XLStyle::Border& rhs) noexcept { return !(lhs == rhs); }

    inline bool operator==(const XLStyle::Alignment& lhs, const XLStyle::Alignment& rhs) noexcept
    {
        return lhs.horizontal == rhs.horizontal && lhs.vertical == rhs.vertical &&
               lhs.wrapText == rhs.wrapText && lhs.textRotation == rhs.textRotation &&
               lhs.indent == rhs.indent;
    }

    inline bool operator!=(const XLStyle::Alignment& lhs, const XLStyle::Alignment& rhs) noexcept { return !(lhs == rhs); }

    inline bool operator==(const XLStyle& lhs, const XLStyle& rhs) noexcept
    {
        return lhs.font == rhs.font && lhs.fill == rhs.fill && lhs.border == rhs.border &&
               lhs.alignment == rhs.alignment && lhs.numberFormat == rhs.numberFormat;
    }

    inline bool operator!=(const XLStyle& lhs, const XLStyle& rhs) noexcept { return !(lhs == rhs); }

    // ===== Hash Functor ===== //

    struct XLStyleHash
    {
        static void hashCombine(size_t& seed, size_t val) noexcept
        {
            seed ^= val + 0x9e3779b97f4a7c15ULL + (seed << 6) + (seed >> 2);
        }

        static size_t hashColor(const XLColor& c) noexcept
        {
            uint32_t val = (static_cast<uint32_t>(c.alpha()) << 24) |
                           (static_cast<uint32_t>(c.red()) << 16) |
                           (static_cast<uint32_t>(c.green()) << 8) |
                           static_cast<uint32_t>(c.blue());
            return std::hash<uint32_t>{}(val);
        }

        size_t operator()(const XLStyle& s) const noexcept
        {
            size_t seed = 0;

            // Font
            if (s.font.name) { hashCombine(seed, 1); hashCombine(seed, std::hash<std::string>{}(*s.font.name)); }
            if (s.font.size) { hashCombine(seed, 2); hashCombine(seed, std::hash<uint32_t>{}(*s.font.size)); }
            if (s.font.color) { hashCombine(seed, 3); hashCombine(seed, hashColor(*s.font.color)); }
            if (s.font.bold) { hashCombine(seed, *s.font.bold ? 4 : 5); }
            if (s.font.italic) { hashCombine(seed, *s.font.italic ? 6 : 7); }
            if (s.font.underline) { hashCombine(seed, *s.font.underline ? 8 : 9); }
            if (s.font.strikethrough) { hashCombine(seed, *s.font.strikethrough ? 10 : 11); }

            // Fill
            if (s.fill.pattern) { hashCombine(seed, 12); hashCombine(seed, static_cast<size_t>(*s.fill.pattern)); }
            if (s.fill.fgColor) { hashCombine(seed, 13); hashCombine(seed, hashColor(*s.fill.fgColor)); }
            if (s.fill.bgColor) { hashCombine(seed, 14); hashCombine(seed, hashColor(*s.fill.bgColor)); }

            // Border
            auto hashBorderElem = [&](size_t offset, const XLStyle::BorderElement& b) {
                if (b.style) { hashCombine(seed, offset); hashCombine(seed, static_cast<size_t>(*b.style)); }
                if (b.color) { hashCombine(seed, offset + 1); hashCombine(seed, hashColor(*b.color)); }
            };
            hashBorderElem(20, s.border.left);
            hashBorderElem(30, s.border.right);
            hashBorderElem(40, s.border.top);
            hashBorderElem(50, s.border.bottom);
            hashBorderElem(60, s.border.diagonal);
            if (s.border.diagonalUp) hashCombine(seed, *s.border.diagonalUp ? 70 : 71);
            if (s.border.diagonalDown) hashCombine(seed, *s.border.diagonalDown ? 72 : 73);

            // Alignment
            if (s.alignment.horizontal) { hashCombine(seed, 80); hashCombine(seed, static_cast<size_t>(*s.alignment.horizontal)); }
            if (s.alignment.vertical) { hashCombine(seed, 81); hashCombine(seed, static_cast<size_t>(*s.alignment.vertical)); }
            if (s.alignment.wrapText) hashCombine(seed, *s.alignment.wrapText ? 82 : 83);
            if (s.alignment.textRotation) { hashCombine(seed, 84); hashCombine(seed, std::hash<uint16_t>{}(*s.alignment.textRotation)); }
            if (s.alignment.indent) { hashCombine(seed, 85); hashCombine(seed, std::hash<uint32_t>{}(*s.alignment.indent)); }

            // NumberFormat
            if (s.numberFormat) { hashCombine(seed, 90); hashCombine(seed, std::hash<std::string>{}(*s.numberFormat)); }

            return seed;
        }
    };

}    // namespace OpenXLSX

namespace std
{
    template<>
    struct hash<OpenXLSX::XLStyle>
    {
        size_t operator()(const OpenXLSX::XLStyle& s) const noexcept
        {
            return OpenXLSX::XLStyleHash{}(s);
        }
    };
}

#endif    // OPENXLSX_XLSTYLE_HPP

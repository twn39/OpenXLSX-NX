#ifndef OPENXLSX_XLSTYLE_HPP
#define OPENXLSX_XLSTYLE_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLColor.hpp"
#include "XLConstants.hpp"
#include <cstdint>
#include <optional>
#include <string>

namespace OpenXLSX
{
    enum XLLineStyle : uint8_t;
    enum XLPatternType : uint8_t;
    enum XLAlignmentStyle : uint8_t;

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
        } border;

        struct Alignment
        {
            std::optional<XLAlignmentStyle> horizontal;
            std::optional<XLAlignmentStyle> vertical;
            std::optional<bool>             wrapText;
        } alignment;

        std::optional<std::string> numberFormat;
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLSTYLE_HPP

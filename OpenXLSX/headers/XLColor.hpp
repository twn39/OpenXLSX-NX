#ifndef OPENXLSX_XLCOLOR_HPP
#define OPENXLSX_XLCOLOR_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <cstdint>    // Pull request #276
#include <string>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"

namespace OpenXLSX
{
    /**
         */
    class OPENXLSX_EXPORT XLColor
    {
friend bool operator==(const XLColor& lhs, const XLColor& rhs);
        friend bool operator!=(const XLColor& lhs, const XLColor& rhs);

    public:
        /**
         * @brief Default constructor. Initializes to opaque black (A=255, R=0, G=0, B=0).
         */
        XLColor();

        /**
         * @brief Constructs a color with explicit ARGB channels.
         */
        XLColor(uint8_t alpha, uint8_t red, uint8_t green, uint8_t blue);

        /**
         * @brief Constructs an opaque color with RGB channels.
         */
        XLColor(uint8_t red, uint8_t green, uint8_t blue);

        /**
         * @brief Constructs a color from an ARGB or RGB hex string (e.g., "FF0000" or "FFFF0000").
         */
        explicit XLColor(std::string_view hexCode);

        XLColor(const XLColor& other);

        XLColor(XLColor&& other) noexcept;

        ~XLColor();

        XLColor& operator=(const XLColor& other);

        XLColor& operator=(XLColor&& other) noexcept;

        /**
         * @brief Sets explicit ARGB channels.
         */
        void set(uint8_t alpha, uint8_t red, uint8_t green, uint8_t blue);

        /**
         * @brief Sets RGB channels. Alpha is implicitly set to 255 (opaque).
         */
        void set(uint8_t red = 0, uint8_t green = 0, uint8_t blue = 0);

        /**
         * @brief Sets the color from an ARGB or RGB hex string.
         */
        void set(std::string_view hexCode);

        /**
         * @brief Retrieves the alpha (opacity) channel value.
         */
        [[nodiscard]] uint8_t alpha() const;

        /**
         * @brief Retrieves the red channel value.
         */
        [[nodiscard]] uint8_t red() const;

        /**
         * @brief Retrieves the green channel value.
         */
        [[nodiscard]] uint8_t green() const;

        /**
         * @brief Retrieves the blue channel value.
         */
        [[nodiscard]] uint8_t blue() const;

        /**
         * @brief Returns the 8-character ARGB hex string representation used in OOXML (e.g., "FFFF0000").
         */
        [[nodiscard]] std::string hex() const;
private:
        uint8_t m_alpha{255};

        uint8_t m_red{0};

        uint8_t m_green{0};

        uint8_t m_blue{0};
    };

}    // namespace OpenXLSX

namespace OpenXLSX
{
    /**
     * @brief Checks if two colors have identical ARGB channels.
     */
    inline bool operator==(const XLColor& lhs, const XLColor& rhs)
    { return lhs.alpha() == rhs.alpha() and lhs.red() == rhs.red() and lhs.green() == rhs.green() and lhs.blue() == rhs.blue(); }

    /**
     * @brief Checks if two colors differ in any of their ARGB channels.
     */
    inline bool operator!=(const XLColor& lhs, const XLColor& rhs) { return !(lhs == rhs); }

}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLCOLOR_HPP

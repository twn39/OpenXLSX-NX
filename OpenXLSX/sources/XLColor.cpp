// ===== External Includes ===== //
#include <cstdint>    // pull requests #216, #232
#include <ios>        // std::hex
#include <sstream>

// ===== OpenXLSX Includes ===== //
#include "XLColor.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLColor::XLColor() = default;

/**
 * @details
 */
XLColor::XLColor(uint8_t alpha, uint8_t red, uint8_t green, uint8_t blue) : m_alpha(alpha), m_red(red), m_green(green), m_blue(blue) {}

/**
 * @details
 */
XLColor::XLColor(uint8_t red, uint8_t green, uint8_t blue) : m_red(red), m_green(green), m_blue(blue) {}

/**
 * @details
 */
XLColor::XLColor(std::string_view hexCode) { set(hexCode); }

/**
 * @details
 */
XLColor::XLColor(const XLColor& other) = default;

/**
 * @details
 */
XLColor::XLColor(XLColor&& other) noexcept = default;

/**
 * @details
 */
XLColor::~XLColor() = default;

/**
 * @details
 */
XLColor& XLColor::operator=(const XLColor& other) = default;

/**
 * @details
 */
XLColor& XLColor::operator=(XLColor&& other) noexcept = default;

/**
 * @details
 */
void XLColor::set(uint8_t alpha, uint8_t red, uint8_t green, uint8_t blue)
{
    m_alpha = alpha;
    m_red   = red;
    m_green = green;
    m_blue  = blue;
}

/**
 * @details
 */
void XLColor::set(uint8_t red, uint8_t green, uint8_t blue)
{
    m_red   = red;
    m_green = green;
    m_blue  = blue;
}

/**
 * @details
 */
#include <charconv>

void XLColor::set(std::string_view hexCode)
{
    constexpr size_t hexCodeSizeWithoutAlpha = 6;
    constexpr size_t hexCodeSizeWithAlpha    = 8;

    if (hexCode.size() != hexCodeSizeWithoutAlpha && hexCode.size() != hexCodeSizeWithAlpha) {
        throw XLInputError("Invalid color code");
    }
    
    // Helper to parse 2 hex chars into uint8_t
    auto parseHex = [](std::string_view str) -> uint8_t {
        uint8_t val = 0;
        auto [ptr, ec] = std::from_chars(str.data(), str.data() + 2, val, 16);
        if (ec != std::errc()) throw XLInputError("Invalid hex string");
        return val;
    };

    if (hexCode.size() == hexCodeSizeWithoutAlpha) {
        m_alpha = 255; // Default alpha is opaque
        m_red   = parseHex(hexCode.substr(0, 2));
        m_green = parseHex(hexCode.substr(2, 2));
        m_blue  = parseHex(hexCode.substr(4, 2));
    } else {
        m_alpha = parseHex(hexCode.substr(0, 2));
        m_red   = parseHex(hexCode.substr(2, 2));
        m_green = parseHex(hexCode.substr(4, 2));
        m_blue  = parseHex(hexCode.substr(6, 2));
    }
}

/**
 * @details
 */
uint8_t XLColor::alpha() const { return m_alpha; }

/**
 * @details
 */
uint8_t XLColor::red() const { return m_red; }

/**
 * @details
 */
uint8_t XLColor::green() const { return m_green; }

/**
 * @details
 */
uint8_t XLColor::blue() const { return m_blue; }

/**
 * @details
 */
#include <fmt/format.h>

std::string XLColor::hex() const
{
    return fmt::format("{:02x}{:02x}{:02x}{:02x}", m_alpha, m_red, m_green, m_blue);
}

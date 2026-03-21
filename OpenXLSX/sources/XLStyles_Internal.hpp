#ifndef OPENXLSX_XLSTYLES_INTERNAL_HPP
#define OPENXLSX_XLSTYLES_INTERNAL_HPP

#include <pugixml.hpp>
#include <string_view>
#include <fmt/format.h>
#include "XLException.hpp"

namespace OpenXLSX {

    inline void copyXMLNode(XMLNode& destination, const XMLNode& source)
    {
        if (not source.empty()) {
            // ===== Copy all XML child nodes
            for (XMLNode child = source.first_child(); not child.empty(); child = child.next_sibling()) destination.append_copy(child);
            // ===== Copy all XML attributes
            for (XMLAttribute attr = source.first_attribute(); not attr.empty(); attr = attr.next_attribute())
                destination.append_copy(attr);
        }
    }

    template<typename E>
    struct EnumStringMap {
        std::string_view str;
        E val;
    };

    template<typename E, size_t N>
    constexpr E EnumFromString(std::string_view str, const EnumStringMap<E> (&mapping)[N], E invalidVal) {
        for (const auto& kv : mapping) {
            if (kv.str == str) return kv.val;
        }
        return invalidVal;
    }

    template<typename E, size_t N>
    constexpr const char* EnumToString(E val, const EnumStringMap<E> (&mapping)[N], const char* invalidStr) {
        for (const auto& kv : mapping) {
            if (kv.val == val) return kv.str.data();
        }
        return invalidStr;
    }

    /**
     * @brief Format val as a string with decimalPlaces
     */
    inline std::string formatDoubleAsString(double val, int decimalPlaces = 2)
    {
        return fmt::format("{:.{}f}", val, decimalPlaces);
    }

    /**
     * @brief Check that a double value is within range, and format it as a string with decimalPlaces
     */
    inline std::string checkAndFormatDoubleAsString(double val, double min, double max, double absTolerance, int decimalPlaces = 2)
    {
        if (max <= min or absTolerance < 0.0)
            throw XLException("checkAndFormatDoubleAsString: max must be greater than min and absTolerance must be >= 0.0");
        if (val < min - absTolerance or val > max + absTolerance) return "";    // range check
        if (val < min)
            val = min;
        else if (val > max)
            val = max;    // fix rounding errors within tolerance
        return formatDoubleAsString(val, decimalPlaces);
    }

} // namespace OpenXLSX

#endif // OPENXLSX_XLSTYLES_INTERNAL_HPP
/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLIMAGEPARSER_HPP
#define OPENXLSX_XLIMAGEPARSER_HPP

#include "OpenXLSX-Exports.hpp"
#include <cstdint>
#include <string>

#include <gsl/span>

namespace OpenXLSX
{

    struct XLImageSize
    {
        uint32_t    width{0};
        uint32_t    height{0};
        std::string extension{""};
        bool        valid{false};
    };

    /**
     * @brief Parses binary image headers (PNG, JPG, GIF) to extract width and height without external dependencies.
     */
    class OPENXLSX_EXPORT XLImageParser
    {
    public:
        static XLImageSize parseDimensions(const std::string& path);
        static XLImageSize parseDimensions(gsl::span<const uint8_t> data);
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLIMAGEPARSER_HPP

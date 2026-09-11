/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLEMUCONVERTER_HPP
#define OPENXLSX_XLEMUCONVERTER_HPP

#include <cstdint>

namespace OpenXLSX
{
    /**
     * @brief Utility for handling Excel's complex coordinate and measurement conversions.
     * Excel uses EMU (English Metric Units) for drawings. 1 inch = 914400 EMUs.
     * At 96 DPI (Standard), 1 pixel = 9525 EMUs.
     */
    class XLEMUConverter
    {
    public:
        static constexpr int32_t PIXELS_TO_EMU_FACTOR = 9525;

        static constexpr int32_t pixelsToEmu(int32_t pixels) noexcept { return pixels * PIXELS_TO_EMU_FACTOR; }

        static constexpr int32_t emuToPixels(int32_t emus) noexcept { return emus / PIXELS_TO_EMU_FACTOR; }
    };
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLEMUCONVERTER_HPP

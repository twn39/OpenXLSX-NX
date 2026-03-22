#ifndef OPENXLSX_XLEMUCONVERTER_HPP
#define OPENXLSX_XLEMUCONVERTER_HPP

#include <cstdint>

namespace OpenXLSX {
    /**
     * @brief Utility for handling Excel's complex coordinate and measurement conversions.
     * Excel uses EMU (English Metric Units) for drawings. 1 inch = 914400 EMUs.
     * At 96 DPI (Standard), 1 pixel = 9525 EMUs.
     */
    class XLEMUConverter {
    public:
        static constexpr int32_t PIXELS_TO_EMU_FACTOR = 9525;
        
        static constexpr int32_t pixelsToEmu(int32_t pixels) noexcept {
            return pixels * PIXELS_TO_EMU_FACTOR;
        }

        static constexpr int32_t emuToPixels(int32_t emus) noexcept {
            return emus / PIXELS_TO_EMU_FACTOR;
        }
    };
}

#endif // OPENXLSX_XLEMUCONVERTER_HPP

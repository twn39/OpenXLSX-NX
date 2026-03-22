#ifndef OPENXLSX_XLCONSTANTS_HPP
#define OPENXLSX_XLCONSTANTS_HPP

namespace OpenXLSX
{
    inline constexpr uint16_t MAX_COLS = 16'384;
    inline constexpr uint32_t MAX_ROWS = 1'048'576;

    class XLDistance {
    public:
        XLDistance() = default;
        static XLDistance Pixels(uint32_t px) { return XLDistance(static_cast<uint64_t>(px) * 9525); }
        static XLDistance Centimeters(double cm) { return XLDistance(static_cast<uint64_t>(cm * 360000.0)); }
        static XLDistance Points(double pt) { return XLDistance(static_cast<uint64_t>(pt * 12700.0)); }
        static XLDistance EMU(uint64_t emu) { return XLDistance(emu); }

        [[nodiscard]] uint64_t getEMU() const { return m_emu; }
        [[nodiscard]] uint32_t getPixels() const { return static_cast<uint32_t>(m_emu / 9525); }

    private:
        explicit XLDistance(uint64_t emu) : m_emu(emu) {}
        uint64_t m_emu {0};
    };

    inline namespace DistanceLiterals {
        inline XLDistance operator""_px(unsigned long long px) { return XLDistance::Pixels(static_cast<uint32_t>(px)); }
        inline XLDistance operator""_cm(long double cm) { return XLDistance::Centimeters(static_cast<double>(cm)); }
        inline XLDistance operator""_cm(unsigned long long cm) { return XLDistance::Centimeters(static_cast<double>(cm)); }

        inline XLDistance operator""_pt(long double pt) { return XLDistance::Points(static_cast<double>(pt)); }
        inline XLDistance operator""_pt(unsigned long long pt) { return XLDistance::Points(static_cast<double>(pt)); }

    }


    // anchoring a comment shape below these values was not possible in LibreOffice - TBC with MS Office
    inline constexpr uint16_t MAX_SHAPE_ANCHOR_COLUMN = 13067;    // column "SHO"
    inline constexpr uint32_t MAX_SHAPE_ANCHOR_ROW    = 852177;
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLCONSTANTS_HPP

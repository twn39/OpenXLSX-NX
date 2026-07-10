#ifndef OPENXLSX_XLSTYLES_COMMON_HPP
#define OPENXLSX_XLSTYLES_COMMON_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

// ===== External Includes ===== //
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"

namespace OpenXLSX
{
    using namespace std::literals::string_view_literals;    // enables sv suffix only

    // Forward declaration to avoid circular includes (XLStyle.hpp uses enums from this header)
    struct XLStyle;

    using XLStyleIndex = size_t;    // custom data type for XLStyleIndex

    constexpr const uint32_t XLInvalidUInt16  = 0xffff;             // used to signal "value not defined" for uint16_t return types
    constexpr const uint32_t XLInvalidUInt32  = 0xffffffff;         // used to signal "value not defined" for uint32_t return types
    constexpr const uint32_t XLDeleteProperty = XLInvalidUInt32;    // when 0 or "" is not the same as "property does not exist", this value
    //                                                            //  can be passed to setter functions to delete the property from XML
    //                                                            //  currently supported in: XLDataBarColor::setTheme

    constexpr const bool XLPermitXfID = true;    // use with XLCellFormat constructor to enable xfId() getter and setXfId() setter

    constexpr const bool XLCreateIfMissing = true;     // use with XLCellFormat::alignment(XLCreateIfMissing)
    constexpr const bool XLDoNotCreate     = false;    // use with XLCellFormat::alignment(XLDoNotCreate)

    constexpr const bool XLForceFillType = true;

    constexpr const char* XLDefaultStylesPrefix       = "\n\t";      // indentation to use for newly created root level style node tags
    constexpr const char* XLDefaultStyleEntriesPrefix = "\n\t\t";    // indentation to use for newly created style entry nodes

    constexpr const XLStyleIndex XLDefaultCellFormat = 0;    // default cell format index in xl/styles.xml:<styleSheet>:<cellXfs>

    // ===== As pugixml attributes are not guaranteed to support value range of XLStyleIndex, use 32 bit unsigned int
    constexpr const XLStyleIndex XLInvalidStyleIndex = XLInvalidUInt32;    // as a function return value, indicates no valid index

    constexpr const uint32_t XLDefaultFontSize       = 12;            //
    constexpr const char*    XLDefaultFontColor      = "ff000000";    // default font color
    constexpr const char*    XLDefaultFontColorTheme = "";            // Theme color ID from theme1.xml (e.g., "1" for dark1)
    constexpr const char*    XLDefaultFontName       = "Arial";       //
    constexpr const uint32_t XLDefaultFontFamily     = 0;             // Font Family (0 = N/A, 1 = Roman, 2 = Swiss, 3 = Modern)
    constexpr const uint32_t XLDefaultFontCharset    = 1;             // Character set (0 = ANSI, 1 = Default, 2 = Symbol)

    constexpr const char* XLDefaultLineStyle = "";    // empty string = line not set

    // Forward declarations of styles subdomain classes (defined in XLStyles_*.hpp)
    class XLNumberFormat;
    class XLNumberFormats;
    class XLFont;
    class XLFonts;
    class XLDataBarColor;
    class XLFill;
    class XLFills;
    class XLLine;
    class XLBorder;
    class XLBorders;
    class XLAlignment;
    class XLCellFormat;
    class XLCellFormats;
    class XLCellStyle;
    class XLCellStyles;
    class XLDxf;
    class XLDxfs;
    class XLStyles;

    enum XLUnderlineStyle : uint8_t { 
        XLUnderlineNone = 0, 
        XLUnderlineSingle = 1, 
        XLUnderlineDouble = 2, 
        XLUnderlineSingleAccounting = 3, 
        XLUnderlineDoubleAccounting = 4, 
        XLUnderlineInvalid = 255 
    };

    enum XLFontSchemeStyle : uint8_t {
        XLFontSchemeNone    = 0,     // <scheme val="none"/>
        XLFontSchemeMajor   = 1,     // <scheme val="major"/>
        XLFontSchemeMinor   = 2,     // <scheme val="minor"/>
        XLFontSchemeInvalid = 255    // all other values
    };

    enum XLVerticalAlignRunStyle : uint8_t {
        XLBaseline                = 0,    // <vertAlign val="baseline"/>
        XLSubscript               = 1,    // <vertAlign val="subscript"/>
        XLSuperscript             = 2,    // <vertAlign val="superscript"/>
        XLVerticalAlignRunInvalid = 255
    };

    enum XLFillType : uint8_t {
        XLGradientFill    = 0,      // <gradientFill />
        XLPatternFill     = 1,      // <patternFill />
        XLFillTypeInvalid = 255,    // any child of <fill> that is not one of the above
    };

    enum XLGradientType : uint8_t { XLGradientLinear = 0, XLGradientPath = 1, XLGradientTypeInvalid = 255 };

    enum XLPatternType : uint8_t {
        XLPatternNone            = 0,     // "none"
        XLPatternSolid           = 1,     // "solid"
        XLPatternMediumGray      = 2,     // "mediumGray"
        XLPatternDarkGray        = 3,     // "darkGray"
        XLPatternLightGray       = 4,     // "lightGray"
        XLPatternDarkHorizontal  = 5,     // "darkHorizontal"
        XLPatternDarkVertical    = 6,     // "darkVertical"
        XLPatternDarkDown        = 7,     // "darkDown"
        XLPatternDarkUp          = 8,     // "darkUp"
        XLPatternDarkGrid        = 9,     // "darkGrid"
        XLPatternDarkTrellis     = 10,    // "darkTrellis"
        XLPatternLightHorizontal = 11,    // "lightHorizontal"
        XLPatternLightVertical   = 12,    // "lightVertical"
        XLPatternLightDown       = 13,    // "lightDown"
        XLPatternLightUp         = 14,    // "lightUp"
        XLPatternLightGrid       = 15,    // "lightGrid"
        XLPatternLightTrellis    = 16,    // "lightTrellis"
        XLPatternGray125         = 17,    // "gray125"
        XLPatternGray0625        = 18,    // "gray0625"
        XLPatternTypeInvalid     = 255    // any patternType that is not one of the above
    };
    constexpr const XLFillType    XLDefaultFillType       = XLPatternFill;    // node name for the pattern description is derived from this
    constexpr const XLPatternType XLDefaultPatternType    = XLPatternNone;    // attribute patternType default value: no fill
    constexpr const char*         XLDefaultPatternFgColor = "ffffffff";       // child node fgcolor attribute rgb value
    constexpr const char*         XLDefaultPatternBgColor = "ff000000";       // child node bgcolor attribute rgb value

    enum XLLineType : uint8_t {
        XLLineLeft       = 0,
        XLLineRight      = 1,
        XLLineTop        = 2,
        XLLineBottom     = 3,
        XLLineDiagonal   = 4,
        XLLineVertical   = 5,
        XLLineHorizontal = 6,
        XLLineInvalid    = 255
    };

    enum XLLineStyle : uint8_t {
        XLLineStyleNone             = 0,
        XLLineStyleThin             = 1,
        XLLineStyleMedium           = 2,
        XLLineStyleDashed           = 3,
        XLLineStyleDotted           = 4,
        XLLineStyleThick            = 5,
        XLLineStyleDouble           = 6,
        XLLineStyleHair             = 7,
        XLLineStyleMediumDashed     = 8,
        XLLineStyleDashDot          = 9,
        XLLineStyleMediumDashDot    = 10,
        XLLineStyleDashDotDot       = 11,
        XLLineStyleMediumDashDotDot = 12,
        XLLineStyleSlantDashDot     = 13,
        XLLineStyleInvalid          = 255
    };

    enum XLAlignmentStyle : uint8_t {
        XLAlignGeneral          = 0,     // value="general",          horizontal only
        XLAlignLeft             = 1,     // value="left",             horizontal only
        XLAlignRight            = 2,     // value="right",            horizontal only
        XLAlignCenter           = 3,     // value="center",           both
        XLAlignTop              = 4,     // value="top",              vertical only
        XLAlignBottom           = 5,     // value="bottom",           vertical only
        XLAlignFill             = 6,     // value="fill",             horizontal only
        XLAlignJustify          = 7,     // value="justify",          both
        XLAlignCenterContinuous = 8,     // value="centerContinuous", horizontal only
        XLAlignDistributed      = 9,     // value="distributed",      both
        XLAlignInvalid          = 255    // all other values
    };

    enum XLReadingOrder : uint32_t { XLReadingOrderContextual = 0, XLReadingOrderLeftToRight = 1, XLReadingOrderRightToLeft = 2 };
}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLSTYLES_COMMON_HPP

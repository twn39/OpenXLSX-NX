#ifndef OPENXLSX_XLSTYLES_HPP
#define OPENXLSX_XLSTYLES_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

// Umbrella header: styles subsystem split into subdomain headers.
// Existing code may keep #include "XLStyles.hpp".
// Dependency order: Common -> pools (NumberFormats/Fonts/Fills) -> Borders (needs XLDataBarColor)
//                   -> CellFormats/CellStyles -> Dxfs (aggregates several pools).

#include "XLStyles_Common.hpp"
#include "XLStyles_NumberFormats.hpp"
#include "XLStyles_Fonts.hpp"
#include "XLStyles_Fills.hpp"
#include "XLStyles_Borders.hpp"
#include "XLStyles_CellFormats.hpp"
#include "XLStyles_CellStyles.hpp"
#include "XLStyles_Dxfs.hpp"

#include "XLXmlFile.hpp"

#include <gsl/gsl>
#include <memory>
#include <optional>
#include <string>
#include <string_view>

namespace OpenXLSX
{
    // XLStyles Class

    /**
     * @brief An encapsulation of the styles file (xl/styles.xml) in an Excel document package.
     */
    class OPENXLSX_EXPORT XLStyles : public XLXmlFile
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLStyles();

        /**
         * @brief
         * @param xmlData
         * @param suppressWarnings if true (SUPRESS_WARNINGS), messages such as "XLStyles: Ignoring currently unsupported <dxfs> node" will
         * be silenced
         * @param stylesPrefix Prefix any newly created root style nodes with this text as pugi::node_pcdata
         */
        explicit XLStyles(gsl::not_null<XLXmlData*> xmlData,
                          bool                      suppressWarnings = false,
                          std::string_view          stylesPrefix     = XLDefaultStylesPrefix);

        /**
         * @brief Destructor
         */
        ~XLStyles();

        /**
         * @brief The move constructor.
         * @param other an existing styles object other will be assigned to this
         */
        XLStyles(XLStyles&& other) noexcept;

        /**
         * @brief The copy constructor.
         * @param other an existing styles object other will be also referred by this
         */
        XLStyles(const XLStyles& other);

        /**
         * @brief move assignment
         * @param other
         * @return
         */
        XLStyles& operator=(XLStyles&& other) noexcept;

        /**
         * @brief copy assignment
         * @param other
         * @return
         */
        XLStyles& operator=(const XLStyles& other);

        /**
         * @brief Create a new custom number format with a unique ID and format code
         * @param formatCode The explicit format code string
         * @return The unique number format ID (numFmtId)
         */
        uint32_t createNumberFormat(std::string_view formatCode);

        /**
         * @brief Get the number formats object
         * @return An XLNumberFormats object
         */
        XLNumberFormats& numberFormats() const;

        /**
         * @brief Get the fonts object
         * @return An XLFonts object
         */
        XLFonts& fonts() const;

        /**
         * @brief Get the fills object
         * @return An XLFills object
         */
        XLFills& fills() const;

        /**
         * @brief Get the borders object
         * @return An XLBorders object
         */
        XLBorders& borders() const;

        /**
         * @brief Get the cell style formats object
         * @return An XLCellFormats object
         */
        XLCellFormats& cellStyleFormats() const;

        /**
         * @brief Get the cell formats object
         * @return An XLCellFormats object
         */
        XLCellFormats& cellFormats() const;

        /**
         * @brief Get the cell styles object
         * @return An XLCellStyles object
         */
        XLCellStyles& cellStyles() const;

        /**
         * @brief Get the differential cell formats object (dxfs)
         * @return An XLDxfs object
         */
        XLDxfs& dxfs() const;

        /**
         * @brief Backward compatibility alias for dxfs()
         */
        XLDxfs& diffCellFormats() const { return dxfs(); }

        /**
         * @brief Add a differential cell format (DXF) to the styles and return its index.
         * @param dxf The XLDxf object to add.
         * @return The index of the added DXF.
         */
        XLStyleIndex addDxf(const XLDxf& dxf);

        /**
         * @brief Add a named style with the given properties and return the index to be used in cells.
         * @param name Name of the new style
         * @param fontId Optional font index
         * @param fillId Optional fill index
         * @param borderId Optional border index
         * @param numFmtId Optional number format index
         * @return The XLStyleIndex in cellXfs that corresponds to this named style.
         */
        XLStyleIndex addNamedStyle(std::string_view            name,
                                   std::optional<XLStyleIndex> fontId   = std::nullopt,
                                   std::optional<XLStyleIndex> fillId   = std::nullopt,
                                   std::optional<XLStyleIndex> borderId = std::nullopt,
                                   std::optional<XLStyleIndex> numFmtId = std::nullopt);

        /**
         * @brief Look up a named style by name and return its index for cell formatting.
         * @param name Name of the style
         * @return The XLStyleIndex in cellXfs, or XLInvalidStyleIndex if not found.
         */
        XLStyleIndex namedStyle(std::string_view name) const;

        /**
         * @brief Apply an XLStyle descriptor, deduplicating every sub-pool entry
         *        (font, fill, border, numFmt, cellXf) and returning the cellXfs index.
         * @details This is the primary high-level entry point for bulk formatting:
         *          call it once per distinct visual style; pass the returned index to
         *          every cell that should share that style.  Repeated calls with an
         *          identical descriptor return the same index in O(1) time.
         * @param style A fully populated XLStyle descriptor.
         * @return The XLStyleIndex (in cellXfs) suitable for the cell's 's' attribute.
         */
        XLStyleIndex findOrCreateStyle(const XLStyle& style);

        // ---------- Protected Member Functions ---------- //
    private:
        bool                             m_suppressWarnings;    // if true, will suppress output of warnings where supported
        std::unique_ptr<XLNumberFormats> m_numberFormats;       // handle to the underlying number formats
        std::unique_ptr<XLFonts>         m_fonts;               // handle to the underlying fonts
        std::unique_ptr<XLFills>         m_fills;               // handle to the underlying fills
        std::unique_ptr<XLBorders>       m_borders;             // handle to the underlying border descriptions
        std::unique_ptr<XLCellFormats>   m_cellStyleFormats;    // handle to the underlying cell style formats descriptions
        std::unique_ptr<XLCellFormats>   m_cellFormats;         // handle to the underlying cell formats descriptions
        std::unique_ptr<XLCellStyles>    m_cellStyles;          // handle to the underlying cell styles
        std::unique_ptr<XLDxfs>          m_dxfs;                // handle to the underlying differential cell formats
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLSTYLES_HPP

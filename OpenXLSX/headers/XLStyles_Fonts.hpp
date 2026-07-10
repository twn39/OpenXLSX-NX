#ifndef OPENXLSX_XLSTYLES_FONTS_HPP
#define OPENXLSX_XLSTYLES_FONTS_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

#include "XLStyles_Common.hpp"
#include "XLColor.hpp"
#include "XLXmlParser.hpp"

#include <memory>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>


namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLFont
    {
        friend class XLFonts;    // for access to m_fontNode in XLFonts::create
        friend class XLStyles;
    public:                      // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLFont();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the fonts XMLNode. If no input is provided, a null node is used.
         */
        explicit XLFont(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLFont(const XLFont& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLFont(XLFont&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLFont();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLFont& operator=(const XLFont& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLFont& operator=(XLFont&& other) noexcept = default;

        /**
         * @brief Get the font name
         * @return The font name
         */
        std::string fontName() const;

        /**
         * @brief Get the font charset
         * @return The font charset id
         */
        size_t fontCharset() const;

        /**
         * @brief Get the font family
         * @return The font family id
         */
        size_t fontFamily() const;

        /**
         * @brief Get the font size
         * @return The font size
         */
        size_t fontSize() const;

        /**
         * @brief Get the font color
         * @return The font color
         */
        XLColor fontColor() const;

        /**
         * @brief Get the font bold status
         * @return true = bold, false = not bold
         */
        bool bold() const;

        /**
         * @brief Get the font italic (cursive) status
         * @return true = italic, false = not italice
         */
        bool italic() const;

        /**
         * @brief Get the font strikethrough status
         * @return true = strikethrough, false = no strikethrough
         */
        bool strikethrough() const;

        /**
         * @brief Get the font underline status
         * @return An XLUnderlineStyle value
         */
        XLUnderlineStyle underline() const;

        /**
         * @brief get the font scheme: none, major or minor
         * @return An XLFontSchemeStyle
         */
        XLFontSchemeStyle scheme() const;

        /**
         * @brief get the font vertical alignment run style: baseline, subscript or superscript
         * @return An XLVerticalAlignRunStyle
         */
        XLVerticalAlignRunStyle vertAlign() const;

        /**
         * @brief Get the outline status
         * @return true if the font is formatted with an outline effect (per ECMA-376 legacy macOS font styling)
         */
        bool outline() const;

        /**
         * @brief Get the shadow status
         * @return true if the font has a shadow effect
         */
        bool shadow() const;

        /**
         * @brief Get the condense status
         * @return true if the font is condensed
         */
        bool condense() const;

        /**
         * @brief Get the extend status
         * @return true if the font is extended (spacing between characters is increased)
         */
        bool extend() const;

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        XLFont& setFontName(std::string_view newName);
        XLFont& setFontCharset(size_t newCharset);
        XLFont& setFontFamily(size_t newFamily);
        XLFont& setFontSize(size_t newSize);
        XLFont& setFontColor(XLColor newColor);
        XLFont& setBold(bool set = true);
        XLFont& setItalic(bool set = true);
        XLFont& setStrikethrough(bool set = true);
        XLFont& setUnderline(XLUnderlineStyle style = XLUnderlineSingle);
        XLFont& setScheme(XLFontSchemeStyle newScheme);
        XLFont& setVertAlign(XLVerticalAlignRunStyle newVertAlign);
        XLFont& setOutline(bool set = true);
        XLFont& setShadow(bool set = true);
        XLFont& setCondense(bool set = true);
        XLFont& setExtend(bool set = true);

        /**
         * @brief Return a string summary of the font properties
         * @return string with info about the font object
         */
        std::string summary() const;

    private:                                 // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_fontNode; /**< An XMLNode object with the font item */
    };

    /**
     * @brief An encapsulation of the XLSX fonts
     */
    class OPENXLSX_EXPORT XLFonts
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLFonts();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLFonts(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLFonts(const XLFonts& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLFonts(XLFonts&& other);

        /**
         * @brief
         */
        ~XLFonts();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLFonts& operator=(const XLFonts& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLFonts& operator=(XLFonts&& other) noexcept = default;

        /**
         * @brief Get the count of fonts
         * @return The amount of font entries
         */
        size_t count() const;

        /**
         * @brief Get the font identified by index
         * @param index The index within the XML sequence
         * @return An XLFont object
         */
        XLFont fontByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to fontByIndex
         * @param index The index within the XML sequence
         * @return An XLFont object
         */
        XLFont operator[](XLStyleIndex index) const { return fontByIndex(index); }

        /**
         * @brief Append a new XLFont, based on copyFrom, and return its index in fonts node
         * @param copyFrom Can provide an XLFont to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLFont copyFrom = XLFont{}, std::string_view styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

        /**
         * @brief Find an existing font matching copyFrom's properties, or create a new one.
         * @details Uses a canonical XML fingerprint for O(1) cache lookups after the first call.
         *          Prevents duplicate font entries when the same style is applied to many cells.
         * @param copyFrom The font descriptor to match or create.
         * @param styleEntriesPrefix XML indentation prefix for newly created nodes.
         * @return The index of the matching or newly created font.
         */
        XLStyleIndex findOrCreate(XLFont copyFrom, std::string_view styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:                                                               // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode>                              m_fontsNode; /**< An XMLNode object with the fonts item */
        std::vector<XLFont>                                   m_fonts;
        mutable std::unordered_map<std::string, XLStyleIndex> m_fingerprintCache; /**< fingerprint -> index dedup cache */
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLSTYLES_FONTS_HPP

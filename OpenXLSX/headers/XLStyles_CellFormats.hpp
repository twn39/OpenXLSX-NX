#ifndef OPENXLSX_XLSTYLES_CELLFORMATS_HPP
#define OPENXLSX_XLSTYLES_CELLFORMATS_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

#include "XLStyles_Common.hpp"
#include "XLXmlFile.hpp"
#include "XLXmlParser.hpp"

#include <memory>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>


namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLAlignment
    {
        friend class XLStyles;
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLAlignment();

        /**
         * @brief Constructor. New items should only be created through an XLBorder object.
         * @param node An XMLNode object with the alignment XMLNode. If no input is provided, a null node is used.
         */
        explicit XLAlignment(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLAlignment(const XLAlignment& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLAlignment(XLAlignment&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLAlignment();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLAlignment& operator=(const XLAlignment& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLAlignment& operator=(XLAlignment&& other) noexcept = default;

        /**
         * @brief Get the horizontal alignment
         * @return An XLAlignmentStyle
         */
        XLAlignmentStyle horizontal() const;

        /**
         * @brief Get the vertical alignment
         * @return An XLAlignmentStyle
         */
        XLAlignmentStyle vertical() const;

        /**
         * @brief Get text rotation
         * @return A value indicating rotation: 0 to 90 is counter-clockwise, 90 to 180 is 90 degrees offset clockwise (e.g. 180 = 90 deg
         * clockwise), 255 is vertical text.
         */
        uint16_t textRotation() const;

        /**
         * @brief Check whether text wrapping is enabled
         * @return true if enabled, false otherwise
         */
        bool wrapText() const;

        /**
         * @brief Get the indent setting
         * @return An integer value, where an increment of 1 represents 3 spaces.
         */
        uint32_t indent() const;

        /**
         * @brief Get the relative indent setting
         * @return An integer value, where an increment of 1 represents 1 space, in addition to indent()*3 spaces
         * @note can be negative
         */
        int32_t relativeIndent() const;

        /**
         * @brief Check whether justification of last line is enabled
         * @return true if enabled, false otherwise
         */
        bool justifyLastLine() const;

        /**
         * @brief Check whether shrink to fit is enabled
         * @return true if enabled, false otherwise
         */
        bool shrinkToFit() const;

        /**
         * @brief Get the reading order setting
         * @return An integer value: 0 - Context Dependent, 1 - Left-to-Right, 2 - Right-to-Left (any other value should be invalid)
         */
        uint32_t readingOrder() const;

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setHorizontal(XLAlignmentStyle newStyle);
        bool setVertical(XLAlignmentStyle newStyle);
        /**
         * @details on setTextRotation from XLSX specification:
         * Text rotation in cells. Expressed in degrees. Values range from 0 to 180. The first letter of the text
         *  is considered the center-point of the arc.
         * For 0 - 90, the value represents degrees above horizon. For 91-180 the degrees below the horizon is calculated as:
         * [degrees below horizon] = 90 - [newRotation].
         * Examples: setTextRotation( 45): / (text is formatted along a line from lower left to upper right)
         *           setTextRotation(135): \ (text is formatted along a line from upper left to lower right)
         */
        bool setTextRotation(uint16_t newRotation);
        bool setWrapText(bool set = true);
        bool setIndent(uint32_t newIndent);
        bool setRelativeIndent(int32_t newIndent);
        bool setJustifyLastLine(bool set = true);
        bool setShrinkToFit(bool set = true);
        bool setReadingOrder(
            uint32_t newReadingOrder);    // can be used with XLReadingOrderContextual, XLReadingOrderLeftToRight, XLReadingOrderRightToLeft

        /**
         * @brief Return a string summary of the alignment properties
         * @return string with info about the alignment object
         */
        std::string summary() const;

    private:                                      // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_alignmentNode; /**< An XMLNode object with the alignment item */
    };

    /**
     * @brief An encapsulation of a cell format item
     */
    class OPENXLSX_EXPORT XLCellFormat
    {
        friend class XLCellFormats;    // for access to m_cellFormatNode in XLCellFormats::create
        friend class XLStyles;
    public:                            // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLCellFormat();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the xf XMLNode. If no input is provided, a null node is used.
         * @param permitXfId true (XLPermitXfID) -> getter xfId and setter setXfId are enabled, otherwise will throw XLException if invoked
         */
        explicit XLCellFormat(const XMLNode& node, bool permitXfId);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLCellFormat(const XLCellFormat& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLCellFormat(XLCellFormat&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLCellFormat();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLCellFormat& operator=(const XLCellFormat& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLCellFormat& operator=(XLCellFormat&& other) noexcept = default;

        /**
         * @brief Get the number format id
         * @return The identifier of a number format, built-in (predefined by office) or defind in XLNumberFormats
         */
        uint32_t numberFormatId() const;

        /**
         * @brief Get the font index
         * @return The index(!) of a font as defined in XLFonts
         */
        XLStyleIndex fontIndex() const;

        /**
         * @brief Get the fill index
         * @return The index(!) of a fill as defined in XLFills
         */
        XLStyleIndex fillIndex() const;

        /**
         * @brief Get the border index
         * @return The index(!) of a border as defined in XLBorders
         */
        XLStyleIndex borderIndex() const;

        /**
         * @brief Get the id of a referred <xf> entry
         * @return The id referring to an index in cell style formats (cellStyleXfs)
         * @throw XLException when invoked from cellStyleFormats
         * @note - only permitted for cellFormats
         */
        XLStyleIndex xfId() const;

        /**
         * @brief Report whether number format is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyNumberFormat() const;

        /**
         * @brief Report whether font is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyFont() const;

        /**
         * @brief Report whether fill is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyFill() const;

        /**
         * @brief Report whether border is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyBorder() const;

        /**
         * @brief Report whether alignment is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyAlignment() const;

        /**
         * @brief Report whether protection is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyProtection() const;

        /**
         * @brief Report whether quotePrefix is enabled
         * @return true for a setting enabled, or false if disabled
         * @note from documentation: A boolean value indicating whether the text string in a cell should be prefixed by a single quote mark
         *       (e.g., 'text). In these cases, the quote is not stored in the Shared Strings Part.
         */
        bool quotePrefix() const;

        /**
         * @brief Report whether pivot button is applied
         * @return true for a setting enabled, or false if disabled
         * @note from documentation: A boolean value indicating whether the cell rendering includes a pivot table dropdown button.
         * @todo need to find a use case for this
         */
        bool pivotButton() const;

        /**
         * @brief Report whether protection locked is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool locked() const;

        /**
         * @brief Report whether protection hidden is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool hidden() const;

        /**
         * @brief Return a reference to applicable alignment
         * @param createIfMissing triggers creation of alignment node - should be used with setter functions of XLAlignment
         * @return An XLAlignment object reference
         */
        XLAlignment alignment(bool createIfMissing = false) const;

        /**
         * @brief Unsupported getter
         */
        XLUnsupportedElement extLst() const { return XLUnsupportedElement{}; }    // <xf><extLst>...</extLst></xf>

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        XLCellFormat& setNumberFormatId(uint32_t newNumFmtId);
        XLCellFormat& setFontIndex(XLStyleIndex newFontIndex);
        XLCellFormat& setFillIndex(XLStyleIndex newFillIndex);
        XLCellFormat& setBorderIndex(XLStyleIndex newBorderIndex);
        XLCellFormat& setXfId(XLStyleIndex newXfId);    // NOTE: throws when invoked from cellStyleFormats
        XLCellFormat& setApplyNumberFormat(bool set = true);
        XLCellFormat& setApplyFont(bool set = true);
        XLCellFormat& setApplyFill(bool set = true);
        XLCellFormat& setApplyBorder(bool set = true);
        XLCellFormat& setApplyAlignment(bool set = true);
        XLCellFormat& setApplyProtection(bool set = true);
        XLCellFormat& setQuotePrefix(bool set = true);
        XLCellFormat& setPivotButton(bool set = true);
        XLCellFormat& setLocked(bool set = true);
        XLCellFormat& setHidden(bool set = true);
        /**
         * @brief Unsupported setter
         */
        XLCellFormat& setExtLst(XLUnsupportedElement const& newExtLst);

        /**
         * @brief Return a string summary of the cell format properties
         * @return string with info about the cell format object
         */
        std::string summary() const;

    private:                                                                // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode>                          m_cellFormatNode; /**< An XMLNode object with the cell format (xf) item */
        bool                                              m_permitXfId{false};
        inline static const std::vector<std::string_view> m_nodeOrder = {"alignment", "protection"};
    };

    /**
     * @brief An encapsulation of the XLSX cell style formats
     */
    class OPENXLSX_EXPORT XLCellFormats
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLCellFormats();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the cell formats (cellXfs or cellStyleXfs) item. If no input is provided, a null node is used.
         * @param permitXfId Pass-through to XLCellFormat constructor: true (XLPermitXfID) -> setter setXfId is enabled, otherwise throws
         */
        explicit XLCellFormats(const XMLNode& node, bool permitXfId = false);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLCellFormats(const XLCellFormats& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLCellFormats(XLCellFormats&& other);

        /**
         * @brief
         */
        ~XLCellFormats();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLCellFormats& operator=(const XLCellFormats& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLCellFormats& operator=(XLCellFormats&& other) noexcept = default;

        /**
         * @brief Get the count of cell style format descriptions
         * @return The amount of cell style format description entries
         */
        size_t count() const;

        /**
         * @brief Get the cell style format description identified by index
         * @param index The index within the XML sequence
         * @return An XLCellFormat object
         */
        XLCellFormat cellFormatByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to cellFormatByIndex
         * @param index The index within the XML sequence
         * @return An XLCellFormat object
         */
        XLCellFormat operator[](XLStyleIndex index) const { return cellFormatByIndex(index); }

        /**
         * @brief Append a new XLCellFormat, based on copyFrom, and return its index
         *       in cellXfs (for XLStyles::cellFormats) or cellStyleXfs (for XLStyles::cellStyleFormats)
         * @param copyFrom Can provide an XLCellFormat to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLCellFormat copyFrom = XLCellFormat{}, std::string_view styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

        /**
         * @brief Find an existing cell format matching copyFrom's properties, or create a new one.
         * @details Uses a canonical XML fingerprint for O(1) cache lookups after the first call.
         *          This is the key deduplication point for bulk formatting operations.
         */
        XLStyleIndex findOrCreate(XLCellFormat copyFrom, std::string_view styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:                                                                     // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode>                              m_cellFormatsNode; /**< An XMLNode object with the cell formats item */
        std::vector<XLCellFormat>                             m_cellFormats;
        bool                                                  m_permitXfId{false};
        mutable std::unordered_map<std::string, XLStyleIndex> m_fingerprintCache; /**< fingerprint -> index dedup cache */
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLSTYLES_CELLFORMATS_HPP

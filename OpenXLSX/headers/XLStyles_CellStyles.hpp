#ifndef OPENXLSX_XLSTYLES_CELLSTYLES_HPP
#define OPENXLSX_XLSTYLES_CELLSTYLES_HPP

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
#include <vector>


namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLCellStyle
    {
        friend class XLCellStyles;    // for access to m_cellStyleNode in XLCellStyles::create

    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLCellStyle();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the cellStyle item. If no input is provided, a null node is used.
         */
        explicit XLCellStyle(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLCellStyle(const XLCellStyle& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLCellStyle(XLCellStyle&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLCellStyle();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLCellStyle& operator=(const XLCellStyle& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLCellStyle& operator=(XLCellStyle&& other) noexcept = default;

        /**
         * @brief Test if this is an empty node
         * @return true if underlying XMLNode is empty
         */
        bool empty() const;

        /**
         * @brief Get the name of the cell style
         * @return The name for this cell style entry
         */
        std::string name() const;

        /**
         * @brief Get the id of the cell style format
         * @return The id referring to an index in cell style formats (cellStyleXfs)
         */
        XLStyleIndex xfId() const;

        /**
         * @brief Get the built-in id of the cell style
         * @return The built-in id of the cell style
         * @todo need to find a use case for this
         */
        uint32_t builtinId() const;

        /**
         * @brief Get the outline style id (attribute iLevel) of the cell style
         * @return The outline style id of the cell style
         * @todo need to find a use case for this
         */
        uint32_t outlineStyle() const;

        /**
         * @brief Get the hidden flag of the cell style
         * @return The hidden flag status (true: applications should not show this style)
         */
        bool hidden() const;

        /**
         * @brief Get the custom buildin flag
         * @return true if this cell style shall customize a built-in style
         */
        bool customBuiltin() const;

        /**
         * @brief Unsupported getter
         */
        XLUnsupportedElement extLst() const { return XLUnsupportedElement{}; }    // <cellStyle><extLst>...</extLst></cellStyle>

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setName(std::string_view newName);
        bool setXfId(XLStyleIndex newXfId);
        bool setBuiltinId(uint32_t newBuiltinId);
        bool setOutlineStyle(uint32_t newOutlineStyle);
        bool setHidden(bool set = true);
        bool setCustomBuiltin(bool set = true);
        /**
         * @brief Unsupported setter
         */
        bool setExtLst(XLUnsupportedElement const& newExtLst);

        /**
         * @brief Return a string summary of the cell style properties
         * @return string with info about the cell style object
         */
        std::string summary() const;

    private:                                      // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_cellStyleNode; /**< An XMLNode object with the cell style item */
    };

    /**
     * @brief An encapsulation of the XLSX cell styles
     */
    class OPENXLSX_EXPORT XLCellStyles
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLCellStyles();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the cellStyles item. If no input is provided, a null node is used.
         */
        explicit XLCellStyles(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLCellStyles(const XLCellStyles& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLCellStyles(XLCellStyles&& other);

        /**
         * @brief
         */
        ~XLCellStyles();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLCellStyles& operator=(const XLCellStyles& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLCellStyles& operator=(XLCellStyles&& other) noexcept = default;

        /**
         * @brief Get the count of cell styles
         * @return The amount of entries in the cell styles
         */
        size_t count() const;

        /**
         * @brief Get the cell style identified by index
         * @return An XLCellStyle object
         */
        XLCellStyle cellStyleByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to cellStyleByIndex
         * @param index The index within the XML sequence
         * @return An XLCellStyle object
         */
        XLCellStyle operator[](XLStyleIndex index) const { return cellStyleByIndex(index); }

        /**
         * @brief Append a new XLCellStyle, based on copyFrom, and return its index in cellStyles node
         * @param copyFrom Can provide an XLCellStyle to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLCellStyle copyFrom = XLCellStyle{}, std::string_view styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:                                       // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_cellStylesNode; /**< An XMLNode object with the cell styles item */
        std::vector<XLCellStyle> m_cellStyles;
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLSTYLES_CELLSTYLES_HPP

#ifndef OPENXLSX_XLSTYLES_BORDERS_HPP
#define OPENXLSX_XLSTYLES_BORDERS_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

#include "XLStyles_Common.hpp"
#include "XLStyles_Fills.hpp"    // XLDataBarColor (shared CT_Color wrapper used by line color)
#include "XLColor.hpp"
#include "XLXmlParser.hpp"

#include <memory>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>


namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLLine
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLLine();

        /**
         * @brief Constructor. New items should only be created through an XLBorder object.
         * @param node An XMLNode object with the line XMLNode. If no input is provided, a null node is used.
         */
        explicit XLLine(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLLine(const XLLine& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLLine(XLLine&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLLine();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLLine& operator=(const XLLine& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLLine& operator=(XLLine&& other) noexcept = default;

        /**
         * @brief Get the line style
         * @return An XLLineStyle enum
         */
        XLLineStyle style() const;

        /**
         * @brief Evaluate XLLine as bool
         * @return true if line is set, false if not
         */
        explicit operator bool() const;

        XLDataBarColor
            color() const;    // <line><color /></line> where node can be left, right, top, bottom, diagonal, vertical, horizontal

        /**
         * @note Regarding setter functions for style parameters:
         * @note Please refer to XLBorder setLine / setLeft / setRight / setTop / setBottom / setDiagonal
         * @note  and XLDataBarColor setter functions that can be invoked via color()
         */

        /**
         * @brief Return a string summary of the line properties
         * @return string with info about the line object
         */
        std::string summary() const;

    private:                                 // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_lineNode; /**< An XMLNode object with the line item */
    };

    /**
     * @brief An encapsulation of a border item
     */
    class OPENXLSX_EXPORT XLBorder
    {
        friend class XLBorders;    // for access to m_borderNode in XLBorders::create
        friend class XLStyles;
    public:                        // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLBorder();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the border XMLNode. If no input is provided, a null node is used.
         */
        explicit XLBorder(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLBorder(const XLBorder& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLBorder(XLBorder&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLBorder();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLBorder& operator=(const XLBorder& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLBorder& operator=(XLBorder&& other) noexcept = default;

        /**
         * @brief Get the diagonal up property
         * @return true if set, otherwise false
         */
        bool diagonalUp() const;

        /**
         * @brief Get the diagonal down property
         * @return true if set, otherwise false
         */
        bool diagonalDown() const;

        /**
         * @brief Get the outline property
         * @return true if set, otherwise false
         */
        bool outline() const;

        /**
         * @brief Get the left line property
         * @return An XLLine object
         */
        XLLine left() const;

        /**
         * @brief Get the left line property
         * @return An XLLine object
         */
        XLLine right() const;

        /**
         * @brief Get the left line property
         * @return An XLLine object
         */
        XLLine top() const;

        /**
         * @brief Get the bottom line property
         * @return An XLLine object
         */
        XLLine bottom() const;

        /**
         * @brief Get the diagonal line property
         * @return An XLLine object
         */
        XLLine diagonal() const;

        /**
         * @brief Get the vertical line property
         * @return An XLLine object
         */
        XLLine vertical() const;

        /**
         * @brief Get the horizontal line property
         * @return An XLLine object
         */
        XLLine horizontal() const;

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @param value2 (optional) that shall be set
         * @return true for success, false for failure
         */
        XLBorder& setDiagonalUp(bool set = true);
        XLBorder& setDiagonalDown(bool set = true);
        XLBorder& setOutline(bool set = true);
        bool      setLine(XLLineType lineType, XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool      setLeft(XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool      setRight(XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool      setTop(XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool      setBottom(XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool      setDiagonal(XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool      setVertical(XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool      setHorizontal(XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);

        /**
         * @brief Return a string summary of the font properties
         * @return string with info about the font object
         */
        std::string summary() const;

    private:                                                            // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode>                          m_borderNode; /**< An XMLNode object with the font item */
        inline static const std::vector<std::string_view> m_nodeOrder =
            {"left", "right", "top", "bottom", "diagonal", "vertical", "horizontal"};
    };

    /**
     * @brief An encapsulation of the XLSX borders
     */
    class OPENXLSX_EXPORT XLBorders
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLBorders();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the borders item. If no input is provided, a null node is used.
         */
        explicit XLBorders(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLBorders(const XLBorders& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLBorders(XLBorders&& other);

        /**
         * @brief
         */
        ~XLBorders();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLBorders& operator=(const XLBorders& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLBorders& operator=(XLBorders&& other) noexcept = default;

        /**
         * @brief Get the count of border descriptions
         * @return The amount of border description entries
         */
        size_t count() const;

        /**
         * @brief Get the border description identified by index
         * @param index The index within the XML sequence
         * @return An XLBorder object
         */
        XLBorder borderByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to borderByIndex
         * @param index The index within the XML sequence
         * @return An XLBorder object
         */
        XLBorder operator[](XLStyleIndex index) const { return borderByIndex(index); }

        /**
         * @brief Append a new XLBorder, based on copyFrom, and return its index in borders node
         * @param copyFrom Can provide an XLBorder to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLBorder copyFrom = XLBorder{}, std::string_view styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

        /**
         * @brief Find an existing border matching copyFrom's properties, or create a new one.
         * @details Uses a canonical XML fingerprint for O(1) cache lookups after the first call.
         */
        XLStyleIndex findOrCreate(XLBorder copyFrom, std::string_view styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:                                                                 // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode>                              m_bordersNode; /**< An XMLNode object with the borders item */
        std::vector<XLBorder>                                 m_borders;
        mutable std::unordered_map<std::string, XLStyleIndex> m_fingerprintCache; /**< fingerprint -> index dedup cache */
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLSTYLES_BORDERS_HPP

#ifndef OPENXLSX_XLSTYLES_FILLS_HPP
#define OPENXLSX_XLSTYLES_FILLS_HPP

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
    class OPENXLSX_EXPORT XLDataBarColor
    {
    public:
        /**
         * @brief
         */
        XLDataBarColor();

        /**
         * @brief Constructor. New items should only be created through an XLGradientStop or XLLine object.
         * @param node An XMLNode object with a data bar color XMLNode. If no input is provided, a null node is used.
         */
        explicit XLDataBarColor(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLDataBarColor(const XLDataBarColor& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLDataBarColor(XLDataBarColor&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLDataBarColor() = default;

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLDataBarColor& operator=(const XLDataBarColor& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLDataBarColor& operator=(XLDataBarColor&& other) noexcept = default;

        /**
         * @brief Get the line color from the rgb attribute
         * @return An XLColor object
         */
        XLColor rgb() const;

        /**
         * @brief Get the line color tint
         * @return A double value as stored in the "tint" attribute (should be between [-1.0;+1.0]), 0.0 if attribute does not exist
         */
        double tint() const;

        /**
         * @brief currently unsupported getter stubs
         */
        bool     automatic() const;    // <color auto="true" />
        uint32_t indexed() const;      // <color indexed="1" />
        uint32_t theme() const;        // <color theme="1" />

        /**
         * @brief Setter functions for data bar color parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setRgb(XLColor newColor);
        bool set(XLColor newColor) { return setRgb(newColor); }    // alias for setRgb
        bool setTint(double newTint);
        bool setAutomatic(bool set = true);
        bool setIndexed(uint32_t newIndex);
        bool setTheme(uint32_t newTheme);

        /**
         * @brief Return a string summary of the color properties
         * @return string with info about the color object
         */
        std::string summary() const;

    private:                                  // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_colorNode; /**< An XMLNode object with the color item */
    };

    // XLFills Class

    /**
     * @brief An encapsulation of an fill::gradientFill::stop item
     */
    class OPENXLSX_EXPORT XLGradientStop
    {
        friend class XLGradientStops;    // for access to m_stopNode in XLGradientStops::create
    public:                              // ---------- Public Member Functions ----------- //
        /**
         * @brief
         */
        XLGradientStop();

        /**
         * @brief Constructor. New items should only be created through an XLGradientStops object.
         * @param node An XMLNode object with the gradient stop XMLNode. If no input is provided, a null node is used.
         */
        explicit XLGradientStop(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLGradientStop(const XLGradientStop& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLGradientStop(XLGradientStop&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLGradientStop() = default;

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLGradientStop& operator=(const XLGradientStop& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLGradientStop& operator=(XLGradientStop&& other) noexcept = default;

        /**
         * @brief Getter functions
         * @return The requested gradient stop property
         */
        XLDataBarColor color() const;       // <stop><color /></stop>
        double         position() const;    // <stop position="1.2" />

        /**
         * @brief Setter functions
         * @param value that shall be set
         * @return true for success, false for failure
         * @note for color setters, use the color() getter with the XLDataBarColor setter functions
         */
        bool setPosition(double newPosition);

        /**
         * @brief Return a string summary of the stop properties
         * @return string with info about the stop object
         */
        std::string summary() const;

    private:                                 // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_stopNode; /**< An XMLNode object with the stop item */
    };

    /**
     * @brief An encapsulation of an array of fill::gradientFill::stop items
     */
    class OPENXLSX_EXPORT XLGradientStops
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLGradientStops();

        /**
         * @brief Constructor. New items should only be created through an XLFill object.
         * @param node An XMLNode object with the gradientFill item. If no input is provided, a null node is used.
         */
        explicit XLGradientStops(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLGradientStops(const XLGradientStops& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLGradientStops(XLGradientStops&& other);

        /**
         * @brief
         */
        ~XLGradientStops();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLGradientStops& operator=(const XLGradientStops& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLGradientStops& operator=(XLGradientStops&& other) noexcept = default;

        /**
         * @brief Get the count of gradient stops
         * @return The amount of stop entries
         */
        size_t count() const;

        /**
         * @brief Get the gradient stop entry identified by index
         * @param index The index within the XML sequence
         * @return An XLGradientStop object
         */
        XLGradientStop stopByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to stopByIndex
         * @param index The index within the XML sequence
         * @return An XLGradientStop object
         */
        XLGradientStop operator[](XLStyleIndex index) const { return stopByIndex(index); }

        /**
         * @brief Append a new XLGradientStop, based on copyFrom, and return its index in fills node
         * @param copyFrom Can provide an XLGradientStop to use as template for the new style
         * @param stopEntriesPrefix Prefix the newly created stop XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLGradientStop copyFrom = XLGradientStop{}, std::string_view styleEntriesPrefix = "");

        std::string summary() const;

    private:
        std::unique_ptr<XMLNode>    m_gradientNode; /**< An XMLNode object with the gradientFill item */
        std::vector<XLGradientStop> m_gradientStops;
    };

    /**
     * @brief An encapsulation of a fill item
     */
    class OPENXLSX_EXPORT XLFill
    {
        friend class XLFills;    // for access to m_fillNode in XLFills::create
        friend class XLStyles;
    public:                      // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLFill();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the fill XMLNode. If no input is provided, a null node is used.
         */
        explicit XLFill(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLFill(const XLFill& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLFill(XLFill&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLFill();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLFill& operator=(const XLFill& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLFill& operator=(XLFill&& other) noexcept = default;

        /**
         * @brief Get the name of the element describing a fill
         * @return The XLFillType derived from the name of the first child element of the fill node
         */
        XLFillType fillType() const;

        /**
         * @brief Create & set the base XML element describing the fill
         * @param newFillType that shall be set
         * @param force erase an existing fillType() if not equal newFillType
         * @return true for success, false for failure
         */
        bool setFillType(XLFillType newFillType, bool force = false);

    private:    // ---------- Switch to private context for two Methods  --------- //
        /**
         * @brief Throw XLException if fillType() matches typeToThrowOn
         * @param typeToThrowOn throw on this
         * @param functionName include this (calling function name) in the exception
         */
        void throwOnFillType(XLFillType typeToThrowOn, const char* functionName) const;

        /**
         * @brief Fetch a valid first element child of m_fillNode. Create with default if needed
         * @param fillTypeIfEmpty if no conflicting fill type exists, create a node with this fill type
         * @param functionName include this (calling function name) in a potential exception
         * @return An XMLNode containing a fill description
         * @throw XLException if fillTypeIfEmpty is in conflict with a current fillType()
         */
        XMLNode getValidFillDescription(XLFillType fillTypeIfEmpty, const char* functionName);

    public:    // ---------- Switch back to public context ---------------------- //
        /**
         * @brief Getter functions for gradientFill - will throwOnFillType(XLPatternFill, __func__)
         * @return The requested gradientFill property
         */
        XLGradientType  gradientType();    // <gradientFill type="path" />
        double          degree();
        double          left();
        double          right();
        double          top();
        double          bottom();
        XLGradientStops stops();

        /**
         * @brief Getter functions for patternFill - will throwOnFillType(XLGradientFill, __func__)
         * @return The requested patternFill property
         */
        XLPatternType patternType();
        XLColor       color();
        XLColor       backgroundColor();

        /**
         * @brief Setter functions for gradientFill - will throwOnFillType(XLPatternFill, __func__)
         * @param value that shall be set
         * @return true for success, false for failure
         * @note for gradient stops, use the stops() getter with the XLGradientStops access functions (create, stopByIndex, [])
         *       and the XLGradientStop setter functions
         */
        bool setGradientType(XLGradientType newType);
        bool setDegree(double newDegree);
        bool setLeft(double newLeft);
        bool setRight(double newRight);
        bool setTop(double newTop);
        bool setBottom(double newBottom);

        /**
         * @brief Setter functions for patternFill - will throwOnFillType(XLGradientFill, __func__)
         * @param value that shall be set
         * @return true for success, false for failure
         */
        XLFill& setPatternType(XLPatternType newPatternType);
        XLFill& setColor(XLColor newColor);
        XLFill& setBackgroundColor(XLColor newBgColor);

        /**
         * @brief Return a string summary of the fill properties
         * @return string with info about the fill object
         */
        std::string summary();

    private:                                 // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_fillNode; /**< An XMLNode object with the fill item */
    };

    /**
     * @brief An encapsulation of the XLSX fills
     */
    class OPENXLSX_EXPORT XLFills
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLFills();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the fills item. If no input is provided, a null node is used.
         */
        explicit XLFills(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLFills(const XLFills& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLFills(XLFills&& other);

        /**
         * @brief
         */
        ~XLFills();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLFills& operator=(const XLFills& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLFills& operator=(XLFills&& other) noexcept = default;

        /**
         * @brief Get the count of fills
         * @return The amount of fill entries
         */
        size_t count() const;

        /**
         * @brief Get the fill entry identified by index
         * @param index The index within the XML sequence
         * @return An XLFill object
         */
        XLFill fillByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to fillByIndex
         * @param index The index within the XML sequence
         * @return An XLFill object
         */
        XLFill operator[](XLStyleIndex index) const { return fillByIndex(index); }

        /**
         * @brief Append a new XLFill, based on copyFrom, and return its index in fills node
         * @param copyFrom Can provide an XLFill to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLFill copyFrom = XLFill{}, std::string_view styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

        /**
         * @brief Find an existing fill matching copyFrom's properties, or create a new one.
         * @details Uses a canonical XML fingerprint for O(1) cache lookups after the first call.
         */
        XLStyleIndex findOrCreate(XLFill copyFrom, std::string_view styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:                                                               // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode>                              m_fillsNode; /**< An XMLNode object with the fills item */
        std::vector<XLFill>                                   m_fills;
        mutable std::unordered_map<std::string, XLStyleIndex> m_fingerprintCache; /**< fingerprint -> index dedup cache */
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLSTYLES_FILLS_HPP

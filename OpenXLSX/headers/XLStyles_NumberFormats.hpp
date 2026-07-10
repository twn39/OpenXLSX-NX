#ifndef OPENXLSX_XLSTYLES_NUMBERFORMATS_HPP
#define OPENXLSX_XLSTYLES_NUMBERFORMATS_HPP

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
    class OPENXLSX_EXPORT XLNumberFormat
    {
        friend class XLNumberFormats;    // for access to m_numberFormatNode in XLNumberFormats::create
    public:                              // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLNumberFormat();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLNumberFormat(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLNumberFormat(const XLNumberFormat& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLNumberFormat(XLNumberFormat&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLNumberFormat();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLNumberFormat& operator=(const XLNumberFormat& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLNumberFormat& operator=(XLNumberFormat&& other) noexcept = default;

        /**
         * @brief Get the id of the number format
         * @return The id for this number format
         */
        uint32_t numberFormatId() const;

        /**
         * @brief Get the code of the number format
         * @return The format code for this number format
         */
        std::string formatCode() const;

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setNumberFormatId(uint32_t newNumberFormatId);
        bool setFormatCode(std::string_view newFormatCode);

        /**
         * @brief Return a string summary of the number format
         * @return string with info about the number format object
         */
        std::string summary() const;

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_numberFormatNode; /**< An XMLNode object with the number format item */
    };

    /**
     * @brief An encapsulation of the XLSX number formats (numFmts)
     */
    class OPENXLSX_EXPORT XLNumberFormats
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLNumberFormats();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLNumberFormats(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLNumberFormats(const XLNumberFormats& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLNumberFormats(XLNumberFormats&& other);

        /**
         * @brief
         */
        ~XLNumberFormats();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLNumberFormats& operator=(const XLNumberFormats& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLNumberFormats& operator=(XLNumberFormats&& other) noexcept = default;

        /**
         * @brief Get the count of number formats
         * @return The amount of entries in the number formats
         */
        size_t count() const;

        /**
         * @brief Get the number format identified by index
         * @param index an array index within XLStyles::numberFormats()
         * @return An XLNumberFormat object
         * @throw XLException when index is out of m_numberFormats range
         */
        XLNumberFormat numberFormatByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to numberFormatByIndex
         */
        XLNumberFormat operator[](XLStyleIndex index) const { return numberFormatByIndex(index); }

        /**
         * @brief Get the number format identified by numberFormatId
         * @param numberFormatId a numFmtId reference to a number format
         * @return An XLNumberFormat object
         * @throw XLException if no match for numberFormatId is found within m_numberFormats
         */
        XLNumberFormat numberFormatById(uint32_t numberFormatId) const;

        /**
         * @brief Get the numFmtId from the number format identified by index
         * @param index an array index within XLStyles::numberFormats()
         * @return the uint32_t numFmtId corresponding to index
         * @throw XLException when index is out of m_numberFormats range
         */
        uint32_t numberFormatIdFromIndex(XLStyleIndex index) const;

        /**
         * @brief Create a new custom number format with a unique ID.
         * @param formatCode The explicit format code string
         * @return The unique number format ID (numFmtId)
         */
        uint32_t createNumberFormat(std::string_view formatCode);

        /**
         * @brief Get a free (unused) number format ID, skipping reserved IDs (0-163).
         * @return A free ID starting from 164.
         */
        uint32_t getFreeNumberFormatId() const;

        /**
         * @brief Append a new XLNumberFormat, based on copyFrom, and return its index in numFmts node
         * @param copyFrom Can provide an XLNumberFormat to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLNumberFormat copyFrom = XLNumberFormat{}, std::string_view styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:                                             // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode>    m_numberFormatsNode; /**< An XMLNode object with the number formats item */
        std::vector<XLNumberFormat> m_numberFormats;
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLSTYLES_NUMBERFORMATS_HPP

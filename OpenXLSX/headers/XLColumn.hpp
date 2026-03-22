#ifndef OPENXLSX_XLCOLUMN_HPP
#define OPENXLSX_XLCOLUMN_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <memory>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLStyles.hpp"    // XLStyleIndex
#include "XLXmlParser.hpp"
#include "XLStyle.hpp"
#include "XLStyles.hpp"

namespace OpenXLSX
{
    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLColumn
    {
    public:
        /**
         * @brief Constructor
         * @param columnNode The XMLNode for the column.
         */
        explicit XLColumn(XMLNode columnNode);

        /**
         * @brief Get the width of the column.
         * @return The width of the column.
         */
        float width() const;

        /**
         * @brief Set the width of the column
         * @param width The width of the column
         */
        void setWidth(float width);

        /**
         * @brief Auto fit the column width based on the content of its cells.
         * @note This is a convenience method that internally delegates to XLWorksheet.
         */
        void autoFit();

        /**
         * @brief Is the column hidden?
         * @return The state of the column.
         */
        bool isHidden() const;

        /**
         * @brief Set the column to be shown or hidden.
         * @param state The state of the column.
         */
        void setHidden(bool state);

        /**
         * @brief Get the outline level of the column.
         * @return The outline level (0-7).
         */
        uint8_t outlineLevel() const;

        /**
         * @brief Set the outline level of the column.
         * @param level The outline level (0-7).
         */
        void setOutlineLevel(uint8_t level);

        /**
         * @brief Is the column collapsed?
         * @return true if the column is collapsed.
         */
        bool isCollapsed() const;

        /**
         * @brief Set the column to be collapsed or expanded.
         * @param state The collapsed state of the column.
         */
        void setCollapsed(bool state);

        /**
         * @brief Get the XMLNode object for the column.
         * @return The XMLNode for the column
         */
        XMLNode columnNode() const;

        /**
         * @brief Get the array index of xl/styles.xml:<styleSheet>:<cellXfs> for the style assigned to the column.
         *        This value is stored in the col attributes like so: style="2"
         * @returns The index of the applicable format style
         */
        XLStyleIndex format() const;

        /**
         * @brief Set the column style as a reference to the array index of xl/styles.xml:<styleSheet>:<cellXfs>
         * @param cellFormatIndex The style to set, corresponding to the index of XLStyles::cellStyles()
         * @returns true on success, false on failure
         */
        bool setFormat(XLStyleIndex cellFormatIndex);

    private:
        /**
         * @brief Get or create an attribute of the column node.
         * @param attrName The name of the attribute.
         * @return The XMLAttribute object.
         */
        XMLAttribute getOrCreateAttribute(const char* attrName);

        XMLNode m_columnNode; /**< The XMLNode object for the column. */
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLCOLUMN_HPP

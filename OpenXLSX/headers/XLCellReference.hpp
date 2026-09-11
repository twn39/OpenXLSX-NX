/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLCELLREFERENCE_HPP
#define OPENXLSX_XLCELLREFERENCE_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <cstdint>    // Pull request #276
#include <string>
#include <string_view>
#include <utility>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"

namespace OpenXLSX
{
    /**
     * @brief
     */
    struct XLCoordinates
    {
        uint32_t row;
        uint16_t column;
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLCellReference final
    {
        friend bool operator==(const XLCellReference& lhs, const XLCellReference& rhs) noexcept;
        friend bool operator!=(const XLCellReference& lhs, const XLCellReference& rhs) noexcept;
        friend bool operator<(const XLCellReference& lhs, const XLCellReference& rhs) noexcept;
        friend bool operator>(const XLCellReference& lhs, const XLCellReference& rhs) noexcept;
        friend bool operator<=(const XLCellReference& lhs, const XLCellReference& rhs) noexcept;
        friend bool operator>=(const XLCellReference& lhs, const XLCellReference& rhs) noexcept;

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief Constructor taking a cell address as argument.
         * @param cellAddress The address of the cell, e.g. 'A1'.
         * @details Initializes a reference using an Excel-style coordinate string (e.g., 'A1'). Serves as the primary parser for cell
         * identities read directly from the DOM.
         */
        XLCellReference(std::string_view cellAddress = "");    // NOLINT

        /**
         * @brief Constructor taking the cell coordinates as arguments.
         * @param row The row number of the cell.
         * @param column The column number of the cell.
         */
        XLCellReference(uint32_t row, uint16_t column);

        /**
         * @brief Constructor taking the row number and the column letter as arguments.
         * @param row The row number of the cell.
         * @param column The column letter of the cell.
         */
        XLCellReference(uint32_t row, std::string_view column);

        /**
         * @brief Copy constructor
         * @param other The object to be copied.
         */
        XLCellReference(const XLCellReference& other);

        /**
         * @brief
         * @param other
         */
        XLCellReference(XLCellReference&& other) noexcept;

        /**
         * @brief Destructor. Default implementation used.
         */
        ~XLCellReference();

        /**
         * @brief Assignment operator.
         * @param other The object to be copied/assigned.
         * @return A reference to the new object.
         */
        XLCellReference& operator=(const XLCellReference& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLCellReference& operator=(XLCellReference&& other) noexcept;

        /**
         * @brief
         * @return
         */
        XLCellReference& operator++();

        /**
         * @brief
         * @return
         */
        XLCellReference operator++(int);    // NOLINT

        /**
         * @brief
         * @return
         */
        XLCellReference& operator--();

        /**
         * @brief
         * @return
         */
        XLCellReference operator--(int);    // NOLINT

        /**
         * @brief Get the row number of the XLCellReference.
         * @return The row.
         */
        uint32_t row() const noexcept;

        /**
         * @brief Set the row number for the XLCellReference.
         * @param row The row number.
         */
        void setRow(uint32_t row);

        /**
         * @brief Get the column number of the XLCellReference.
         * @return The column number.
         */
        uint16_t column() const noexcept;

        /**
         * @brief Set the column number of the XLCellReference.
         * @param column The column number.
         */
        void setColumn(uint16_t column);

        /**
         * @brief Set both row and column number of the XLCellReference.
         * @param row The row number.
         * @param column The column number.
         */
        void setRowAndColumn(uint32_t row, uint16_t column);

        /**
         * @brief Get the address of the XLCellReference
         * @return The address, e.g. 'A1'
         */
        std::string address() const;

        /**
         * @brief Set the address of the XLCellReference
         * @param address The address, e.g. 'A1'
         * @pre The address input string must be a valid Excel cell reference. Otherwise the behaviour is undefined.
         */
        void setAddress(std::string_view address);

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Functions
        //----------------------------------------------------------------------------------------------------------------------

        // private:

        /**
         * @brief
         * @param row
         * @return
         */
        static std::string rowAsString(uint32_t row);

        /**
         * @brief
         * @param row
         * @return
         */
        static uint32_t rowAsNumber(std::string_view row);

        /**
         * @brief Static helper function to convert column number to column letter (e.g. column 1 becomes 'A')
         * @param column The column number.
         * @return The column letter
         */
        static std::string columnAsString(uint16_t column);

        /**
         * @brief Static helper function to convert column letter to column number (e.g. column 'A' becomes 1)
         * @param column The column letter, e.g. 'A'
         * @return The column number.
         */
        static uint16_t columnAsNumber(std::string_view column);

        /**
         * @brief Static helper function to convert cell address to coordinates.
         * @param address The address to be converted, e.g. 'A1'
         * @return A std::pair<row, column>
         */
        static XLCoordinates coordinatesFromAddress(std::string_view address);

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------
    private:
        uint32_t    m_row{1};            /**< The row */
        uint16_t    m_column{1};         /**< The column */
        std::string m_cellAddress{"A1"}; /**< The address, e.g. 'A1' */
    };

    /**
     * @brief Determines exact coordinate identity, allowing comparisons without resolving their DOM node equivalence.
     */
    inline bool operator==(const XLCellReference& lhs, const XLCellReference& rhs) noexcept
    { return lhs.row() == rhs.row() and lhs.column() == rhs.column(); }

    /**
     * @brief Detects coordinate divergence, effectively identical to `!(lhs == rhs)`.
     */
    inline bool operator!=(const XLCellReference& lhs, const XLCellReference& rhs) noexcept { return !(lhs == rhs); }

    /**
     * @brief Evaluates precedence primarily by row, then by column, allowing cell ranges to be sorted efficiently and sequentially from
     * left-to-right, top-to-bottom.
     */
    inline bool operator<(const XLCellReference& lhs, const XLCellReference& rhs) noexcept
    { return lhs.row() < rhs.row() or (lhs.row() <= rhs.row() and lhs.column() < rhs.column()); }

    /**
     * @brief Inverts the less-than operator logic to verify strict left-to-right, top-to-bottom traversal dominance.
     */
    inline bool operator>(const XLCellReference& lhs, const XLCellReference& rhs) noexcept { return (rhs < lhs); }

    /**
     * @brief Asserts whether a cell sequentially precedes or occupies the exact same coordinate as another.
     */
    inline bool operator<=(const XLCellReference& lhs, const XLCellReference& rhs) noexcept { return !(lhs > rhs); }

    /**
     * @brief Asserts whether a cell sequentially follows or occupies the exact same coordinate as another.
     */
    inline bool operator>=(const XLCellReference& lhs, const XLCellReference& rhs) noexcept { return !(lhs < rhs); }
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLCELLREFERENCE_HPP

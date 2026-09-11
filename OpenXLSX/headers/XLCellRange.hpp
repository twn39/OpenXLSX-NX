/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLCELLRANGE_HPP
#define OPENXLSX_XLCELLRANGE_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

// ===== External Includes ===== //
#include <memory>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCell.hpp"
#include "XLCellIterator.hpp"
#include "XLCellReference.hpp"
#include "XLStyle.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @brief Represents a rectangular area of cells within a worksheet.
     * @details This class provides a high-level interface for bulk operations on cells,
     * such as clearing values, setting formats, or iterating over a subset of the worksheet.
     * It maintains a cache of column styles to ensure efficient cell creation during iteration.
     */
    class OPENXLSX_EXPORT XLCellRange
    {
        friend class XLCellIterator;

    public:
        /**
         * @brief Constructs an uninitialized range.
         */
        XLCellRange();

        /**
         * @param dataNode The XML node containing sheet data.
         * @param topLeft Top-left boundary of the range.
         * @param bottomRight Bottom-right boundary of the range.
         * @param sharedStrings Reference to the workbook's shared strings table.
         * @throws XLInputError if topLeft is not truly to the top-left of bottomRight.
         */
        explicit XLCellRange(const XMLNode&         dataNode,
                             const XLCellReference& topLeft,
                             const XLCellReference& bottomRight,
                             const XLSharedStrings& sharedStrings);

        ~XLCellRange()                                       = default;
        XLCellRange(const XLCellRange& other)                = default;
        XLCellRange(XLCellRange&& other) noexcept            = default;
        XLCellRange& operator=(const XLCellRange& other)     = default;
        XLCellRange& operator=(XLCellRange&& other) noexcept = default;

        /**
         * @brief Scans the worksheet for column-level styles and caches them.
         * @details This is called automatically during construction to ensure XLCellIterator
         * can apply default styles to newly created cells without repeated XML lookups.
         */
        void fetchColumnStyles();

        [[nodiscard]] XLCellReference topLeft() const;
        [[nodiscard]] XLCellReference bottomRight() const;

        /**
         * @return The range reference string (e.g., "A1:C3"), or an empty string if uninitialized.
         */
        [[nodiscard]] std::string address() const;

        [[nodiscard]] uint32_t numRows() const;
        [[nodiscard]] uint16_t numColumns() const;

        [[nodiscard]] XLCellIterator begin() const;
        [[nodiscard]] XLCellIterator end() const;

        /**
         * @brief Apply a high-level style to all cells within this range.
         * @param style The requested high-level style object.
         */
        void applyStyle(const XLStyle& style);

        /**
         * @brief Set border style specifically for the outer edges of this range.
         * @param style The line style (e.g. XLLineStyleThick)
         * @param color The line color (e.g. XLColor("000000"))
         */
        void setBorderOutline(XLLineStyle style, XLColor color);

        /**
         * @brief Read the range into a 2D matrix of type T.
         */
        template<typename T>
        std::vector<std::vector<T>> getValue() const;

        template<typename T>
        void setValue(const std::vector<std::vector<T>>& matrix);
        /**
         * @brief Returns true if the range is uninitialized or points to an invalid worksheet node.
         */
        [[nodiscard]] bool empty() const;

        /**
         * @brief Clears the values of all cells within the range.
         */
        void clear();

        /**
         * @brief Assigns a single value to every cell in the range.
         */
        template<typename T,
                 typename = std::enable_if_t<
                     std::is_integral_v<T> or std::is_floating_point_v<T> or std::is_same_v<std::decay_t<T>, std::string> ||
                     std::is_same_v<std::decay_t<T>, std::string_view> or std::is_same_v<std::decay_t<T>, const char*> ||
                     std::is_same_v<std::decay_t<T>, char*> or std::is_same_v<T, XLDateTime>>>
        XLCellRange& operator=(T value)
        {
            for (auto it = begin(); it != end(); ++it) it->value() = value;
            return *this;
        }

        /**
         * @brief Applies a cell format (style) to all cells in the range.
         * @param cellFormatIndex The index in the workbook's style sheet.
         * @return true on success.
         */
        XLCellRange& setFormat(XLStyleIndex cellFormatIndex);

        /**
         * @brief Calculates the intersection area between this range and another.
         * @return A new range representing the common area, or an empty range if they don't overlap.
         */
        [[nodiscard]] XLCellRange intersect(const XLCellRange& other) const;

    private:
        XMLNode                   m_dataNode;
        XLCellReference           m_topLeft;
        XLCellReference           m_bottomRight;
        XLSharedStringsRef        m_sharedStrings;
        std::vector<XLStyleIndex> m_columnStyles; /**< Cached column-level styles for fast cell creation. */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

// Template implementations
namespace OpenXLSX
{
    template<typename T>
    std::vector<std::vector<T>> XLCellRange::getValue() const
    {
        std::vector<std::vector<T>> matrix;
        if (numRows() == 0 || numColumns() == 0) return matrix;

        matrix.resize(numRows(), std::vector<T>(numColumns()));

        uint32_t startRow = topLeft().row();
        uint16_t startCol = topLeft().column();

        for (auto& cell : *this) {
            uint32_t r   = cell.cellReference().row() - startRow;
            uint16_t c   = cell.cellReference().column() - startCol;
            matrix[r][c] = cell.value().get<T>();
        }
        return matrix;
    }

    template<typename T>
    void XLCellRange::setValue(const std::vector<std::vector<T>>& matrix)
    {
        if (matrix.empty() || matrix[0].empty()) return;

        uint32_t startRow = topLeft().row();
        uint16_t startCol = topLeft().column();

        for (auto& cell : *this) {
            uint32_t r = cell.cellReference().row() - startRow;
            uint16_t c = cell.cellReference().column() - startCol;
            if (r < matrix.size() && c < matrix[r].size()) { cell.value() = matrix[r][c]; }
        }
    }
}    // namespace OpenXLSX
#endif    // OPENXLSX_XLCELLRANGE_HPP

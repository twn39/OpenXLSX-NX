/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLCELLITERATOR_HPP
#define OPENXLSX_XLCELLITERATOR_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

#include <algorithm>

#include "OpenXLSX-Exports.hpp"
#include "XLCell.hpp"
#include "XLCellReference.hpp"
#include "XLIterator.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @brief locate the XML row node within sheetDataNode for the row at rowNumber
     * @return the XMLNode pointing to the row, or an empty XMLNode if the row does not exist
     */
    XMLNode findRowNode(XMLNode sheetDataNode, uint32_t rowNumber);

    /**
     * @brief locate the XML cell node within rownode for the cell at columnNumber
     * @return the XMLNode pointing to the cell, or an empty XMLNode if the cell does not exist
     */
    XMLNode findCellNode(XMLNode rowNode, uint16_t columnNumber);

    /**
     * @brief A forward iterator for iterating over a range of cells.
     * @details The iterator performs lazy-loading of cells to minimize XML overhead.
     */
    class OPENXLSX_EXPORT XLCellIterator
    {
    public:
        using iterator_category = std::forward_iterator_tag;
        using value_type        = XLCell;
        using difference_type   = int64_t;
        using pointer           = XLCell*;
        using reference         = XLCell&;

        enum class CellStatus : uint8_t { NotLoaded = 0, NoSuchCell = 1, Loaded = 2 };

        /**
         * @param colStyles Vector of column styles used to initialize newly created cells.
         */
        explicit XLCellIterator(const XLCellRange& cellRange, XLIteratorLocation loc, std::vector<XLStyleIndex> const* colStyles);

        ~XLCellIterator()                                          = default;
        XLCellIterator(const XLCellIterator& other)                = default;
        XLCellIterator(XLCellIterator&& other) noexcept            = default;
        XLCellIterator& operator=(const XLCellIterator& other)     = default;
        XLCellIterator& operator=(XLCellIterator&& other) noexcept = default;

    private:
        static constexpr bool XLCreateIfMissing      = true;
        static constexpr bool XLDoNotCreateIfMissing = false;

        /**
         * @brief Updates the internal cell pointer to match the current coordinates.
         * @param createIfMissing If true, a new XML node will be inserted if the cell doesn't exist.
         */
        void updateCurrentCell(bool createIfMissing) const;

    public:
        XLCellIterator& operator++();
        XLCellIterator  operator++(int);

        [[nodiscard]] reference operator*();
        [[nodiscard]] pointer   operator->();

        [[nodiscard]] bool operator==(const XLCellIterator& rhs) const noexcept;
        [[nodiscard]] bool operator!=(const XLCellIterator& rhs) const noexcept;

        /**
         * @return true if the cell exists in the XML structure.
         */
        [[nodiscard]] bool cellExists() const;

        /**
         * @return true if the iterator has moved past the last cell in the range.
         */
        [[nodiscard]] bool endReached() const noexcept { return m_endReached; }

        /**
         * @return The number of cells between this iterator and 'last'.
         */
        [[nodiscard]] uint64_t distance(const XLCellIterator& last) const;

        /**
         * @return The A1-style address of the current cell.
         */
        [[nodiscard]] std::string address() const;

    private:
        mutable XMLNode                  m_dataNode;
        XLCellReference                  m_topLeft;
        XLCellReference                  m_bottomRight;
        XLSharedStringsRef               m_sharedStrings;
        bool                             m_endReached;
        mutable XMLNode                  m_hintNode; /**< Cached node to speed up subsequent row/cell lookups. */
        mutable uint32_t                 m_hintRow;
        mutable XLCell                   m_currentCell;
        mutable CellStatus               m_currentCellStatus;
        uint32_t                         m_currentRow;
        uint16_t                         m_currentColumn;
        std::vector<XLStyleIndex> const* m_colStyles;
    };

    inline std::ostream& operator<<(std::ostream& os, const XLCellIterator& it)
    {
        os << it.address();
        return os;
    }
}    // namespace OpenXLSX

namespace std
{
    using OpenXLSX::XLCellIterator;
    template<>
    inline std::iterator_traits<XLCellIterator>::difference_type distance<XLCellIterator>(XLCellIterator first, XLCellIterator last)
    { return static_cast<std::iterator_traits<XLCellIterator>::difference_type>(first.distance(last)); }
}    // namespace std

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLCELLITERATOR_HPP

/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

// ===== External Includes ===== //
#include <fmt/format.h>
#include <gsl/gsl>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCellRange.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLCellRange::XLCellRange()
    : m_dataNode(XMLNode{}),
      m_topLeft(XLCellReference("A1")),
      m_bottomRight(XLCellReference("A1")),
      m_sharedStrings(XLSharedStringsDefaulted),
      m_columnStyles{}
{}

/**
 * @details Validates that the boundaries form a logical rectangle.
 */
XLCellRange::XLCellRange(const XMLNode&         dataNode,
                         const XLCellReference& topLeft,
                         const XLCellReference& bottomRight,
                         const XLSharedStrings& sharedStrings)
    : m_dataNode(dataNode),
      m_topLeft(topLeft),
      m_bottomRight(bottomRight),
      m_sharedStrings(sharedStrings),
      m_columnStyles{}
{
    if (m_topLeft.row() > m_bottomRight.row() or m_topLeft.column() > m_bottomRight.column()) {
        throw XLInputError(
            fmt::format("XLCellRange constructor: topLeft ({}) does not point to a lower or equal row and column than bottomRight ({})",
                        topLeft.address(),
                        bottomRight.address()));
    }

    fetchColumnStyles();
}

/**
 * @details Column styles are cached to ensure XLCellIterator can initialize new cells
 * with the correct worksheet-level formatting without re-traversing the XML tree.
 */
void XLCellRange::fetchColumnStyles()
{
    if (m_dataNode.empty()) return;
    XMLNode cols = m_dataNode.parent().child("cols");
    if (cols.empty()) return;

    m_columnStyles.clear();
    uint16_t vecPos = 0;

    for (auto col : cols.children("col")) {
        uint16_t minCol = gsl::narrow_cast<uint16_t>(col.attribute("min").as_uint(0));
        uint16_t maxCol = gsl::narrow_cast<uint16_t>(col.attribute("max").as_uint(0));

        if (minCol > maxCol || minCol == 0 || maxCol == 0) {
            throw XLInputError(
                fmt::format("XLCellRange: invalid column range {}:{}", col.attribute("min").value(), col.attribute("max").value()));
        }

        if (maxCol > m_columnStyles.size()) { m_columnStyles.resize(maxCol, XLDefaultCellFormat); }

        // Fill non-explicitly styled columns with the default format.
        while (vecPos + 1 < minCol) { m_columnStyles[vecPos++] = XLDefaultCellFormat; }

        XLStyleIndex colStyle = col.attribute("style").as_uint(XLDefaultCellFormat);
        while (vecPos < maxCol) { m_columnStyles[vecPos++] = colStyle; }
    }
}

/**
 * @details
 */
XLCellReference XLCellRange::topLeft() const { return m_topLeft; }

/**
 * @details
 */
XLCellReference XLCellRange::bottomRight() const { return m_bottomRight; }

/**
 * @details
 */
std::string XLCellRange::address() const
{
    if (empty()) return "";
    return m_topLeft.address() + ":" + m_bottomRight.address();
}

/**
 * @details
 */
uint32_t XLCellRange::numRows() const
{
    if (empty()) return 0;
    return m_bottomRight.row() + 1 - m_topLeft.row();
}

/**
 * @details
 */
uint16_t XLCellRange::numColumns() const
{
    if (empty()) return 0;
    return m_bottomRight.column() + 1 - m_topLeft.column();
}

/**
 * @details
 */
XLCellIterator XLCellRange::begin() const { return XLCellIterator(*this, XLIteratorLocation::Begin, &m_columnStyles); }

/**
 * @details
 */
XLCellIterator XLCellRange::end() const { return XLCellIterator(*this, XLIteratorLocation::End, &m_columnStyles); }

/**
 * @details
 */
bool XLCellRange::empty() const { return m_dataNode.empty(); }

/**
 * @details
 */
void XLCellRange::clear()
{
    for (auto& cell : *this) cell.value().clear();
}

/**
 * @details
 */
XLCellRange& XLCellRange::setFormat(XLStyleIndex cellFormatIndex)
{
    for (auto& cell : *this) { cell.setCellFormat(cellFormatIndex); }
    return *this;
}

/**
 * @details Performs a coordinate-based intersection. If ranges are on different sheets,
 * the intersection is always empty.
 */
XLCellRange XLCellRange::intersect(const XLCellRange& other) const
{
    if (empty() || other.empty() || m_dataNode != other.m_dataNode) return XLCellRange();

    uint32_t top    = std::max(m_topLeft.row(), other.m_topLeft.row());
    uint16_t left   = std::max(m_topLeft.column(), other.m_topLeft.column());
    uint32_t bottom = std::min(m_bottomRight.row(), other.m_bottomRight.row());
    uint16_t right  = std::min(m_bottomRight.column(), other.m_bottomRight.column());

    if (top > bottom || left > right) return XLCellRange();

    return XLCellRange(m_dataNode, XLCellReference(top, left), XLCellReference(bottom, right), m_sharedStrings.get());
}

void XLCellRange::applyStyle(const XLStyle& style)
{
    for (auto& cell : *this) { cell.setStyle(style); }
}

void XLCellRange::setBorderOutline(XLLineStyle style, XLColor color)
{
    uint32_t top    = topLeft().row();
    uint32_t bottom = bottomRight().row();
    uint16_t left   = topLeft().column();
    uint16_t right  = bottomRight().column();

    for (auto& cell : *this) {
        uint32_t r = cell.cellReference().row();
        uint16_t c = cell.cellReference().column();

        XLStyle cellStyle;
        bool    isEdge = false;

        if (r == top) {
            cellStyle.border.top.style = style;
            cellStyle.border.top.color = color;
            isEdge                     = true;
        }
        if (r == bottom) {
            cellStyle.border.bottom.style = style;
            cellStyle.border.bottom.color = color;
            isEdge                        = true;
        }
        if (c == left) {
            cellStyle.border.left.style = style;
            cellStyle.border.left.color = color;
            isEdge                      = true;
        }
        if (c == right) {
            cellStyle.border.right.style = style;
            cellStyle.border.right.color = color;
            isEdge                       = true;
        }

        if (isEdge) { cell.setStyle(cellStyle); }
    }
}

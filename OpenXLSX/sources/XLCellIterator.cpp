// ===== External Includes ===== //
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCellIterator.hpp"
#include "XLCellRange.hpp"
#include "XLCellReference.hpp"
#include "XLException.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{
    /**
     * @details
     */
    XMLNode findRowNode(XMLNode sheetDataNode, uint32_t rowNumber)
    {
        if (rowNumber < 1 or rowNumber > OpenXLSX::MAX_ROWS) {
            using namespace std::literals::string_literals;
            throw XLCellAddressError("rowNumber "s + std::to_string(rowNumber) + " is outside valid range [1;"s +
                                     std::to_string(OpenXLSX::MAX_ROWS) + "]"s);
        }

        XMLNode rowNode = sheetDataNode.last_child_of_type(pugi::node_element);

        // If sheetData is empty or the requested row is beyond existing rows, it definitely doesn't exist.
        if (rowNode.empty() or (rowNumber > rowNode.attribute("r").as_ullong())) return XMLNode{};

        // Determine the most efficient search direction (from start or from end).
        if (rowNode.attribute("r").as_ullong() - rowNumber < rowNumber) {
            while (not rowNode.empty() and (rowNode.attribute("r").as_ullong() > rowNumber))
                rowNode = rowNode.previous_sibling_of_type(pugi::node_element);
            if (rowNode.empty() or (rowNode.attribute("r").as_ullong() != rowNumber)) return XMLNode{};
        }
        else {
            rowNode = sheetDataNode.first_child_of_type(pugi::node_element);

            while (rowNode.attribute("r").as_ullong() < rowNumber) rowNode = rowNode.next_sibling_of_type(pugi::node_element);
            if (rowNode.attribute("r").as_ullong() > rowNumber) return XMLNode{};
        }

        return rowNode;
    }

    /**
     * @details
     */
    XMLNode findCellNode(XMLNode rowNode, uint16_t columnNumber)
    {
        if (columnNumber < 1 or columnNumber > OpenXLSX::MAX_COLS) {
            using namespace std::literals::string_literals;
            throw XLException("XLWorksheet::column: columnNumber "s + std::to_string(columnNumber) + " is outside allowed range [1;"s +
                              std::to_string(MAX_COLS) + "]"s);
        }
        if (rowNode.empty()) return XMLNode{};

        XMLNode cellNode = rowNode.last_child_of_type(pugi::node_element);

        // Using extractColumnFromCellRef is faster than constructing a full XLCellReference object.
        uint16_t lastCellCol = cellNode.empty() ? 0 : extractColumnFromCellRef(cellNode.attribute("r").value());

        if (cellNode.empty() or (lastCellCol < columnNumber)) return XMLNode{};

        // Bi-directional search for efficiency.
        if (lastCellCol - columnNumber < columnNumber) {
            uint16_t currentCol = lastCellCol;
            while (not cellNode.empty() and (currentCol > columnNumber)) {
                cellNode   = cellNode.previous_sibling_of_type(pugi::node_element);
                currentCol = cellNode.empty() ? 0 : extractColumnFromCellRef(cellNode.attribute("r").value());
            }
            if (cellNode.empty() or (currentCol < columnNumber)) return XMLNode{};
        }
        else {
            cellNode = rowNode.first_child_of_type(pugi::node_element);

            uint16_t currentCol = extractColumnFromCellRef(cellNode.attribute("r").value());
            while (currentCol < columnNumber) {
                cellNode   = cellNode.next_sibling_of_type(pugi::node_element);
                currentCol = extractColumnFromCellRef(cellNode.attribute("r").value());
            }
            if (currentCol > columnNumber) return XMLNode{};
        }
        return cellNode;
    }
}    // namespace OpenXLSX

/**
 * @details
 */
XLCellIterator::XLCellIterator(const XLCellRange& cellRange, XLIteratorLocation loc, std::vector<XLStyleIndex> const* colStyles)
    : m_dataNode(cellRange.m_dataNode),
      m_topLeft(cellRange.m_topLeft),
      m_bottomRight(cellRange.m_bottomRight),
      m_sharedStrings(cellRange.m_sharedStrings),
      m_endReached(false),
      m_hintNode(),
      m_hintRow(0),
      m_currentCell(),
      m_currentCellStatus(CellStatus::NotLoaded),
      m_currentRow(0),
      m_currentColumn(0),
      m_colStyles(colStyles)
{
    if (loc == XLIteratorLocation::End)
        m_endReached = true;
    else {
        m_currentRow    = m_topLeft.row();
        m_currentColumn = m_topLeft.column();
    }
    if (m_colStyles == nullptr) throw XLInternalError("XLCellIterator constructor parameter colStyles must not be nullptr");
}

/**
 * @details
 */
void XLCellIterator::updateCurrentCell(bool createIfMissing) const
{
    if (m_currentCellStatus == CellStatus::Loaded) return;
    if (!createIfMissing and m_currentCellStatus == CellStatus::NoSuchCell) return;

    if (m_endReached) throw XLInputError("XLCellIterator updateCurrentCell: iterator should not be dereferenced when endReached() == true");

    if (m_hintNode.empty()) {
        // Fallback for first lookup.
        if (createIfMissing)
            m_currentCell =
                XLCell(getCellNode(getRowNode(m_dataNode, m_currentRow), m_currentColumn, 0, *m_colStyles), m_sharedStrings.get());
        else
            m_currentCell = XLCell(findCellNode(findRowNode(m_dataNode, m_currentRow), m_currentColumn), m_sharedStrings.get());
    }
    else {
        if (m_currentRow == m_hintRow) {
            XMLNode  cellNode = m_hintNode.next_sibling_of_type(pugi::node_element);
            uint16_t colNo    = 0;
            while (not cellNode.empty()) {
                colNo = extractColumnFromCellRef(cellNode.attribute("r").value());
                if (colNo >= m_currentColumn) break;
                cellNode = cellNode.next_sibling_of_type(pugi::node_element);
            }
            if (colNo != m_currentColumn) cellNode = XMLNode{};

            if (createIfMissing and cellNode.empty()) {
                cellNode = m_hintNode.parent().insert_child_after("c", m_hintNode);
                char cellAddrBuf[16];
                makeCellAddress(m_currentRow, m_currentColumn, cellAddrBuf);
                setDefaultCellAttributes(cellNode, cellAddrBuf, m_hintNode.parent(), m_currentColumn, *m_colStyles);
            }
            m_currentCell = XLCell(cellNode, m_sharedStrings.get());
        }
        else if (m_currentRow > m_hintRow) {
            XMLNode  rowNode = m_hintNode.parent().next_sibling_of_type(pugi::node_element);
            uint32_t rowNo   = 0;
            while (not rowNode.empty()) {
                rowNo = static_cast<uint32_t>(rowNode.attribute("r").as_ullong());
                if (rowNo >= m_currentRow) break;
                rowNode = rowNode.next_sibling_of_type(pugi::node_element);
            }
            if (rowNo != m_currentRow) rowNode = XMLNode{};

            if (createIfMissing and rowNode.empty()) {
                rowNode = m_dataNode.insert_child_after("row", m_hintNode.parent());
                rowNode.append_attribute("r").set_value(m_currentRow);
            }

            if (rowNode.empty())
                m_currentCell = XLCell{};
            else {
                if (createIfMissing)
                    m_currentCell = XLCell(getCellNode(rowNode, m_currentColumn, m_currentRow, *m_colStyles), m_sharedStrings.get());
                else
                    m_currentCell = XLCell(findCellNode(rowNode, m_currentColumn), m_sharedStrings.get());
            }
        }
        else
            throw XLInternalError("XLCellIterator::updateCurrentCell: an internal error occured (m_currentRow < m_hintRow)");
    }

    if (m_currentCell.empty())
        m_currentCellStatus = CellStatus::NoSuchCell;
    else {
        // Cache the result to optimize subsequent lookups.
        m_hintNode          = m_currentCell.m_cellNode;
        m_hintRow           = m_currentRow;
        m_currentCellStatus = CellStatus::Loaded;
    }
}

/**
 * @details
 */
XLCellIterator& XLCellIterator::operator++()
{
    if (m_endReached) throw XLInputError("XLCellIterator: tried to increment beyond end operator");

    if (m_currentColumn < m_bottomRight.column())
        ++m_currentColumn;
    else if (m_currentRow < m_bottomRight.row()) {
        ++m_currentRow;
        m_currentColumn = m_topLeft.column();
    }
    else
        m_endReached = true;

    m_currentCellStatus = CellStatus::NotLoaded;

    return *this;
}

/**
 * @details
 */
XLCellIterator XLCellIterator::operator++(int)    // NOLINT
{
    auto oldIter(*this);
    ++(*this);
    return oldIter;
}

/**
 * @details
 */
XLCell& XLCellIterator::operator*()
{
    updateCurrentCell(XLCreateIfMissing);
    return m_currentCell;
}

/**
 * @details
 */
XLCellIterator::pointer XLCellIterator::operator->()
{
    updateCurrentCell(XLCreateIfMissing);
    return &m_currentCell;
}

/**
 * @details
 */
bool XLCellIterator::operator==(const XLCellIterator& rhs) const noexcept
{
    if (m_endReached and rhs.m_endReached) return true;

    if ((m_currentColumn != rhs.m_currentColumn) or (m_currentRow != rhs.m_currentRow)) return false;

    // Check if iterators belong to the same underlying XML document node (same worksheet)
    if (m_dataNode != rhs.m_dataNode) return false;

    // Check if iterators were created with the same range boundaries
    if (m_topLeft.address() != rhs.m_topLeft.address() || m_bottomRight.address() != rhs.m_bottomRight.address()) return false;

    return true;
}

/**
 * @details
 */
bool XLCellIterator::operator!=(const XLCellIterator& rhs) const noexcept { return !(*this == rhs); }

/**
 * @details
 */
bool XLCellIterator::cellExists() const
{
    updateCurrentCell(XLDoNotCreateIfMissing);
    return not m_currentCell.empty();
}

/**
 * @details
 */
uint64_t XLCellIterator::distance(const XLCellIterator& last) const
{
    uint32_t row     = (m_endReached ? m_bottomRight.row() : m_currentRow);
    uint16_t col     = (m_endReached ? m_bottomRight.column() + 1 : m_currentColumn);
    uint32_t lastRow = (last.m_endReached ? last.m_bottomRight.row() : last.m_currentRow);
    uint16_t lastCol = (last.m_endReached ? last.m_bottomRight.column() + 1 : last.m_currentColumn);

    uint16_t rowWidth = m_bottomRight.column() - m_topLeft.column() + 1;
    int64_t  distance = (static_cast<int64_t>(lastRow) - row) * rowWidth + static_cast<int64_t>(lastCol) - col;
    if (distance < 0) throw XLInputError("XLCellIterator::distance is negative");

    return static_cast<uint64_t>(distance);
}

/**
 * @details
 */
std::string XLCellIterator::address() const
{
    uint32_t row = (m_endReached ? m_bottomRight.row() : m_currentRow);
    uint16_t col = (m_endReached ? m_bottomRight.column() + 1 : m_currentColumn);
    char     cellAddrBuf[16];
    makeCellAddress(row, col, cellAddrBuf);
    return (m_endReached ? "END(" : "") + std::string(cellAddrBuf) + (m_endReached ? ")" : "");
}

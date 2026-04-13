#include "XLRowData.hpp"
#include "XLCell.hpp"
#include "XLRow.hpp"
#include "XLUtilities.hpp"
#include <algorithm>
#include <cassert>
namespace OpenXLSX
{
    XLRowDataIterator::XLRowDataIterator(const XLRowDataRange& rowDataRange, XLIteratorLocation loc)
        : m_dataRange(std::make_unique<XLRowDataRange>(rowDataRange)),
          m_cellNode(getCellNode((m_dataRange->size() ? *m_dataRange->m_rowNode : XMLNode{}), m_dataRange->m_firstCol)),
          m_currentCell(loc == XLIteratorLocation::End ? XLCell() : XLCell(m_cellNode, m_dataRange->m_sharedStrings.get()))
    {}
    XLRowDataIterator::~XLRowDataIterator() = default;
    XLRowDataIterator::XLRowDataIterator(const XLRowDataIterator& other)
        : m_dataRange(std::make_unique<XLRowDataRange>(*other.m_dataRange)),
          m_cellNode(other.m_cellNode),
          m_currentCell(other.m_currentCell)
    {}
    XLRowDataIterator::XLRowDataIterator(XLRowDataIterator&& other) noexcept = default;
    XLRowDataIterator& XLRowDataIterator::operator=(const XLRowDataIterator& other)
    {
        if (&other != this) {
            XLRowDataIterator temp = other;
            std::swap(temp, *this);
        }
        return *this;
    }
    XLRowDataIterator& XLRowDataIterator::operator=(XLRowDataIterator&& other) noexcept = default;
    XLRowDataIterator& XLRowDataIterator::operator++()
    {
        if (not m_currentCell)    // 2025-07-14 BUGFIX issue #368: check that m_currentCell is valid
            throw XLInputError("XLRowDataIterator: tried to increment beyond end operator");
        const uint16_t cellNumber = m_currentCell.cellReference().column() + 1;
        XMLNode        cellNode   = m_currentCell.m_cellNode.next_sibling_of_type(pugi::node_element);
        if (cellNumber > m_dataRange->m_lastCol)
            m_currentCell = XLCell();

        else if (cellNode.empty() or extractColumnFromCellRef(cellNode.attribute("r").value()) > cellNumber) {
            cellNode = m_dataRange->m_rowNode->insert_child_after("c", m_currentCell.m_cellNode);
            char cellAddrBuf[16];
            makeCellAddress(static_cast<uint32_t>(m_dataRange->m_rowNode->attribute("r").as_ullong()), cellNumber, cellAddrBuf);
            setDefaultCellAttributes(cellNode, cellAddrBuf, *m_dataRange->m_rowNode, cellNumber);
            m_currentCell = XLCell(cellNode, m_dataRange->m_sharedStrings.get());
        }
        else {
            assert(extractColumnFromCellRef(cellNode.attribute("r").value()) == cellNumber);
            m_currentCell = XLCell(cellNode, m_dataRange->m_sharedStrings.get());
        }
        return *this;
    }
    XLRowDataIterator XLRowDataIterator::operator++(int)
    {
        auto oldIter(*this);
        ++(*this);
        return oldIter;
    }
    XLCell&                    XLRowDataIterator::operator*() { return m_currentCell; }
    XLRowDataIterator::pointer XLRowDataIterator::operator->() { return &m_currentCell; }
    bool                       XLRowDataIterator::operator==(const XLRowDataIterator& rhs) const
    {
        if (static_cast<bool>(m_currentCell) != static_cast<bool>(rhs.m_currentCell)) return false;
        if (not m_currentCell) return true;
        return m_currentCell == rhs.m_currentCell;
    }
    bool XLRowDataIterator::operator!=(const XLRowDataIterator& rhs) const { return !(*this == rhs); }
}    // namespace OpenXLSX
namespace OpenXLSX
{
    XLRowDataRange::XLRowDataRange(const XMLNode& rowNode, uint16_t firstColumn, uint16_t lastColumn, const XLSharedStrings& sharedStrings)
        : m_rowNode(std::make_unique<XMLNode>(rowNode)),
          m_firstCol(firstColumn),
          m_lastCol(lastColumn),
          m_sharedStrings(sharedStrings)
    {
        if (lastColumn < firstColumn) {
            m_firstCol = 1;
            m_lastCol  = 1;
            throw XLOverflowError("lastColumn is less than firstColumn.");
        }
    }
    XLRowDataRange::XLRowDataRange()
        : m_rowNode(nullptr),
          m_firstCol(1),    // first col of 1
          m_lastCol(0),     // and last col of 0 will ensure that size returns 0
          m_sharedStrings(XLSharedStringsDefaulted)
    {}
    XLRowDataRange::XLRowDataRange(const XLRowDataRange& other)
        : m_rowNode((other.m_rowNode != nullptr) ? std::make_unique<XMLNode>(*other.m_rowNode) : nullptr),
          m_firstCol(other.m_firstCol),
          m_lastCol(other.m_lastCol),
          m_sharedStrings(other.m_sharedStrings)
    {}
    XLRowDataRange::XLRowDataRange(XLRowDataRange&& other) noexcept = default;
    XLRowDataRange::~XLRowDataRange()                               = default;
    XLRowDataRange& XLRowDataRange::operator=(const XLRowDataRange& other)
    {
        if (&other != this) {
            XLRowDataRange temp(other);
            std::swap(temp, *this);
        }
        return *this;
    }
    XLRowDataRange&   XLRowDataRange::operator=(XLRowDataRange&& other) noexcept = default;
    uint16_t          XLRowDataRange::size() const { return m_lastCol - m_firstCol + 1; }
    XLRowDataIterator XLRowDataRange::begin()
    { return XLRowDataIterator{*this, (size() > 0 ? XLIteratorLocation::Begin : XLIteratorLocation::End)}; }
    XLRowDataIterator XLRowDataRange::end() { return XLRowDataIterator{*this, XLIteratorLocation::End}; }
}    // namespace OpenXLSX
namespace OpenXLSX
{
    XLRowDataProxy::~XLRowDataProxy() = default;
    XLRowDataProxy& XLRowDataProxy::operator=(const XLRowDataProxy& other)
    {
        if (&other != this) { *this = other.getValues(); }
        return *this;
    }
    XLRowDataProxy::XLRowDataProxy(XLRow* row, XMLNode* rowNode) : m_row(row), m_rowNode(rowNode) {}
    XLRowDataProxy::XLRowDataProxy(const XLRowDataProxy& other)                = default;
    XLRowDataProxy::XLRowDataProxy(XLRowDataProxy&& other) noexcept            = default;
    XLRowDataProxy& XLRowDataProxy::operator=(XLRowDataProxy&& other) noexcept = default;
    XLRowDataProxy& XLRowDataProxy::operator=(const std::vector<XLCellValue>& values)
    {
        if (values.size() > MAX_COLS) throw XLOverflowError("vector<XLCellValue> size exceeds maximum number of columns.");
        if (values.empty()) return *this;
        deleteCellValues(static_cast<uint16_t>(values.size()));
        XMLNode        curNode{};
        uint16_t       colNo  = static_cast<uint16_t>(values.size());
        const uint32_t rowNum = m_row->rowNumber();
        char           addrBuffer[16];    // Buffer for cell address (e.g., "XFD1048576")
        for (auto value = values.rbegin(); value != values.rend(); ++value) {
            curNode = m_rowNode->prepend_child("c");
            makeCellAddress(rowNum, colNo, addrBuffer);
            setDefaultCellAttributes(curNode, addrBuffer, *m_rowNode, colNo);
            XLCell(curNode, m_row->m_sharedStrings.get()).value() = *value;
            --colNo;
        }
        return *this;
    }
    XLRowDataProxy& XLRowDataProxy::operator=(const std::vector<bool>& values)
    {
        if (values.size() > MAX_COLS) throw XLOverflowError("vector<bool> size exceeds maximum number of columns.");
        if (values.empty()) return *this;
        auto range = XLRowDataRange(*m_rowNode, 1, static_cast<uint16_t>(values.size()), m_row->m_sharedStrings.get());
        auto dst   = range.begin();

        auto src = values.begin();
        while (true) {
            dst->value() = static_cast<bool>(*src);
            ++src;
            if (src == values.end()) break;
            ++dst;
        }
        return *this;
    }
    XLRowDataProxy::         operator std::vector<XLCellValue>() const { return getValues(); }
    XLRowDataProxy::         operator std::deque<XLCellValue>() const { return convertContainer<std::deque<XLCellValue>>(); }
    XLRowDataProxy::         operator std::list<XLCellValue>() const { return convertContainer<std::list<XLCellValue>>(); }
    std::vector<XLCellValue> XLRowDataProxy::getValues() const
    {
        const XMLNode  lastElementChild = m_rowNode->last_child_of_type(pugi::node_element);
        const uint16_t numCells = (lastElementChild.empty() ? 0 : extractColumnFromCellRef(lastElementChild.attribute("r").value()));
        std::vector<XLCellValue> result(static_cast<uint64_t>(numCells));
        if (numCells > 0) {
            XMLNode node = lastElementChild;    // avoid unneeded call to first_child_of_type by iterating backwards, vector is random

            while (not node.empty()) {
                result[extractColumnFromCellRef(node.attribute("r").value()) - 1] = XLCell(node, m_row->m_sharedStrings.get()).value();
                node                                                              = node.previous_sibling_of_type(pugi::node_element);
            }
        }
        return result;
    }
    const XLSharedStrings& XLRowDataProxy::getSharedStrings() const { return m_row->m_sharedStrings.get(); }
    void                   XLRowDataProxy::deleteCellValues(uint16_t count)
    {
        XMLNode cellNode = m_rowNode->first_child_of_type(pugi::node_element);
        while (not cellNode.empty()) {
            if (extractColumnFromCellRef(cellNode.attribute("r").value()) <= count) {
                XMLNode nextNode    = cellNode.next_sibling();
                XMLNode nextElement = cellNode.next_sibling_of_type(pugi::node_element);

                m_rowNode->remove_child(cellNode);

                while (not nextNode.empty() and nextNode != nextElement) {
                    XMLNode toDelete = nextNode;
                    nextNode         = nextNode.next_sibling();
                    m_rowNode->remove_child(toDelete);
                }

                cellNode = nextElement;
            }
            else {
                cellNode = cellNode.next_sibling_of_type(pugi::node_element);
            }
        }
    }
    void XLRowDataProxy::prependCellValue(const XLCellValue& value, uint16_t col)
    {
        XMLNode curNode = m_rowNode->prepend_child("c");
        char    addrBuffer[16];
        makeCellAddress(m_row->rowNumber(), col, addrBuffer);
        setDefaultCellAttributes(curNode, addrBuffer, *m_rowNode, col);
        XLCell(curNode, m_row->m_sharedStrings.get()).value() = value;
    }
    void XLRowDataProxy::clear() { m_rowNode->remove_children(); }
}    // namespace OpenXLSX

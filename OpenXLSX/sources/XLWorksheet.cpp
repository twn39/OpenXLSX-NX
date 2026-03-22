#include <algorithm>
#include <fmt/format.h>
#include <pugixml.hpp>
#include "XLWorksheet.hpp"
#include "XLUtilities.hpp"
#include "XLCellRange.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

/**
 * @details The constructor does some slight reconfiguration of the XML file, in order to make parsing easier.
 */
XLWorksheet::XLWorksheet(XLXmlData* xmlData) : XLSheetBase(xmlData)
{
    if (xmlDocument().document_element().child("cols").type() != pugi::node_null) {
        auto currentNode = xmlDocument().document_element().child("cols").first_child_of_type(pugi::node_element);
        while (not currentNode.empty()) {
            uint16_t min{};
            uint16_t max{};
            try {
                min = static_cast<uint16_t>(std::stoi(currentNode.attribute("min").value()));
                max = static_cast<uint16_t>(std::stoi(currentNode.attribute("max").value()));
            }
            catch (...) {
                throw XLInternalError("Worksheet column min and/or max attributes are invalid.");
            }
            if (min != max) {
                currentNode.attribute("min").set_value(max);
                for (uint16_t i = min; i < max; i++) {
                    auto newnode = xmlDocument().document_element().child("cols").insert_child_before("col", currentNode);
                    auto attr    = currentNode.first_attribute();
                    while (not attr.empty()) {
                        newnode.append_attribute(attr.name()) = attr.value();
                        attr                                  = attr.next_attribute();
                    }
                    newnode.attribute("min") = i;
                    newnode.attribute("max") = i;
                }
            }
            currentNode = currentNode.next_sibling_of_type(pugi::node_element);
        }
    }
}

XLWorksheet::XLWorksheet(const XLWorksheet& other) : XLSheetBase<XLWorksheet>(other)
{
    m_relationships = other.m_relationships;
    m_merges        = other.m_merges;
    m_dataValidations = other.m_dataValidations;
    m_drawing       = other.m_drawing;
    m_vmlDrawing    = other.m_vmlDrawing;
    m_comments      = other.m_comments;
    m_tables        = other.m_tables;
}

XLWorksheet::XLWorksheet(XLWorksheet&& other) : XLSheetBase<XLWorksheet>(std::move(other))
{
    m_relationships = std::move(other.m_relationships);
    m_merges        = std::move(other.m_merges);
    m_dataValidations = std::move(other.m_dataValidations);
    m_drawing       = std::move(other.m_drawing);
    m_vmlDrawing    = std::move(other.m_vmlDrawing);
    m_comments      = std::move(other.m_comments);
    m_tables        = std::move(other.m_tables);
}

XLWorksheet& XLWorksheet::operator=(const XLWorksheet& other)
{
    if (&other != this) {
        XLSheetBase<XLWorksheet>::operator=(other);
        m_relationships = other.m_relationships;
        m_merges        = other.m_merges;
        m_dataValidations = other.m_dataValidations;
        m_drawing       = other.m_drawing;
        m_vmlDrawing    = other.m_vmlDrawing;
        m_comments      = other.m_comments;
        m_tables        = other.m_tables;
    }
    return *this;
}

XLWorksheet& XLWorksheet::operator=(XLWorksheet&& other)
{
    if (&other != this) {
        XLSheetBase<XLWorksheet>::operator=(std::move(other));
        m_relationships = std::move(other.m_relationships);
        m_merges        = std::move(other.m_merges);
        m_dataValidations = std::move(other.m_dataValidations);
        m_drawing       = std::move(other.m_drawing);
        m_vmlDrawing    = std::move(other.m_vmlDrawing);
        m_comments      = std::move(other.m_comments);
        m_tables        = std::move(other.m_tables);
    }
    return *this;
}

XLColor XLWorksheet::getColor_impl() const
{
    auto node = xmlDocument().document_element().child("sheetPr").child("tabColor");
    if (node.empty()) return XLColor();
    return XLColor(node.attribute("rgb").value());
}

void XLWorksheet::setColor_impl(const XLColor& color) { setTabColor(xmlDocument(), color); }

bool XLWorksheet::isSelected_impl() const { return tabIsSelected(xmlDocument()); }

void XLWorksheet::setSelected_impl(bool selected) { setTabSelected(xmlDocument(), selected); }

bool XLWorksheet::isActive_impl() const
{ return parentDoc().execQuery(XLQuery(XLQueryType::QuerySheetIsActive).setParam("sheetID", relationshipID())).result<bool>(); }

bool XLWorksheet::setActive_impl()
{ return parentDoc().execCommand(XLCommand(XLCommandType::SetSheetActive).setParam("sheetID", relationshipID())); }

XLCellAssignable XLWorksheet::cell(const std::string& ref) const { return cell(XLCellReference(ref)); }
XLCellAssignable XLWorksheet::cell(const XLCellReference& ref) const { return cell(ref.row(), ref.column()); }

XLCellAssignable XLWorksheet::cell(uint32_t rowNumber, uint16_t columnNumber) const
{
    const XMLNode rowNode  = getRowNode(xmlDocument().document_element().child("sheetData"), rowNumber);
    const XMLNode cellNode = getCellNode(rowNode, columnNumber, rowNumber);
    return XLCellAssignable(XLCell(cellNode, parentDoc().sharedStrings()));
}

XLCellAssignable XLWorksheet::findCell(const std::string& ref) const { return findCell(XLCellReference(ref)); }
XLCellAssignable XLWorksheet::findCell(const XLCellReference& ref) const { return findCell(ref.row(), ref.column()); }

XLCellAssignable XLWorksheet::findCell(uint32_t rowNumber, uint16_t columnNumber) const
{
    return XLCellAssignable(XLCell(findCellNode(findRowNode(xmlDocument().document_element().child("sheetData"), rowNumber), columnNumber),
                                   parentDoc().sharedStrings()));
}

XLCellRange XLWorksheet::range() const { return range(XLCellReference("A1"), lastCell()); }
XLCellRange XLWorksheet::range(const XLCellReference& topLeft, const XLCellReference& bottomRight) const
{ return XLCellRange(xmlDocument().document_element().child("sheetData"), topLeft, bottomRight, parentDoc().sharedStrings()); }

XLCellRange XLWorksheet::range(std::string const& topLeft, std::string const& bottomRight) const
{ return range(XLCellReference(topLeft), XLCellReference(bottomRight)); }

XLCellRange XLWorksheet::range(std::string const& rangeReference) const
{
    size_t pos = rangeReference.find_first_of(':');
    if (pos == std::string::npos) {
        return range(rangeReference, rangeReference);
    }
    std::string topLeft = rangeReference.substr(0, pos);
    std::string bottomRight = rangeReference.substr(pos + 1);

    bool topIsCol = std::all_of(topLeft.begin(), topLeft.end(), [](unsigned char c){ return std::isalpha(c); });
    bool bottomIsCol = std::all_of(bottomRight.begin(), bottomRight.end(), [](unsigned char c){ return std::isalpha(c); });
    if (topIsCol && bottomIsCol) {
        return range(topLeft + "1", bottomRight + std::to_string(OpenXLSX::MAX_ROWS));
    }

    bool topIsRow = std::all_of(topLeft.begin(), topLeft.end(), [](unsigned char c){ return std::isdigit(c); });
    bool bottomIsRow = std::all_of(bottomRight.begin(), bottomRight.end(), [](unsigned char c){ return std::isdigit(c); });
    if (topIsRow && bottomIsRow) {
        return range("A" + topLeft, XLCellReference::columnAsString(OpenXLSX::MAX_COLS) + bottomRight);
    }

    return range(topLeft, bottomRight);
}


XLRowRange XLWorksheet::rows() const
{
    const auto sheetDataNode = xmlDocument().document_element().child("sheetData");
    return XLRowRange(sheetDataNode,
                      1,
                      (sheetDataNode.last_child_of_type(pugi::node_element).empty()
                           ? 1
                           : static_cast<uint32_t>(sheetDataNode.last_child_of_type(pugi::node_element).attribute("r").as_ullong())),
                      parentDoc().sharedStrings());
}

XLRowRange XLWorksheet::rows(uint32_t rowCount) const
{ return XLRowRange(xmlDocument().document_element().child("sheetData"), 1, rowCount, parentDoc().sharedStrings()); }

XLRowRange XLWorksheet::rows(uint32_t firstRow, uint32_t lastRow) const
{ return XLRowRange(xmlDocument().document_element().child("sheetData"), firstRow, lastRow, parentDoc().sharedStrings()); }

void XLWorksheet::appendRow(const std::vector<XLCellValue>& values)
{
    row(rowCount() + 1).values() = values;
}


XLRow XLWorksheet::row(uint32_t rowNumber) const
{ return XLRow{getRowNode(xmlDocument().document_element().child("sheetData"), rowNumber), parentDoc().sharedStrings()}; }

XLColumn XLWorksheet::column(uint16_t columnNumber) const
{
    using namespace std::literals::string_literals;
    if (columnNumber < 1 or columnNumber > OpenXLSX::MAX_COLS)
        throw XLException("XLWorksheet::column: columnNumber "s + std::to_string(columnNumber) + " is outside allowed range [1;"s + std::to_string(MAX_COLS) + "]"s);

    if (xmlDocument().document_element().child("cols").empty())
        xmlDocument().document_element().insert_child_before("cols", xmlDocument().document_element().child("sheetData"));

    auto columnNode = xmlDocument().document_element().child("cols").find_child([&](const XMLNode node) {
        return (columnNumber >= node.attribute("min").as_int() and columnNumber <= node.attribute("max").as_int()) ||
               node.attribute("min").as_int() > columnNumber;
    });

    uint16_t minColumn{};
    uint16_t maxColumn{};
    if (not columnNode.empty()) {
        minColumn = static_cast<uint16_t>(columnNode.attribute("min").as_int());
        maxColumn = static_cast<uint16_t>(columnNode.attribute("max").as_int());
    }

    if (not columnNode.empty() and (minColumn == columnNumber) and (maxColumn == columnNumber)) {}
    else if (not columnNode.empty() and (columnNumber >= minColumn) and (minColumn != maxColumn)) {
        columnNode.attribute("min").set_value(maxColumn);
        for (int i = minColumn; i < maxColumn; ++i) {
            auto node = xmlDocument().document_element().child("cols").insert_copy_before(columnNode, columnNode);
            node.attribute("min").set_value(i);
            node.attribute("max").set_value(i);
        }
        while (not columnNode.empty() and columnNode.attribute("min").as_int() != columnNumber)
            columnNode = columnNode.previous_sibling_of_type(pugi::node_element);
        if (columnNode.empty())
            throw XLInternalError("XLWorksheet::"s + __func__ + ": column node for index "s + std::to_string(columnNumber) + "not found after splitting column nodes"s);
    }
    else if (not columnNode.empty() and minColumn > columnNumber) {
        columnNode                                 = xmlDocument().document_element().child("cols").insert_child_before("col", columnNode);
        columnNode.append_attribute("min")         = columnNumber;
        columnNode.append_attribute("max")         = columnNumber;
        columnNode.append_attribute("width")       = 9.8;
        columnNode.append_attribute("customWidth") = 0;
    }
    else if (columnNode.empty()) {
        columnNode                                 = xmlDocument().document_element().child("cols").append_child("col");
        columnNode.append_attribute("min")         = columnNumber;
        columnNode.append_attribute("max")         = columnNumber;
        columnNode.append_attribute("width")       = 9.8;
        columnNode.append_attribute("customWidth") = 0;
    }

    if (columnNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLInternalError("XLWorksheet::"s + __func__ + ": was unable to find or create node for column "s + std::to_string(columnNumber));
    }
    return XLColumn(columnNode);
}

XLColumn XLWorksheet::column(std::string const& columnRef) const { return column(XLCellReference::columnAsNumber(columnRef)); }

void XLWorksheet::setAutoFilter(const XLCellRange& range)
{
    XMLNode docElement     = xmlDocument().document_element();
    XMLNode autoFilterNode = docElement.child("autoFilter");
    if (autoFilterNode.empty()) { autoFilterNode = appendAndGetNode(docElement, "autoFilter", m_nodeOrder); }
    appendAndSetAttribute(autoFilterNode, "ref", range.address());
}

void XLWorksheet::clearAutoFilter() { xmlDocument().document_element().remove_child("autoFilter"); }
bool XLWorksheet::hasAutoFilter() const { return !xmlDocument().document_element().child("autoFilter").empty(); }

std::string XLWorksheet::autoFilter() const
{
    XMLNode autoFilterNode = xmlDocument().document_element().child("autoFilter");
    if (autoFilterNode.empty()) return "";
    return autoFilterNode.attribute("ref").value();
}

XLAutoFilter XLWorksheet::autofilterObject() const
{
    XMLNode autoFilterNode = xmlDocument().document_element().child("autoFilter");
    return XLAutoFilter(autoFilterNode);
}

XLCellReference XLWorksheet::lastCell() const noexcept { return {rowCount(), columnCount()}; }

uint16_t XLWorksheet::columnCount() const noexcept
{
    uint16_t   maxCount  = 0;
    XLRowRange rowsRange = rows();
    for (XLRowIterator rowIt = rowsRange.begin(); rowIt != rowsRange.end(); ++rowIt) {
        if (rowIt.rowExists()) {
            uint16_t cellCount = rowIt->cellCount();
            maxCount           = std::max(cellCount, maxCount);
        }
    }
    return maxCount;
}

uint32_t XLWorksheet::rowCount() const noexcept
{
    auto node = xmlDocument().document_element().child("sheetData").last_child_of_type(pugi::node_element);
    if (node.empty()) return 0;
    return static_cast<uint32_t>(node.attribute("r").as_ullong());
}

bool XLWorksheet::deleteRow(uint32_t rowNumber)
{
    XMLNode row     = xmlDocument().document_element().child("sheetData").first_child_of_type(pugi::node_element);
    XMLNode lastRow = xmlDocument().document_element().child("sheetData").last_child_of_type(pugi::node_element);
    if (row.empty() or rowNumber < row.attribute("r").as_ullong() or rowNumber > lastRow.attribute("r").as_ullong()) return false;
    if (rowNumber - row.attribute("r").as_ullong() < lastRow.attribute("r").as_ullong() - rowNumber)
        while (not row.empty() and (row.attribute("r").as_ullong() < rowNumber)) row = row.next_sibling_of_type(pugi::node_element);
    else {
        row = lastRow;
        while (not row.empty() and (row.attribute("r").as_ullong() > rowNumber)) row = row.previous_sibling_of_type(pugi::node_element);
    }
    if (row.empty() or row.attribute("r").as_ullong() != rowNumber) return false;
    return xmlDocument().document_element().child("sheetData").remove_child(row);
}

void XLWorksheet::updateSheetName(const std::string& oldName, const std::string& newName)
{
    std::string oldNameTemp = oldName;
    std::string newNameTemp = newName;
    std::string formula;
    if (oldName.find(' ') != std::string::npos) oldNameTemp = "\'" + oldName + "\'";
    if (newName.find(' ') != std::string::npos) newNameTemp = "\'" + newName + "\'";
    oldNameTemp += '!';
    newNameTemp += '!';
    XMLNode row = xmlDocument().document_element().child("sheetData").first_child_of_type(pugi::node_element);
    for (; not row.empty(); row = row.next_sibling_of_type(pugi::node_element)) {
        for (XMLNode cell = row.first_child_of_type(pugi::node_element); not cell.empty();
             cell         = cell.next_sibling_of_type(pugi::node_element))
        {
            if (!XLCell(cell, parentDoc().sharedStrings()).hasFormula()) continue;
            formula = XLCell(cell, parentDoc().sharedStrings()).formula().get();
            if (formula.find('[') == std::string::npos and formula.find(']') == std::string::npos) {
                while (formula.find(oldNameTemp) != std::string::npos) {
                    formula.replace(formula.find(oldNameTemp), oldNameTemp.length(), newNameTemp);
                }
                XLCell(cell, parentDoc().sharedStrings()).formula() = formula;
            }
        }
    }
}

void XLWorksheet::updateDimension()
{
    uint32_t rows = rowCount();
    uint16_t cols = columnCount();
    auto rootNode = xmlDocument().document_element();
    auto dimensionNode = rootNode.child("dimension");
    if (dimensionNode.empty()) dimensionNode = rootNode.prepend_child("dimension");
    dimensionNode.remove_children();
    if (rows == 0 || cols == 0) {
        dimensionNode.attribute("ref").set_value("A1");
    } else {
        std::string ref = "A1:" + XLCellReference::columnAsString(cols) + std::to_string(rows);
        auto attr = dimensionNode.attribute("ref");
        if (attr.empty()) attr = dimensionNode.append_attribute("ref");
        attr.set_value(ref.c_str());
    }
}



std::string XLWorksheet::makeInternalLocation(std::string_view sheetName, std::string_view cellRef)
{
    std::string result;
    if (sheetName.find(' ') != std::string_view::npos) {
        result.reserve(sheetName.size() + cellRef.size() + 3);
        result += "'";
        result += sheetName;
        result += "'!";
        result += cellRef;
    }
    else {
        result.reserve(sheetName.size() + cellRef.size() + 1);
        result += sheetName;
        result += "!";
        result += cellRef;
    }
    return result;
}

uint16_t XLWorksheet::sheetXmlNumber() const
{
    constexpr const char* searchPattern = "xl/worksheets/sheet";
    std::string           xmlPath       = getXmlPath();
    size_t                pos           = xmlPath.find(searchPattern);
    if (pos == std::string::npos) return 0;
    pos += strlen(searchPattern);
    size_t pos2 = pos;
    while (std::isdigit(xmlPath[pos2])) ++pos2;
    if (pos2 == pos or xmlPath.substr(pos2) != ".xml") return 0;
    return static_cast<uint16_t>(std::stoi(xmlPath.substr(pos, pos2 - pos)));
}

std::optional<XLCell> XLWorksheet::peekCell(const std::string& ref) const { return peekCell(XLCellReference(ref)); }
std::optional<XLCell> XLWorksheet::peekCell(const XLCellReference& ref) const { return peekCell(ref.row(), ref.column()); }

std::optional<XLCell> XLWorksheet::peekCell(uint32_t rowNumber, uint16_t columnNumber) const
{
    XMLNode cellNode = findCellNode(findRowNode(xmlDocument().document_element().child("sheetData"), rowNumber), columnNumber);
    if (!cellNode) return std::nullopt;
    return XLCell(cellNode, parentDoc().sharedStrings());
}


XLStreamWriter XLWorksheet::streamWriter()
{
    if (m_xmlData->m_isStreamed && !m_xmlData->m_streamFilePath.empty()) {
        std::error_code ec;
        if (std::filesystem::exists(m_xmlData->m_streamFilePath, ec)) {
            std::filesystem::remove(m_xmlData->m_streamFilePath, ec);
        }
        m_xmlData->m_streamFilePath.clear();
        m_xmlData->m_isStreamed = false;
    }

    XMLDocument& doc = xmlDocument();
    XMLNode root = doc.document_element();
    XMLNode sheetData = root.child("sheetData");
    
    if (sheetData.empty()) {
        sheetData = appendAndGetNode(root, "sheetData", m_nodeOrder);
    }
    
    while (sheetData.first_child()) {
        sheetData.remove_child(sheetData.first_child());
    }

    struct StringWriter : pugi::xml_writer {
        std::string result;
        void write(const void* data, size_t size) override {
            result.append(static_cast<const char*>(data), size);
        }
    } sw;
    
    doc.save(sw, "", pugi::format_raw | pugi::format_no_declaration);
    std::string xmlStr = sw.result;
    
    std::string topHalf = "";
    std::string bottomHalf = "";
    
    // Split accurately
    size_t pos = xmlStr.find("<sheetData/>");
    if (pos != std::string::npos) {
        topHalf = xmlStr.substr(0, pos) + "<sheetData>";
        bottomHalf = "</sheetData>" + xmlStr.substr(pos + 12);
    } else {
        pos = xmlStr.find("<sheetData></sheetData>");
        if (pos != std::string::npos) {
            topHalf = xmlStr.substr(0, pos) + "<sheetData>";
            bottomHalf = "</sheetData>" + xmlStr.substr(pos + 23);
        } else {
            size_t endTag = xmlStr.find("</worksheet>");
            topHalf = xmlStr.substr(0, endTag) + "<sheetData>";
            bottomHalf = "</sheetData></worksheet>";
        }
    }
    
    XLStreamWriter writer(this);
    writer.m_bottomHalf = bottomHalf;
    
    if (writer.m_stream.is_open()) {
        std::string header = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
        writer.m_stream << header << topHalf;
    }
    
    m_xmlData->m_isStreamed = true;
    m_xmlData->m_streamFilePath = writer.getTempFilePath();
    
    return writer;
}


#include "XLWorksheet.hpp"
#include "XLCellRange.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLUtilities.hpp"
#include <algorithm>
#include <charconv>
#include <fmt/format.h>
#include <pugixml.hpp>
#include <sstream>

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

            std::string_view minStr = currentNode.attribute("min").value();
            std::string_view maxStr = currentNode.attribute("max").value();

            auto resMin = std::from_chars(minStr.data(), minStr.data() + minStr.size(), min);
            auto resMax = std::from_chars(maxStr.data(), maxStr.data() + maxStr.size(), max);

            if (resMin.ec != std::errc() || resMax.ec != std::errc()) {
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
    m_relationships   = other.m_relationships;
    m_merges          = other.m_merges;
    m_dataValidations = other.m_dataValidations;
    m_drawing         = other.m_drawing;
    m_vmlDrawing      = other.m_vmlDrawing;
    m_comments        = other.m_comments;
    m_threadedComments = other.m_threadedComments;
    m_tables          = other.m_tables;
}

XLWorksheet::XLWorksheet(XLWorksheet&& other) : XLSheetBase<XLWorksheet>(std::move(other))
{
    m_relationships   = std::move(other.m_relationships);
    m_merges          = std::move(other.m_merges);
    m_dataValidations = std::move(other.m_dataValidations);
    m_drawing         = std::move(other.m_drawing);
    m_vmlDrawing      = std::move(other.m_vmlDrawing);
    m_comments        = std::move(other.m_comments);
    m_threadedComments = std::move(other.m_threadedComments);
    m_tables          = std::move(other.m_tables);
}

XLWorksheet& XLWorksheet::operator=(const XLWorksheet& other)
{
    if (&other != this) {
        XLSheetBase<XLWorksheet>::operator=(other);
        m_relationships   = other.m_relationships;
        m_merges          = other.m_merges;
        m_dataValidations = other.m_dataValidations;
        m_drawing         = other.m_drawing;
        m_vmlDrawing      = other.m_vmlDrawing;
        m_comments        = other.m_comments;
        m_threadedComments = other.m_threadedComments;
        m_tables          = other.m_tables;
    }
    return *this;
}

XLWorksheet& XLWorksheet::operator=(XLWorksheet&& other)
{
    if (&other != this) {
        XLSheetBase<XLWorksheet>::operator=(std::move(other));
        m_relationships   = std::move(other.m_relationships);
        m_merges          = std::move(other.m_merges);
        m_dataValidations = std::move(other.m_dataValidations);
        m_drawing         = std::move(other.m_drawing);
        m_vmlDrawing      = std::move(other.m_vmlDrawing);
        m_comments        = std::move(other.m_comments);
        m_threadedComments = std::move(other.m_threadedComments);
        m_tables          = std::move(other.m_tables);
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
    const XMLNode rowNode  = getRowNode(xmlDocument().document_element().child("sheetData"), rowNumber, &m_hintRowNumber, &m_hintRowNode);
    const XMLNode cellNode = getCellNode(rowNode, columnNumber, rowNumber, {}, &m_hintColNumber, &m_hintCellNode);
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
    if (pos == std::string::npos) { return range(rangeReference, rangeReference); }
    std::string topLeft     = rangeReference.substr(0, pos);
    std::string bottomRight = rangeReference.substr(pos + 1);

    bool topIsCol    = std::all_of(topLeft.begin(), topLeft.end(), [](unsigned char c) { return std::isalpha(c); });
    bool bottomIsCol = std::all_of(bottomRight.begin(), bottomRight.end(), [](unsigned char c) { return std::isalpha(c); });
    if (topIsCol && bottomIsCol) { return range(topLeft + "1", bottomRight + std::to_string(OpenXLSX::MAX_ROWS)); }

    bool topIsRow    = std::all_of(topLeft.begin(), topLeft.end(), [](unsigned char c) { return std::isdigit(c); });
    bool bottomIsRow = std::all_of(bottomRight.begin(), bottomRight.end(), [](unsigned char c) { return std::isdigit(c); });
    if (topIsRow && bottomIsRow) { return range("A" + topLeft, XLCellReference::columnAsString(OpenXLSX::MAX_COLS) + bottomRight); }

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

void XLWorksheet::appendRow(const std::vector<XLCellValue>& values) { row(rowCount() + 1).values() = values; }

XLRow XLWorksheet::row(uint32_t rowNumber) const
{ return XLRow{getRowNode(xmlDocument().document_element().child("sheetData"), rowNumber), parentDoc().sharedStrings()}; }

XLColumn XLWorksheet::column(uint16_t columnNumber) const
{
    using namespace std::literals::string_literals;
    if (columnNumber < 1 or columnNumber > OpenXLSX::MAX_COLS)
        throw XLException("XLWorksheet::column: columnNumber "s + std::to_string(columnNumber) + " is outside allowed range [1;"s +
                          std::to_string(MAX_COLS) + "]"s);

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
            throw XLInternalError("XLWorksheet::"s + __func__ + ": column node for index "s + std::to_string(columnNumber) +
                                  "not found after splitting column nodes"s);
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
        throw XLInternalError("XLWorksheet::"s + __func__ + ": was unable to find or create node for column "s +
                              std::to_string(columnNumber));
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

void XLWorksheet::addSortCondition(const std::string& ref, uint16_t colId, bool descending)
{
    XMLNode rootNode       = xmlDocument().document_element();
    XMLNode autoFilterNode = rootNode.child("autoFilter");
    if (autoFilterNode.empty()) {
        autoFilterNode                         = appendAndGetNode(rootNode, "autoFilter", m_nodeOrder);
        autoFilterNode.append_attribute("ref") = ref.c_str();
    }

    XMLNode sortStateNode = autoFilterNode.child("sortState");
    if (sortStateNode.empty()) { sortStateNode = autoFilterNode.append_child("sortState"); }

    // Set or update the ref attribute on sortState
    if (sortStateNode.attribute("ref").empty()) { sortStateNode.append_attribute("ref") = ref.c_str(); }
    else {
        sortStateNode.attribute("ref") = ref.c_str();
    }

    // The ref string for a specific column should be just that column in the range, excluding the header
    std::string colRef = XLCellReference::columnAsString(XLCellReference(ref.substr(0, ref.find(':'))).column() + colId) +
                         std::to_string(XLCellReference(ref.substr(0, ref.find(':'))).row() + 1) + ":" +
                         XLCellReference::columnAsString(XLCellReference(ref.substr(0, ref.find(':'))).column() + colId) +
                         std::to_string(XLCellReference(ref.substr(ref.find(':') + 1)).row());

    XMLNode sortConditionNode                 = sortStateNode.append_child("sortCondition");
    sortConditionNode.append_attribute("ref") = colRef.c_str();
    if (descending) { sortConditionNode.append_attribute("descending") = "1"; }
}

void XLWorksheet::applyAutoFilter()
{
    XLAutoFilter filter = autofilterObject();
    if (!filter) return;

    std::string ref = filter.ref();
    if (ref.empty()) return;

    XLCellRange filterRange = range(XLCellReference(ref.substr(0, ref.find(':'))), XLCellReference(ref.substr(ref.find(':') + 1)));

    // Default assumption is that the first row is the header row
    uint32_t firstRow = filterRange.topLeft().row();
    uint32_t lastRow  = filterRange.bottomRight().row();
    uint16_t firstCol = filterRange.topLeft().column();
    uint16_t lastCol  = filterRange.bottomRight().column();

    // Iterate through all rows in the data area (skipping header)
    for (uint32_t rowIdx = firstRow + 1; rowIdx <= lastRow; ++rowIdx) {
        bool rowMatches = true;

        // Check each column's filter conditions
        for (uint16_t colIdx = firstCol; colIdx <= lastCol; ++colIdx) {
            uint16_t filterColId = colIdx - firstCol;

            // Optimization: check if there are actually any filter conditions
            XMLNode realColNode;
            for (auto child : xmlDocument().document_element().child("autoFilter").children("filterColumn")) {
                if (child.attribute("colId").as_uint() == filterColId) {
                    realColNode = child;
                    break;
                }
            }

            if (realColNode.empty()) continue;    // No filters on this column

            std::string cellValue  = cell(rowIdx, colIdx).value().getString();
            bool        colMatches = false;

            // 1. Basic Value Filters (<filters><filter val="..."/></filters>)
            XMLNode filtersNode = realColNode.child("filters");
            if (!filtersNode.empty()) {
                colMatches = false;
                for (XMLNode f : filtersNode.children("filter")) {
                    if (f.attribute("val").value() == cellValue) {
                        colMatches = true;
                        break;
                    }
                }
            }
            
            // 2. Custom Filters (<customFilters and="1">...)
            XMLNode customFiltersNode = realColNode.child("customFilters");
            if (!customFiltersNode.empty()) {
                bool isAnd = customFiltersNode.attribute("and").as_bool(false);
                bool matchedFirst = false;
                bool matchedSecond = false;
                
                auto evaluateCustomFilter = [](XMLNode f, const std::string& cvStr) -> bool {
                    if (f.empty()) return false;
                    std::string op = f.attribute("operator").value();
                    std::string valStr = f.attribute("val").value();
                    if (op.empty()) op = "equal"; // Default if omitted
                    
                    try {
                        double cv = std::stod(cvStr);
                        double thresh = std::stod(valStr);
                        if (op == "equal") return cv == thresh;
                        if (op == "notEqual") return cv != thresh;
                        if (op == "greaterThan") return cv > thresh;
                        if (op == "greaterThanOrEqual") return cv >= thresh;
                        if (op == "lessThan") return cv < thresh;
                        if (op == "lessThanOrEqual") return cv <= thresh;
                    } catch(...) {
                        // String comparison fallback
                        if (op == "equal") return cvStr == valStr;
                        if (op == "notEqual") return cvStr != valStr;
                    }
                    return false;
                };

                auto childIt = customFiltersNode.children("customFilter").begin();
                if (childIt != customFiltersNode.children("customFilter").end()) {
                    matchedFirst = evaluateCustomFilter(*childIt, cellValue);
                    ++childIt;
                    if (childIt != customFiltersNode.children("customFilter").end()) {
                        matchedSecond = evaluateCustomFilter(*childIt, cellValue);
                        colMatches = isAnd ? (matchedFirst && matchedSecond) : (matchedFirst || matchedSecond);
                    } else {
                        colMatches = matchedFirst; // Only 1 condition
                    }
                }
            }

            if ((!filtersNode.empty() || !customFiltersNode.empty()) && !colMatches) {
                rowMatches = false;
                break;
            }
        }

        if (rowMatches) { row(rowIdx).setHidden(false); }
        else {
            row(rowIdx).setHidden(true);
        }
    }
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
    uint32_t rows          = rowCount();
    uint16_t cols          = columnCount();
    auto     rootNode      = xmlDocument().document_element();
    auto     dimensionNode = rootNode.child("dimension");
    if (dimensionNode.empty()) dimensionNode = rootNode.prepend_child("dimension");
    dimensionNode.remove_children();
    if (rows == 0 || cols == 0) { dimensionNode.attribute("ref").set_value("A1"); }
    else {
        std::string ref  = "A1:" + XLCellReference::columnAsString(cols) + std::to_string(rows);
        auto        attr = dimensionNode.attribute("ref");
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

    uint16_t         num = 0;
    std::string_view numStr(xmlPath.data() + pos, pos2 - pos);
    std::from_chars(numStr.data(), numStr.data() + numStr.size(), num);
    return num;
}

std::optional<XLCell> XLWorksheet::peekCell(const std::string& ref) const { return peekCell(XLCellReference(ref)); }
std::optional<XLCell> XLWorksheet::peekCell(const XLCellReference& ref) const { return peekCell(ref.row(), ref.column()); }

std::optional<XLCell> XLWorksheet::peekCell(uint32_t rowNumber, uint16_t columnNumber) const
{
    XMLNode cellNode = findCellNode(findRowNode(xmlDocument().document_element().child("sheetData"), rowNumber), columnNumber);
    if (!cellNode) return std::nullopt;
    return XLCell(cellNode, parentDoc().sharedStrings());
}

void XLWorksheet::addSparkline(const std::string& location, const std::string& dataRange, XLSparklineType type)
{
    XLSparklineOptions options;
    options.type = type;
    addSparkline(location, dataRange, options);
}

void XLWorksheet::addSparkline(const std::string& location, const std::string& dataRange, const XLSparklineOptions& options)
{
    XMLNode rootNode = xmlDocument().document_element();

    // Ensure x14 namespace is registered at the root worksheet level
    if (rootNode.attribute("xmlns:x14").empty()) {
        rootNode.append_attribute("xmlns:x14") = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    }

    // Add x14 to mc:Ignorable attribute
    auto ignorableAttr = rootNode.attribute("mc:Ignorable");
    if (!ignorableAttr.empty()) {
        std::string ignorable = ignorableAttr.value();
        if (ignorable.find("x14 ") == std::string::npos && ignorable.find(" x14") == std::string::npos && ignorable != "x14") {
            ignorableAttr.set_value((ignorable + " x14").c_str());
        }
    }
    else {
        rootNode.append_attribute("mc:Ignorable") = "x14ac x14";
    }

    XMLNode extLst = rootNode.child("extLst");
    if (extLst.empty()) { extLst = appendAndGetNode(rootNode, "extLst", m_nodeOrder); }

    XMLNode extNode = extLst.find_child_by_attribute("ext", "uri", "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}");
    if (extNode.empty()) {
        extNode                         = extLst.append_child("ext");
        extNode.append_attribute("uri") = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
    }

    XMLNode sparklineGroups = extNode.child("x14:sparklineGroups");
    if (sparklineGroups.empty()) {
        sparklineGroups                              = extNode.append_child("x14:sparklineGroups");
        sparklineGroups.append_attribute("xmlns:xm") = "http://schemas.microsoft.com/office/excel/2006/main";
    }

    XMLNode sparklineGroup                                 = sparklineGroups.append_child("x14:sparklineGroup");
    sparklineGroup.append_attribute("displayEmptyCellsAs") = options.displayEmptyCellsAs.c_str();

    if (options.markers) sparklineGroup.append_attribute("markers").set_value("1");
    if (options.high) sparklineGroup.append_attribute("high").set_value("1");
    if (options.low) sparklineGroup.append_attribute("low").set_value("1");
    if (options.first) sparklineGroup.append_attribute("first").set_value("1");
    if (options.last) sparklineGroup.append_attribute("last").set_value("1");
    if (options.negative) sparklineGroup.append_attribute("negative").set_value("1");
    if (options.displayXAxis) sparklineGroup.append_attribute("displayXAxis").set_value("1");

    // Colors
    sparklineGroup.append_child("x14:colorSeries").append_attribute("rgb").set_value(options.seriesColor.c_str());
    sparklineGroup.append_child("x14:colorNegative").append_attribute("rgb").set_value(options.negativeColor.c_str());
    sparklineGroup.append_child("x14:colorAxis").append_attribute("rgb").set_value("FF000000"); // Base line color
    sparklineGroup.append_child("x14:colorMarkers").append_attribute("rgb").set_value(options.markersColor.c_str());
    sparklineGroup.append_child("x14:colorFirst").append_attribute("rgb").set_value(options.firstMarkerColor.c_str());
    sparklineGroup.append_child("x14:colorLast").append_attribute("rgb").set_value(options.lastMarkerColor.c_str());
    sparklineGroup.append_child("x14:colorHigh").append_attribute("rgb").set_value(options.highMarkerColor.c_str());
    sparklineGroup.append_child("x14:colorLow").append_attribute("rgb").set_value(options.lowMarkerColor.c_str());

    XLSparkline sparkline(sparklineGroup);
    sparkline.setType(options.type);
    sparkline.setLocation(location);

    // Format the data range to include the sheet name (e.g. "Sheet1!A1:B2") if it doesn't already have one
    std::string formattedRange = dataRange;
    if (formattedRange.find('!') == std::string::npos) {
        std::string sName = name();
        // Quote sheet name if it contains spaces
        if (sName.find(' ') != std::string::npos) { formattedRange = "'" + sName + "'!" + dataRange; }
        else {
            formattedRange = sName + "!" + dataRange;
        }
    }
    sparkline.setDataRange(formattedRange);
}

XLStreamReader XLWorksheet::streamReader() const { return XLStreamReader(this); }

XLStreamWriter XLWorksheet::streamWriter()
{
    if (m_xmlData->m_isStreamed && !m_xmlData->m_streamFilePath.empty()) {
        std::error_code ec;
        if (std::filesystem::exists(m_xmlData->m_streamFilePath, ec)) { std::filesystem::remove(m_xmlData->m_streamFilePath, ec); }
        m_xmlData->m_streamFilePath.clear();
        m_xmlData->m_isStreamed = false;
    }

    XMLDocument& doc       = xmlDocument();
    XMLNode      root      = doc.document_element();
    XMLNode      sheetData = root.child("sheetData");

    if (sheetData.empty()) { sheetData = appendAndGetNode(root, "sheetData", m_nodeOrder); }

    struct StringWriter : pugi::xml_writer
    {
        std::string result;
        void        write(const void* data, size_t size) override { result.append(static_cast<const char*>(data), size); }
    } sw;

    doc.save(sw, "", pugi::format_raw | pugi::format_no_declaration);
    std::string xmlStr = sw.result;

    std::string topHalf    = "";
    std::string bottomHalf = "";

    // Split accurately
    size_t pos = xmlStr.find("<sheetData/>");
    if (pos != std::string::npos) {
        topHalf    = xmlStr.substr(0, pos) + "<sheetData>";
        bottomHalf = "</sheetData>" + xmlStr.substr(pos + 12);
    }
    else {
        pos = xmlStr.find("</sheetData>");
        if (pos != std::string::npos) {
            // We have existing rows: <sheetData><row ... /></sheetData>
            // So we take everything UP TO </sheetData> as topHalf, allowing stream to append inside sheetData.
            topHalf    = xmlStr.substr(0, pos);
            bottomHalf = xmlStr.substr(pos);    // starts with </sheetData>
        }
        else {
            pos = xmlStr.find("<sheetData></sheetData>");
            if (pos != std::string::npos) {
                topHalf    = xmlStr.substr(0, pos) + "<sheetData>";
                bottomHalf = "</sheetData>" + xmlStr.substr(pos + 23);
            }
            else {
                size_t endTag = xmlStr.find("</worksheet>");
                topHalf       = xmlStr.substr(0, endTag) + "<sheetData>";
                bottomHalf    = "</sheetData></worksheet>";
            }
        }
    }

    XLStreamWriter writer(this);
    writer.m_bottomHalf = bottomHalf;

    if (writer.m_stream.is_open()) {
        std::string header = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
        writer.m_stream << header << topHalf;
    }

    m_xmlData->m_isStreamed     = true;
    m_xmlData->m_streamFilePath = writer.getTempFilePath();

    return writer;
}

void XLWorksheet::autoFitColumn(uint16_t columnNumber)
{
    float maxWidth = 0.0f;
    // Iterate through all rows
    for (uint32_t r = 1; r <= rowCount(); ++r) {
        auto c = cell(r, columnNumber);
        if (c.value().type() != XLValueType::Empty) {
            std::string text = c.value().getString();
            
            // Basic estimation algorithm counting UTF-8 glyphs rather than raw bytes.
            // CJK characters (3 bytes usually) typically render wider (approx 2.0 Excel units)
            // ASCII characters render narrower (approx 1.1 Excel units for Calibri 11)
            float estimatedWidth = 0.0f;
            for (size_t i = 0; i < text.length();) {
                unsigned char byte = text[i];
                if ((byte & 0x80) == 0) {
                    // ASCII
                    estimatedWidth += 1.1f;
                    i += 1;
                } else if ((byte & 0xE0) == 0xC0) {
                    // 2-byte sequence
                    estimatedWidth += 1.5f;
                    i += 2;
                } else if ((byte & 0xF0) == 0xE0) {
                    // 3-byte sequence (Typical for CJK)
                    estimatedWidth += 2.1f;
                    i += 3;
                } else if ((byte & 0xF8) == 0xF0) {
                    // 4-byte sequence (Emojis, rare symbols)
                    estimatedWidth += 2.5f;
                    i += 4;
                } else {
                    // Invalid UTF-8, just step forward
                    estimatedWidth += 1.0f;
                    i += 1;
                }
            }
            
            if (estimatedWidth > maxWidth) { maxWidth = estimatedWidth; }
        }
    }

    // Add padding and cap width
    maxWidth += 1.5f;
    if (maxWidth > 255.0f) maxWidth = 255.0f;
    if (maxWidth < 8.43f) maxWidth = 8.43f;    // Default width

    column(columnNumber).setWidth(maxWidth);
}

// =============================================================================
// Row/Column Insert & Delete — private helpers
// =============================================================================

/**
 * @details Rewrite a single cell-address string applying row and/or column offsets.
 * Components that carry an absolute-reference '$' marker are left unchanged.
 * A component whose numeric value after shift would be <= 0 is not shifted (clamped).
 *
 * Supported formats: "A1", "$A1", "A$1", "$A$1", and sheet-qualified "Sheet1!A1".
 * If the address contains '!' the sheet prefix is preserved verbatim.
 */
std::string XLWorksheet::shiftCellRef(std::string_view ref, int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol)
{
    // Handle sheet-qualified refs like "Sheet1!A1" or "'My Sheet'!B5"
    std::string      sheetPrefix;
    std::string_view localRef = ref;
    auto             bangPos  = ref.rfind('!');
    if (bangPos != std::string_view::npos) {
        sheetPrefix = std::string(ref.substr(0, bangPos + 1));
        localRef    = ref.substr(bangPos + 1);
    }

    // Parse: optional '$' + column letters + optional '$' + row digits
    std::size_t i      = 0;
    bool        absCol = false;
    if (i < localRef.size() && localRef[i] == '$') {
        absCol = true;
        ++i;
    }

    std::size_t colStart = i;
    while (i < localRef.size() && std::isalpha(static_cast<unsigned char>(localRef[i]))) ++i;
    std::size_t colEnd = i;    // localRef[colStart..colEnd) = column letters

    bool absRow = false;
    if (i < localRef.size() && localRef[i] == '$') {
        absRow = true;
        ++i;
    }

    std::size_t rowStart = i;
    while (i < localRef.size() && std::isdigit(static_cast<unsigned char>(localRef[i]))) ++i;
    std::size_t rowEnd = i;

    // If we couldn't parse a complete cell address just return unchanged
    if (colEnd == colStart || rowEnd == rowStart || i != localRef.size()) return std::string(ref);

    std::string colStr(localRef.substr(colStart, colEnd - colStart));
    std::string rowStr(localRef.substr(rowStart, rowEnd - rowStart));

    uint16_t col = XLCellReference::columnAsNumber(colStr);
    uint32_t row = 0;
    std::from_chars(rowStr.data(), rowStr.data() + rowStr.size(), row);

    // Apply column shift (only if not absolute and column is in the affected range)
    if (!absCol && colDelta != 0 && col >= fromCol) {
        int32_t newCol = static_cast<int32_t>(col) + colDelta;
        if (newCol >= 1) col = static_cast<uint16_t>(newCol);
    }
    // Apply row shift (only if not absolute and row is in the affected range)
    if (!absRow && rowDelta != 0 && row >= fromRow) {
        int32_t newRow = static_cast<int32_t>(row) + rowDelta;
        if (newRow >= 1) row = static_cast<uint32_t>(newRow);
    }

    std::string result = sheetPrefix;
    if (absCol) result += '$';
    result += XLCellReference::columnAsString(col);
    if (absRow) result += '$';
    result += std::to_string(row);
    return result;
}

/**
 * @details Tokenise a formula at a character level and rewrite every cell
 * reference / range token without altering anything else (strings, numbers,
 * function names, operators).  This is a lightweight round-trip transformer —
 * it does not evaluate the formula.
 */
std::string XLWorksheet::shiftFormulaRefs(std::string_view formula, int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol)
{
    if (formula.empty()) return {};
    std::string out;
    out.reserve(formula.size() + 16);

    std::size_t       i   = 0;
    const std::size_t len = formula.size();

    while (i < len) {
        char c = formula[i];

        // Pass string literals unchanged
        if (c == '"') {
            out += c;
            ++i;
            while (i < len) {
                if (formula[i] == '"') {
                    out += formula[i];
                    ++i;
                    if (i < len && formula[i] == '"') {
                        out += formula[i];
                        ++i;
                    }    // escaped quote
                    else
                        break;
                }
                else {
                    out += formula[i];
                    ++i;
                }
            }
            continue;
        }

        // Potential cell reference: starts with '$' or an alpha char.
        // The preceding character must NOT be alpha/digit (i.e. not inside a name).
        bool prevIsAlnum = (i > 0 && std::isalnum(static_cast<unsigned char>(formula[i - 1])));

        if (!prevIsAlnum && (c == '$' || std::isalpha(static_cast<unsigned char>(c)))) {
            // Collect the candidate token: letters/digits/$/'!'
            std::size_t start = i;
            if (i < len && formula[i] == '$') ++i;    // leading $
            std::size_t colLetterStart = i;
            while (i < len && std::isalpha(static_cast<unsigned char>(formula[i]))) ++i;
            std::size_t colLetterEnd = i;
            if (i < len && formula[i] == '$') ++i;    // optional $ before row
            std::size_t rowDigitStart = i;
            while (i < len && std::isdigit(static_cast<unsigned char>(formula[i]))) ++i;
            std::size_t rowDigitEnd = i;

            bool looksLikeCellRef = (colLetterEnd > colLetterStart) && (rowDigitEnd > rowDigitStart);

            // Check for sheet-qualified reference coming BEFORE the cell ref,
            // e.g. the '!' was already consumed as part of an identifier segment.
            // We handle it by looking for a trailing '!' right after a pure-alpha+digit candidate.

            if (looksLikeCellRef) {
                std::string_view candidate = formula.substr(start, i - start);
                // Check if this is a range (has ':' following it)
                if (i < len && formula[i] == ':') {
                    ++i;    // consume ':'
                    std::size_t start2 = i;
                    if (i < len && formula[i] == '$') ++i;
                    while (i < len && std::isalpha(static_cast<unsigned char>(formula[i]))) ++i;
                    if (i < len && formula[i] == '$') ++i;
                    while (i < len && std::isdigit(static_cast<unsigned char>(formula[i]))) ++i;
                    std::string_view candidate2 = formula.substr(start2, i - start2);

                    out += shiftCellRef(candidate, rowDelta, colDelta, fromRow, fromCol);
                    out += ':';
                    out += shiftCellRef(candidate2, rowDelta, colDelta, fromRow, fromCol);
                }
                else {
                    // Check for 'Name!ref' — if the alpha segment contains '!' we should not shift
                    if (candidate.find('!') != std::string_view::npos || colLetterEnd - colLetterStart > 3) {
                        // Not a simple cell ref (function name or cross-sheet — already captured)
                        out += candidate;
                    }
                    else {
                        out += shiftCellRef(candidate, rowDelta, colDelta, fromRow, fromCol);
                    }
                }
            }
            else {
                // Not a cell ref (function name, TRUE/FALSE, etc.) — emit verbatim
                out += formula.substr(start, i - start);
            }
            continue;
        }

        out += c;
        ++i;
    }
    return out;
}

/**
 * @details Walk all <row> nodes in sheetData.
 * For positive delta: iterate in reverse (last row first) to avoid collisions.
 * For negative delta:
 *   1. First erase all <row> nodes in the deleted band [fromRow, fromRow+|delta|).
 *   2. Then iterate forward from the first affected row and decrement r= attributes.
 */
void XLWorksheet::shiftSheetDataRows(int32_t delta, uint32_t fromRow)
{
    if (delta == 0) return;
    XMLNode sheetData = xmlDocument().document_element().child("sheetData");
    if (sheetData.empty()) return;

    if (delta < 0) {
        // Step 1: remove rows inside the deleted band
        uint32_t delLast = fromRow + static_cast<uint32_t>(-delta) - 1;
        // We already call deleteRow(rowNumber) before shiftSheetDataRows for each row,
        // so by the time we get here those rows are gone. Nothing to erase in this helper.
        // (The public deleteRow handles removal; shiftSheetDataRows only slides.)

        // Step 2: slide remaining rows down by |delta| (i.e. subtract |delta|)
        XMLNode rowNode = sheetData.first_child_of_type(pugi::node_element);
        while (!rowNode.empty()) {
            auto    r    = static_cast<uint32_t>(rowNode.attribute("r").as_ullong());
            XMLNode next = rowNode.next_sibling_of_type(pugi::node_element);
            if (r >= fromRow) {
                uint32_t newR = static_cast<uint32_t>(static_cast<int32_t>(r) + delta);
                rowNode.attribute("r").set_value(newR);
                // Update each cell in this row
                for (XMLNode cellNode = rowNode.first_child_of_type(pugi::node_element); !cellNode.empty();
                     cellNode         = cellNode.next_sibling_of_type(pugi::node_element))
                {
                    std::string_view cellRef = cellNode.attribute("r").value();
                    // Extract column letters (before digits) and append new row number
                    std::size_t pos = 0;
                    while (pos < cellRef.size() && !std::isdigit(static_cast<unsigned char>(cellRef[pos])) && cellRef[pos] != '$') ++pos;
                    // skip possible '$' before row digits
                    if (pos < cellRef.size() && cellRef[pos] == '$') ++pos;
                    std::string colPart;
                    for (std::size_t k = 0; k < pos; ++k)
                        if (cellRef[k] != '$') colPart += cellRef[k];
                    std::string newRef = colPart + std::to_string(newR);
                    cellNode.attribute("r").set_value(newRef.c_str());
                }
            }
            rowNode = next;
        }
        (void)delLast;
    }
    else {
        // For insertion: iterate rows in REVERSE (highest first) so we don't
        // shift a row that then gets shifted again.
        XMLNode lastRow = sheetData.last_child_of_type(pugi::node_element);
        XMLNode rowNode = lastRow;
        while (!rowNode.empty()) {
            auto    r    = static_cast<uint32_t>(rowNode.attribute("r").as_ullong());
            XMLNode prev = rowNode.previous_sibling_of_type(pugi::node_element);
            if (r >= fromRow) {
                uint32_t newR = static_cast<uint32_t>(static_cast<int32_t>(r) + delta);
                rowNode.attribute("r").set_value(newR);
                for (XMLNode cellNode = rowNode.first_child_of_type(pugi::node_element); !cellNode.empty();
                     cellNode         = cellNode.next_sibling_of_type(pugi::node_element))
                {
                    std::string_view cellRef = cellNode.attribute("r").value();
                    std::size_t      pos     = 0;
                    while (pos < cellRef.size() && !std::isdigit(static_cast<unsigned char>(cellRef[pos])) && cellRef[pos] != '$') ++pos;
                    if (pos < cellRef.size() && cellRef[pos] == '$') ++pos;
                    std::string colPart;
                    for (std::size_t k = 0; k < pos; ++k)
                        if (cellRef[k] != '$') colPart += cellRef[k];
                    std::string newRef = colPart + std::to_string(newR);
                    cellNode.attribute("r").set_value(newRef.c_str());
                }
            }
            rowNode = prev;
        }
    }
}

/**
 * @details Walk every cell in every row and shift the column-letter portion of its
 * r="XN" attribute.  Also update the <cols> node.
 */
void XLWorksheet::shiftSheetDataCols(int32_t delta, uint16_t fromCol)
{
    if (delta == 0) return;
    XMLNode sheetData = xmlDocument().document_element().child("sheetData");
    if (sheetData.empty()) return;

    for (XMLNode rowNode = sheetData.first_child_of_type(pugi::node_element); !rowNode.empty();
         rowNode         = rowNode.next_sibling_of_type(pugi::node_element))
    {
        // Collect cells that need shifting into a vector (to avoid iterator invalidation)
        std::vector<std::pair<XMLNode, std::string>> updates;
        for (XMLNode cellNode = (delta > 0 ? rowNode.last_child_of_type(pugi::node_element)    // reverse for insert
                                           : rowNode.first_child_of_type(pugi::node_element));
             !cellNode.empty();
             cellNode =
                 (delta > 0 ? cellNode.previous_sibling_of_type(pugi::node_element) : cellNode.next_sibling_of_type(pugi::node_element)))
        {
            std::string_view cellRef = cellNode.attribute("r").value();
            // Extract column number
            // cellRef may have leading '$' for absolute refs — strip for column extraction
            const char* p      = cellRef.data();
            bool        absCol = (*p == '$');
            if (absCol) ++p;
            uint16_t    col        = 0;
            const char* digitStart = p;
            while (*digitStart && std::isalpha(static_cast<unsigned char>(*digitStart))) ++digitStart;
            col = XLCellReference::columnAsNumber(std::string(p, digitStart - p));

            // Cells in the deleted band are already removed by deleteColumn() before this
            // function is called, so we simply slide all surviving cells in affected columns.
            if (col >= fromCol) {
                int32_t newCol = static_cast<int32_t>(col) + delta;
                if (newCol < 1) continue;
                // Extract row part (after the digits-start pointer)
                std::string newRef;
                if (absCol) newRef += '$';
                newRef += XLCellReference::columnAsString(static_cast<uint16_t>(newCol));
                newRef += std::string(digitStart);    // row digits (including possible '$' if absRow)
                updates.emplace_back(cellNode, newRef);
            }
        }
        for (auto& [node, ref] : updates) node.attribute("r").set_value(ref.c_str());
    }

    shiftColsNode(delta, fromCol);
}

void XLWorksheet::shiftColsNode(int32_t delta, uint16_t fromCol)
{
    XMLNode colsNode = xmlDocument().document_element().child("cols");
    if (colsNode.empty()) return;

    for (XMLNode colNode = (delta > 0 ? colsNode.last_child_of_type(pugi::node_element) : colsNode.first_child_of_type(pugi::node_element));
         !colNode.empty();
         colNode = (delta > 0 ? colNode.previous_sibling_of_type(pugi::node_element) : colNode.next_sibling_of_type(pugi::node_element)))
    {
        auto minVal = static_cast<uint16_t>(colNode.attribute("min").as_uint());
        auto maxVal = static_cast<uint16_t>(colNode.attribute("max").as_uint());
        if (maxVal < fromCol) continue;
        if (minVal >= fromCol) colNode.attribute("min").set_value(static_cast<uint16_t>(static_cast<int32_t>(minVal) + delta));
        if (maxVal >= fromCol) colNode.attribute("max").set_value(static_cast<uint16_t>(static_cast<int32_t>(maxVal) + delta));
    }
}

void XLWorksheet::shiftFormulas(int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol)
{
    if (rowDelta == 0 && colDelta == 0) return;
    XMLNode sheetData = xmlDocument().document_element().child("sheetData");
    if (sheetData.empty()) return;

    for (XMLNode rowNode = sheetData.first_child_of_type(pugi::node_element); !rowNode.empty();
         rowNode         = rowNode.next_sibling_of_type(pugi::node_element))
    {
        for (XMLNode cellNode = rowNode.first_child_of_type(pugi::node_element); !cellNode.empty();
             cellNode         = cellNode.next_sibling_of_type(pugi::node_element))
        {
            XMLNode fNode = cellNode.child("f");
            if (!fNode.empty()) {
                std::string formula = fNode.text().get();
                std::string shifted = shiftFormulaRefs(formula, rowDelta, colDelta, fromRow, fromCol);
                if (shifted != formula) fNode.text().set(shifted.c_str());
            }
        }
    }
}

/**
 * @details Adjust row and column indices in xdr:from and xdr:to anchor nodes.
 * Drawing XML uses 0-based row/col indices.
 */
void XLWorksheet::shiftDrawingAnchors(int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol)
{
    if (rowDelta == 0 && colDelta == 0) return;
    if (!m_drawing.valid()) return;

    XMLNode root = m_drawing.xmlDocument().document_element();
    if (root.empty()) return;

    // Shift 0-based thresholds
    uint32_t fromRow0 = (fromRow > 0) ? fromRow - 1 : 0;
    uint16_t fromCol0 = (fromCol > 0) ? fromCol - 1 : 0;

    auto shiftAnchorNode = [&](XMLNode anchor) {
        for (const char* child : {"xdr:from", "xdr:to"}) {
            XMLNode an = anchor.child(child);
            if (an.empty()) continue;

            XMLNode rowN = an.child("xdr:row");
            if (!rowN.empty() && rowDelta != 0) {
                auto r = static_cast<uint32_t>(rowN.text().as_uint());
                if (r >= fromRow0) {
                    int32_t newR = static_cast<int32_t>(r) + rowDelta;
                    if (newR < 0) newR = 0;
                    rowN.text().set(static_cast<unsigned int>(newR));
                }
            }
            XMLNode colN = an.child("xdr:col");
            if (!colN.empty() && colDelta != 0) {
                auto co = static_cast<uint16_t>(colN.text().as_uint());
                if (co >= fromCol0) {
                    int32_t newC = static_cast<int32_t>(co) + colDelta;
                    if (newC < 0) newC = 0;
                    colN.text().set(static_cast<unsigned int>(newC));
                }
            }
        }
    };

    for (XMLNode anchor = root.first_child_of_type(pugi::node_element); !anchor.empty();
         anchor         = anchor.next_sibling_of_type(pugi::node_element))
    {
        shiftAnchorNode(anchor);
    }
}

/**
 * @details Walk the dataValidations element and shift the sqref of each rule.
 * sqref is a space-separated list of cell ranges.
 */
void XLWorksheet::shiftDataValidations(int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol)
{
    if (rowDelta == 0 && colDelta == 0) return;
    XMLNode dvNode = xmlDocument().document_element().child("dataValidations");
    if (dvNode.empty()) return;

    for (XMLNode dv = dvNode.first_child_of_type(pugi::node_element); !dv.empty(); dv = dv.next_sibling_of_type(pugi::node_element)) {
        XMLAttribute sqrefAttr = dv.attribute("sqref");
        if (sqrefAttr.empty()) continue;

        std::string        sqref = sqrefAttr.value();
        std::istringstream ss(sqref);
        std::string        segment;
        std::string        newSqref;
        while (std::getline(ss, segment, ' ')) {
            if (segment.empty()) continue;
            if (!newSqref.empty()) newSqref += ' ';
            auto colon = segment.find(':');
            if (colon != std::string::npos) {
                newSqref += shiftCellRef(segment.substr(0, colon), rowDelta, colDelta, fromRow, fromCol);
                newSqref += ':';
                newSqref += shiftCellRef(segment.substr(colon + 1), rowDelta, colDelta, fromRow, fromCol);
            }
            else {
                newSqref += shiftCellRef(segment, rowDelta, colDelta, fromRow, fromCol);
            }
        }
        if (!newSqref.empty()) sqrefAttr.set_value(newSqref.c_str());
    }
}

void XLWorksheet::shiftAutoFilter(int32_t rowDelta, int32_t colDelta, uint32_t fromRow, uint16_t fromCol)
{
    if (rowDelta == 0 && colDelta == 0) return;
    XMLNode afNode = xmlDocument().document_element().child("autoFilter");
    if (afNode.empty()) return;

    XMLAttribute refAttr = afNode.attribute("ref");
    if (refAttr.empty()) return;

    std::string ref   = refAttr.value();
    auto        colon = ref.find(':');
    std::string newRef;
    if (colon != std::string::npos) {
        newRef = shiftCellRef(ref.substr(0, colon), rowDelta, colDelta, fromRow, fromCol) + ":" +
                 shiftCellRef(ref.substr(colon + 1), rowDelta, colDelta, fromRow, fromCol);
    }
    else {
        newRef = shiftCellRef(ref, rowDelta, colDelta, fromRow, fromCol);
    }
    refAttr.set_value(newRef.c_str());
}

// =============================================================================
// Row/Column Insert & Delete — public API
// =============================================================================

bool XLWorksheet::insertRow(uint32_t rowNumber, uint32_t count)
{
    using namespace std::literals::string_literals;
    if (rowNumber < 1 || count == 0) throw XLInputError("XLWorksheet::insertRow: rowNumber must be >= 1 and count > 0"s);

    auto delta = static_cast<int32_t>(count);

    // Shift sheetData (no rows to erase for insert)
    shiftSheetDataRows(delta, rowNumber);

    // Shift all subsystems
    if (m_merges.valid()) m_merges.shiftRows(delta, rowNumber);

    shiftFormulas(delta, 0, rowNumber, 1);
    shiftDrawingAnchors(delta, 0, rowNumber, 1);
    shiftDataValidations(delta, 0, rowNumber, 1);
    shiftAutoFilter(delta, 0, rowNumber, 1);

    return true;
}

bool XLWorksheet::deleteRow(uint32_t rowNumber, uint32_t count)
{
    using namespace std::literals::string_literals;
    if (rowNumber < 1 || count == 0) throw XLInputError("XLWorksheet::deleteRow: rowNumber must be >= 1 and count > 0"s);

    auto delta = -static_cast<int32_t>(count);

    // Step 1: physically remove the target row nodes
    for (uint32_t r = rowNumber; r < rowNumber + count; ++r) deleteRow(r);    // existing 1-argument version

    // Step 2: slide subsequent rows up (fromRow = rowNumber + count because those
    //         are rows that survived and now need to move up by `count`)
    shiftSheetDataRows(delta, rowNumber + count);

    // Step 3: shift all subsystems (fromRow = rowNumber: affects everything from the first
    //         deleted row onward)
    if (m_merges.valid()) m_merges.shiftRows(delta, rowNumber);

    shiftFormulas(delta, 0, rowNumber + count, 1);
    shiftDrawingAnchors(delta, 0, rowNumber + count, 1);
    shiftDataValidations(delta, 0, rowNumber + count, 1);
    shiftAutoFilter(delta, 0, rowNumber + count, 1);

    return true;
}

bool XLWorksheet::insertColumn(uint16_t colNumber, uint16_t count)
{
    using namespace std::literals::string_literals;
    if (colNumber < 1 || count == 0) throw XLInputError("XLWorksheet::insertColumn: colNumber must be >= 1 and count > 0"s);

    auto delta = static_cast<int32_t>(count);

    shiftSheetDataCols(delta, colNumber);

    if (m_merges.valid()) m_merges.shiftCols(delta, colNumber);

    shiftFormulas(0, delta, 1, colNumber);
    shiftDrawingAnchors(0, delta, 1, colNumber);
    shiftDataValidations(0, delta, 1, colNumber);
    shiftAutoFilter(0, delta, 1, colNumber);

    return true;
}

bool XLWorksheet::deleteColumn(uint16_t colNumber, uint16_t count)
{
    using namespace std::literals::string_literals;
    if (colNumber < 1 || count == 0) throw XLInputError("XLWorksheet::deleteColumn: colNumber must be >= 1 and count > 0"s);

    auto delta = -static_cast<int32_t>(count);

    // Remove cells in deleted columns from all rows
    XMLNode sheetData = xmlDocument().document_element().child("sheetData");
    if (!sheetData.empty()) {
        for (XMLNode rowNode = sheetData.first_child_of_type(pugi::node_element); !rowNode.empty();
             rowNode         = rowNode.next_sibling_of_type(pugi::node_element))
        {
            std::vector<XMLNode> toRemove;
            for (XMLNode cellNode = rowNode.first_child_of_type(pugi::node_element); !cellNode.empty();
                 cellNode         = cellNode.next_sibling_of_type(pugi::node_element))
            {
                std::string_view cellRef = cellNode.attribute("r").value();
                const char*      p       = cellRef.data();
                if (*p == '$') ++p;
                const char* digitStart = p;
                while (*digitStart && std::isalpha(static_cast<unsigned char>(*digitStart))) ++digitStart;
                uint16_t col = XLCellReference::columnAsNumber(std::string(p, digitStart - p));
                if (col >= colNumber && col < static_cast<uint16_t>(colNumber + count)) toRemove.push_back(cellNode);
            }
            for (auto& node : toRemove) rowNode.remove_child(node);
        }
    }

    // Slide remaining columns left
    shiftSheetDataCols(delta, colNumber + count);

    if (m_merges.valid()) m_merges.shiftCols(delta, colNumber);

    shiftFormulas(0, delta, 1, colNumber + count);
    shiftDrawingAnchors(0, delta, 1, colNumber + count);
    shiftDataValidations(0, delta, 1, colNumber + count);
    shiftAutoFilter(0, delta, 1, colNumber + count);

    return true;
}

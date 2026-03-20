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
    return range(rangeReference.substr(0, pos), rangeReference.substr(pos + 1, std::string::npos));
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

void XLWorksheet::groupRows(uint32_t rowFirst, uint32_t rowLast, uint8_t outlineLevel, bool collapsed)
{
    auto sheetPr = xmlDocument().document_element().child("sheetPr");
    if (sheetPr.empty()) sheetPr = xmlDocument().document_element().prepend_child("sheetPr");
    auto outlinePr = sheetPr.child("outlinePr");
    if (outlinePr.empty()) outlinePr = sheetPr.append_child("outlinePr");
    if (outlinePr.attribute("summaryBelow").empty()) outlinePr.append_attribute("summaryBelow").set_value("1");
    if (outlinePr.attribute("summaryRight").empty()) outlinePr.append_attribute("summaryRight").set_value("1");

    for (uint32_t r = rowFirst; r <= rowLast; ++r) {
        auto rowObj = row(r);
        rowObj.setOutlineLevel(outlineLevel);
        if (collapsed) rowObj.setHidden(true);
    }
    if (collapsed && outlineLevel > 0) {
        auto rowObj = row(rowLast + 1);
        rowObj.setCollapsed(true);
    }
}

void XLWorksheet::groupColumns(uint16_t colFirst, uint16_t colLast, uint8_t outlineLevel, bool collapsed)
{
    auto sheetPr = xmlDocument().document_element().child("sheetPr");
    if (sheetPr.empty()) sheetPr = xmlDocument().document_element().prepend_child("sheetPr");
    auto outlinePr = sheetPr.child("outlinePr");
    if (outlinePr.empty()) outlinePr = sheetPr.append_child("outlinePr");
    if (outlinePr.attribute("summaryBelow").empty()) outlinePr.append_attribute("summaryBelow").set_value("1");
    if (outlinePr.attribute("summaryRight").empty()) outlinePr.append_attribute("summaryRight").set_value("1");

    for (uint16_t c = colFirst; c <= colLast; ++c) {
        auto colObj = column(c);
        colObj.setOutlineLevel(outlineLevel);
        if (collapsed) colObj.setHidden(true);
    }
    if (collapsed && outlineLevel > 0) {
        auto colObj = column(colLast + 1);
        colObj.setCollapsed(true);
    }
}

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

XLMergeCells& XLWorksheet::merges()
{
    if (!m_merges.valid()) m_merges = XLMergeCells(xmlDocument().document_element(), m_nodeOrder);
    return m_merges;
}

XLDataValidations& XLWorksheet::dataValidations()
{
    if (m_dataValidations.empty()) m_dataValidations = XLDataValidations(xmlDocument().document_element());
    return m_dataValidations;
}

void XLWorksheet::mergeCells(XLCellRange const& rangeToMerge, bool emptyHiddenCells)
{
    if (rangeToMerge.numRows() * rangeToMerge.numColumns() < 2) {
        using namespace std::literals::string_literals;
        throw XLInputError("XLWorksheet::"s + __func__ + ": rangeToMerge must comprise at least 2 cells"s);
    }
    merges().appendMerge(rangeToMerge.address());
    if (emptyHiddenCells) {
        XLCellIterator it = rangeToMerge.begin();
        ++it;
        while (it != rangeToMerge.end()) {
            it->clear(XLKeepCellStyle);
            ++it;
        }
    }
}

void XLWorksheet::mergeCells(const std::string& rangeReference, bool emptyHiddenCells)
{ mergeCells(range(rangeReference), emptyHiddenCells); }

void XLWorksheet::unmergeCells(XLCellRange const& rangeToUnmerge)
{
    int32_t mergeIndex = merges().findMerge(rangeToUnmerge.address());
    if (mergeIndex != -1) merges().deleteMerge(mergeIndex);
    else {
        using namespace std::literals::string_literals;
        throw XLInputError("XLWorksheet::"s + __func__ + ": merged range "s + rangeToUnmerge.address() + " does not exist"s);
    }
}

void XLWorksheet::unmergeCells(const std::string& rangeReference) { unmergeCells(range(rangeReference)); }

XLStyleIndex XLWorksheet::getColumnFormat(uint16_t columnNumber) const { return column(columnNumber).format(); }
XLStyleIndex XLWorksheet::getColumnFormat(const std::string& columnNumber) const
{ return getColumnFormat(XLCellReference::columnAsNumber(columnNumber)); }

bool XLWorksheet::setColumnFormat(uint16_t columnNumber, XLStyleIndex cellFormatIndex)
{
    if (!column(columnNumber).setFormat(cellFormatIndex)) return false;
    XLRowRange allRows = rows();
    for (XLRowIterator rowIt = allRows.begin(); rowIt != allRows.end(); ++rowIt) {
        XLCell curCell = rowIt->findCell(columnNumber);
        if (curCell and !curCell.setCellFormat(cellFormatIndex)) return false;
    }
    return true;
}
bool XLWorksheet::setColumnFormat(const std::string& columnNumber, XLStyleIndex cellFormatIndex)
{ return setColumnFormat(XLCellReference::columnAsNumber(columnNumber), cellFormatIndex); }

XLStyleIndex XLWorksheet::getRowFormat(uint16_t rowNumber) const { return row(rowNumber).format(); }

bool XLWorksheet::setRowFormat(uint32_t rowNumber, XLStyleIndex cellFormatIndex)
{
    if (!row(rowNumber).setFormat(cellFormatIndex)) return false;
    XLRowDataRange rowCells = row(rowNumber).cells();
    for (XLRowDataIterator cellIt = rowCells.begin(); cellIt != rowCells.end(); ++cellIt)
        if (!cellIt->setCellFormat(cellFormatIndex)) return false;
    return true;
}

XLConditionalFormats XLWorksheet::conditionalFormats() const { return XLConditionalFormats(xmlDocument().document_element()); }

void XLWorksheet::addConditionalFormatting(const std::string& sqref, const XLCfRule& rule)
{
    auto cfList = conditionalFormats();
    XLConditionalFormat cfTarget;
    bool found = false;
    for (size_t i = 0; i < cfList.count(); ++i) {
        if (cfList[i].sqref() == sqref) {
            cfTarget = cfList[i];
            found = true;
            break;
        }
    }
    if (!found) {
        size_t idx = cfList.create();
        cfTarget = cfList[idx];
        cfTarget.setSqref(sqref);
    }
    
    // Compute global max priority across all conditional formatting rules in this worksheet
    uint16_t globalMaxPrio = 0;
    for (size_t i = 0; i < cfList.count(); ++i) {
        auto cfmt = cfList[i];
        for (size_t j = 0; j < cfmt.cfRules().count(); ++j) {
            uint16_t prio = cfmt.cfRules()[j].priority();
            if (prio > globalMaxPrio) {
                globalMaxPrio = prio;
            }
        }
    }
    
    // We create a mutable copy to update priority before inserting it
    XLCfRule ruleWithPrio = rule;
    ruleWithPrio.setPriority(globalMaxPrio + 1);
    
    // The underlying XLCfRules::create also computes maxPrio internally but only within its parent CF node.
    // By setting it explicitly beforehand and relying on the `copyFrom` behavior, we can ensure it's copied properly.
    cfTarget.cfRules().create(ruleWithPrio);
}

void XLWorksheet::addConditionalFormatting(const std::string& sqref, const XLCfRule& rule, const XLDxf& dxf)
{
    // Need a mutable copy to update dxfId before passing to create()
    XLCfRule ruleWithDxf = rule;
    if (!dxf.empty()) {
        XLStyleIndex dxfId = parentDoc().styles().addDxf(dxf);
        ruleWithDxf.setDxfId(dxfId);
    }
    addConditionalFormatting(sqref, ruleWithDxf);
}

void XLWorksheet::addConditionalFormatting(const XLCellRange& range, const XLCfRule& rule)
{
    addConditionalFormatting(range.address(), rule);
}

void XLWorksheet::addConditionalFormatting(const XLCellRange& range, const XLCfRule& rule, const XLDxf& dxf)
{
    addConditionalFormatting(range.address(), rule, dxf);
}

void XLWorksheet::removeConditionalFormatting(const std::string& sqref)
{
    auto rootNode = xmlDocument().document_element();
    for (XMLNode node = rootNode.child("conditionalFormatting"); not node.empty(); ) {
        XMLNode next = node.next_sibling("conditionalFormatting");
        if (std::string(node.attribute("sqref").value()) == sqref) {
            rootNode.remove_child(node);
        }
        node = next;
    }
}

void XLWorksheet::removeConditionalFormatting(const XLCellRange& range)
{
    removeConditionalFormatting(range.address());
}

void XLWorksheet::clearAllConditionalFormatting()
{
    auto rootNode = xmlDocument().document_element();
    while (XMLNode node = rootNode.child("conditionalFormatting")) {
        rootNode.remove_child(node);
    }
}

XLPageMargins XLWorksheet::pageMargins() const
{
    XMLNode rootNode = xmlDocument().document_element();
    XMLNode node     = rootNode.child("pageMargins");
    if (node.empty()) node = appendAndGetNode(rootNode, "pageMargins", m_nodeOrder);
    return XLPageMargins(node);
}

XLPrintOptions XLWorksheet::printOptions() const
{
    XMLNode rootNode = xmlDocument().document_element();
    XMLNode node     = rootNode.child("printOptions");
    if (node.empty()) node = appendAndGetNode(rootNode, "printOptions", m_nodeOrder);
    return XLPrintOptions(node);
}

XLPageSetup XLWorksheet::pageSetup() const
{
    XMLNode rootNode = xmlDocument().document_element();
    XMLNode node     = rootNode.child("pageSetup");
    if (node.empty()) node = appendAndGetNode(rootNode, "pageSetup", m_nodeOrder);
    return XLPageSetup(node);
}

void XLWorksheet::setPrintArea(const std::string& sqref)
{
    uint32_t localSheetId = index() - 1; // index() is 1-based, localSheetId is 0-based
    parentDoc().workbook().definedNames().append("_xlnm.Print_Area", "'" + name() + "'!" + sqref, localSheetId);
}

void XLWorksheet::setPrintTitleRows(uint32_t firstRow, uint32_t lastRow)
{
    uint32_t localSheetId = index() - 1;
    std::string rowRef = "'" + name() + "'!$" + std::to_string(firstRow) + ":$" + std::to_string(lastRow);
    
    auto currentName = parentDoc().workbook().definedNames().get("_xlnm.Print_Titles", localSheetId);
    if (currentName.valid()) {
        std::string currentRef = currentName.refersTo();
        if (currentRef.find(":") != std::string::npos && currentRef.find("$1") == std::string::npos) {
            rowRef = rowRef + "," + currentRef;
            parentDoc().workbook().definedNames().remove("_xlnm.Print_Titles", localSheetId);
        }
    }
    
    parentDoc().workbook().definedNames().append("_xlnm.Print_Titles", rowRef, localSheetId);
}

void XLWorksheet::setPrintTitleCols(uint16_t firstCol, uint16_t lastCol)
{
    uint32_t localSheetId = index() - 1;
    std::string colRef = "'" + name() + "'!$" + XLCellReference::columnAsString(firstCol) + ":$" + XLCellReference::columnAsString(lastCol);
    
    auto currentName = parentDoc().workbook().definedNames().get("_xlnm.Print_Titles", localSheetId);
    if (currentName.valid()) {
        std::string currentRef = currentName.refersTo();
        if (currentRef.find(":") != std::string::npos && currentRef.find("$A") == std::string::npos) {
            colRef = currentRef + "," + colRef;
            parentDoc().workbook().definedNames().remove("_xlnm.Print_Titles", localSheetId);
        }
    }
    
    parentDoc().workbook().definedNames().append("_xlnm.Print_Titles", colRef, localSheetId);
}

bool XLWorksheet::protectSheet(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "sheet", (set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::protectObjects(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "objects", (set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::protectScenarios(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "scenarios", (set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}

bool XLWorksheet::allowInsertColumns(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "insertColumns", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowInsertRows(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "insertRows", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowDeleteColumns(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "deleteColumns", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowDeleteRows(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "deleteRows", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowFormatCells(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "formatCells", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowFormatColumns(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "formatColumns", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowFormatRows(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "formatRows", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowInsertHyperlinks(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "insertHyperlinks", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowSort(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "sort", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowAutoFilter(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "autoFilter", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowPivotTables(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "pivotTables", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowSelectLockedCells(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "selectLockedCells", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowSelectUnlockedCells(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "selectUnlockedCells", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}

bool XLWorksheet::setPasswordHash(std::string hash)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "password", hash, XLKeepAttributes, m_nodeOrder).empty() == false;
}

bool XLWorksheet::setPassword(std::string password)
{ return setPasswordHash(password.length() ? ExcelPasswordHashAsString(password) : std::string("")); }
bool XLWorksheet::clearPassword() { return setPasswordHash(""); }
bool XLWorksheet::clearSheetProtection() { return xmlDocument().document_element().remove_child("sheetProtection"); }

bool XLWorksheet::sheetProtected() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return false;
    return node.attribute("sheet").as_bool(false);
}
bool XLWorksheet::objectsProtected() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return false;
    return node.attribute("objects").as_bool(false);
}
bool XLWorksheet::scenariosProtected() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return false;
    return node.attribute("scenarios").as_bool(false);
}

bool XLWorksheet::insertColumnsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("insertColumns").as_bool(true);
}
bool XLWorksheet::insertRowsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("insertRows").as_bool(true);
}
bool XLWorksheet::deleteColumnsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("deleteColumns").as_bool(true);
}
bool XLWorksheet::deleteRowsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("deleteRows").as_bool(true);
}
bool XLWorksheet::formatCellsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("formatCells").as_bool(true);
}
bool XLWorksheet::formatColumnsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("formatColumns").as_bool(true);
}
bool XLWorksheet::formatRowsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("formatRows").as_bool(true);
}
bool XLWorksheet::insertHyperlinksAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("insertHyperlinks").as_bool(true);
}
bool XLWorksheet::sortAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("sort").as_bool(true);
}
bool XLWorksheet::autoFilterAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("autoFilter").as_bool(true);
}
bool XLWorksheet::pivotTablesAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("pivotTables").as_bool(true);
}
bool XLWorksheet::selectLockedCellsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("selectLockedCells").as_bool(false);
}
bool XLWorksheet::selectUnlockedCellsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("selectUnlockedCells").as_bool(false);
}

std::string XLWorksheet::passwordHash() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return "";
    return node.attribute("password").value();
}
bool XLWorksheet::passwordIsSet() const { return !passwordHash().empty(); }

std::string XLWorksheet::sheetProtectionSummary() const
{
    return fmt::format("sheet: {}, objects: {}, scenarios: {}, insertColumns: {}, insertRows: {}, deleteColumns: {}, deleteRows: {}, formatCells: {}, formatColumns: {}, formatRows: {}, insertHyperlinks: {}, sort: {}, autoFilter: {}, pivotTables: {}, selectLockedCells: {}, selectUnlockedCells: {}, passwordHash: \"{}\"",
                       sheetProtected() ? "true" : "false", objectsProtected() ? "true" : "false", scenariosProtected() ? "true" : "false", insertColumnsAllowed() ? "allowed" : "protected", insertRowsAllowed() ? "allowed" : "protected", deleteColumnsAllowed() ? "allowed" : "protected", deleteRowsAllowed() ? "allowed" : "protected", formatCellsAllowed() ? "allowed" : "protected", formatColumnsAllowed() ? "allowed" : "protected", formatRowsAllowed() ? "allowed" : "protected", insertHyperlinksAllowed() ? "allowed" : "protected", sortAllowed() ? "allowed" : "protected", autoFilterAllowed() ? "allowed" : "protected", pivotTablesAllowed() ? "allowed" : "protected", selectLockedCellsAllowed() ? "allowed" : "protected", selectUnlockedCellsAllowed() ? "allowed" : "protected", passwordHash());
}

bool XLWorksheet::hasRelationships() const { return parentDoc().hasSheetRelationships(sheetXmlNumber()); }
bool XLWorksheet::hasDrawing() const { return parentDoc().hasSheetDrawing(sheetXmlNumber()); }
bool XLWorksheet::hasVmlDrawing() const { return parentDoc().hasSheetVmlDrawing(sheetXmlNumber()); }
bool XLWorksheet::hasComments() const { return parentDoc().hasSheetComments(sheetXmlNumber()); }
bool XLWorksheet::hasTables() const { return parentDoc().hasSheetTables(sheetXmlNumber()); }

XLRelationships& XLWorksheet::relationships()
{
    if (!m_relationships.valid()) {
        m_relationships = parentDoc().sheetRelationships(sheetXmlNumber());
    }
    if (!m_relationships.valid()) throw XLException("XLWorksheet::relationships(): could not create relationships XML");
    return m_relationships;
}

XLDrawing& XLWorksheet::drawing()
{
    if (!m_drawing.valid()) {
        XMLNode docElement = xmlDocument().document_element();
        std::ignore        = relationships();

        uint16_t sheetXmlNo = sheetXmlNumber();
        m_drawing           = parentDoc().sheetDrawing(sheetXmlNo);
        if (!m_drawing.valid()) throw XLException("XLWorksheet::drawing(): could not create drawing XML");

        std::string drawingPath         = m_drawing.getXmlPath();
        std::string drawingRelativePath = getPathARelativeToPathB(drawingPath, getXmlPath());

        XLRelationshipItem drawingRelationship;
        if (!m_relationships.targetExists(drawingRelativePath))
            drawingRelationship = m_relationships.addRelationship(XLRelationshipType::Drawing, drawingRelativePath);
        else
            drawingRelationship = m_relationships.relationshipByTarget(drawingRelativePath);

        if (drawingRelationship.empty()) throw XLException("XLWorksheet::drawing(): could not add sheet relationship for Drawing");

        if (docElement.child("drawing").empty()) {
            XMLNode drawingNode = appendAndGetNode(docElement, "drawing", m_nodeOrder);
            appendAndSetAttribute(drawingNode, "r:id", drawingRelationship.id());
        }
    }

    return m_drawing;
}

void XLWorksheet::addImage(const std::string& name, const std::string& data, uint32_t row, uint32_t col, uint32_t width, uint32_t height)
{
    std::string internalPath = parentDoc().addImage(name, data);
    XLDrawing& drw = drawing();
    std::string drawingPath       = drw.getXmlPath();
    std::string imageRelativePath = getPathARelativeToPathB(internalPath, drawingPath);

    XLRelationshipItem imgRel;
    if (!drw.relationships().targetExists(imageRelativePath))
        imgRel = drw.relationships().addRelationship(XLRelationshipType::Image, imageRelativePath);
    else
        imgRel = drw.relationships().relationshipByTarget(imageRelativePath);

    drw.addImage(imgRel.id(), name, "", row, col, width, height);
}

void XLWorksheet::addScaledImage(const std::string& name, const std::string& data, uint32_t row, uint32_t col, double scalingFactor)
{
    std::string internalPath = parentDoc().addImage(name, data);
    XLDrawing& drw = drawing();
    std::string drawingPath       = drw.getXmlPath();
    std::string imageRelativePath = getPathARelativeToPathB(internalPath, drawingPath);

    XLRelationshipItem imgRel;
    if (!drw.relationships().targetExists(imageRelativePath))
        imgRel = drw.relationships().addRelationship(XLRelationshipType::Image, imageRelativePath);
    else
        imgRel = drw.relationships().relationshipByTarget(imageRelativePath);

    drw.addScaledImage(imgRel.id(), name, "", data, row, col, scalingFactor);
}

XLChart XLWorksheet::addChart([[maybe_unused]] XLChartType type, std::string_view name, uint32_t row, uint32_t col, uint32_t width, uint32_t height)
{
    // 1. Create Chart File in Document
    XLChart chart = parentDoc().createChart();

    // 2. Get Drawing for the sheet
    XLDrawing& drw = drawing();
    std::string drawingPath     = drw.getXmlPath();
    std::string chartRelativePath = getPathARelativeToPathB(chart.getXmlPath(), drawingPath);

    // 3. Add Relationship in drawing to the new chart file
    XLRelationshipItem chartRel;
    if (!drw.relationships().targetExists(chartRelativePath)) {
        chartRel = drw.relationships().addRelationship(XLRelationshipType::Chart, chartRelativePath);
    } else {
        chartRel = drw.relationships().relationshipByTarget(chartRelativePath);
    }

    // 4. Add Chart Anchor in Drawing
    drw.addChartAnchor(chartRel.id(), name, row, col, width, height);

    return chart;
}

std::vector<XLDrawingItem> XLWorksheet::images()
{
    std::vector<XLDrawingItem> result;
    if (!m_drawing.valid()) {
        std::string drawingRelativePath = "";
        for (auto& rel : relationships().relationships()) {
            if (rel.type() == XLRelationshipType::Drawing) {
                drawingRelativePath = rel.target();
                break;
            }
        }
        if (drawingRelativePath.empty()) return result;
        m_drawing = drawing();
    }

    uint32_t count = m_drawing.imageCount();
    for (uint32_t i = 0; i < count; ++i) { result.push_back(m_drawing.image(i)); }
    return result;
}

XLVmlDrawing& XLWorksheet::vmlDrawing()
{
    if (!m_vmlDrawing.valid()) {
        XMLNode docElement = xmlDocument().document_element();
        std::ignore        = relationships();

        uint16_t sheetXmlNo = sheetXmlNumber();
        m_vmlDrawing        = parentDoc().sheetVmlDrawing(sheetXmlNo);
        if (!m_vmlDrawing.valid()) throw XLException("XLWorksheet::vmlDrawing(): could not create drawing XML");
        std::string        drawingRelativePath = getPathARelativeToPathB(m_vmlDrawing.getXmlPath(), getXmlPath());
        XLRelationshipItem vmlDrawingRelationship;
        if (!m_relationships.targetExists(drawingRelativePath))
            vmlDrawingRelationship = m_relationships.addRelationship(XLRelationshipType::VMLDrawing, drawingRelativePath);
        else
            vmlDrawingRelationship = m_relationships.relationshipByTarget(drawingRelativePath);
        if (vmlDrawingRelationship.empty())
            throw XLException("XLWorksheet::vmlDrawing(): could not add determine sheet relationship for VML Drawing");
        XMLNode legacyDrawing = appendAndGetNode(docElement, "legacyDrawing", m_nodeOrder);
        if (legacyDrawing.empty()) throw XLException("XLWorksheet::vmlDrawing(): could not add <legacyDrawing> element to worksheet XML");
        appendAndSetAttribute(legacyDrawing, "r:id", vmlDrawingRelationship.id());
    }

    return m_vmlDrawing;
}

XLComments& XLWorksheet::comments()
{
    if (!m_comments.valid()) {
        std::ignore = relationships();
        std::ignore = vmlDrawing();

        uint16_t sheetXmlNo = sheetXmlNumber();
        m_comments = parentDoc().sheetComments(sheetXmlNo);
        if (!m_comments.valid()) throw XLException("XLWorksheet::comments(): could not create comments XML");
        m_comments.setVmlDrawing(m_vmlDrawing);
        std::string commentsRelativePath = getPathARelativeToPathB(m_comments.getXmlPath(), getXmlPath());
        if (!m_relationships.targetExists(commentsRelativePath))
            m_relationships.addRelationship(XLRelationshipType::Comments, commentsRelativePath);
    }

    return m_comments;
}

XLTableCollection& XLWorksheet::tables()
{
    if (!m_tables.valid()) {
        m_tables = XLTableCollection(xmlDocument().document_element(), this);
    }

    return m_tables;
}

void XLWorksheet::addHyperlink(std::string_view cellRef, std::string_view url, std::string_view tooltip)
{
    removeHyperlink(cellRef);
    std::ignore = relationships();
    const auto rel = m_relationships.addRelationship(XLRelationshipType::Hyperlink, std::string(url), true);
    XMLNode docElement     = xmlDocument().document_element();
    XMLNode hyperlinksNode = docElement.child("hyperlinks");
    if (hyperlinksNode.empty()) { hyperlinksNode = appendAndGetNode(docElement, "hyperlinks", m_nodeOrder); }
    XMLNode hyperlinkNode = hyperlinksNode.append_child("hyperlink");
    hyperlinkNode.append_attribute("ref").set_value(std::string(cellRef).c_str());
    hyperlinkNode.append_attribute("r:id").set_value(rel.id().c_str());
    if (!tooltip.empty()) { hyperlinkNode.append_attribute("tooltip").set_value(std::string(tooltip).c_str()); }
}

void XLWorksheet::addInternalHyperlink(std::string_view cellRef, std::string_view location, std::string_view tooltip)
{
    removeHyperlink(cellRef);
    XMLNode docElement     = xmlDocument().document_element();
    XMLNode hyperlinksNode = docElement.child("hyperlinks");
    if (hyperlinksNode.empty()) { hyperlinksNode = appendAndGetNode(docElement, "hyperlinks", m_nodeOrder); }
    XMLNode hyperlinkNode = hyperlinksNode.append_child("hyperlink");
    hyperlinkNode.append_attribute("ref").set_value(std::string(cellRef).c_str());
    hyperlinkNode.append_attribute("location").set_value(std::string(location).c_str());
    hyperlinkNode.append_attribute("display").set_value(std::string(location).c_str());
    if (!tooltip.empty()) { hyperlinkNode.append_attribute("tooltip").set_value(std::string(tooltip).c_str()); }
}

bool XLWorksheet::hasHyperlink(std::string_view cellRef) const
{
    XMLNode docElement     = xmlDocument().document_element();
    XMLNode hyperlinksNode = docElement.child("hyperlinks");
    if (hyperlinksNode.empty()) return false;
    for (auto link : hyperlinksNode.children("hyperlink")) {
        if (std::string_view(link.attribute("ref").value()) == cellRef) return true;
    }
    return false;
}

std::string XLWorksheet::getHyperlink(std::string_view cellRef) const
{
    XMLNode docElement     = xmlDocument().document_element();
    XMLNode hyperlinksNode = docElement.child("hyperlinks");
    if (hyperlinksNode.empty()) return "";
    for (auto link : hyperlinksNode.children("hyperlink")) {
        if (std::string_view(link.attribute("ref").value()) == cellRef) {
            if (link.attribute("location")) return link.attribute("location").value();
            if (link.attribute("r:id")) {
                const auto rId = link.attribute("r:id").value();
                return const_cast<XLWorksheet*>(this)->relationships().relationshipById(rId).target();
            }
        }
    }
    return "";
}

void XLWorksheet::removeHyperlink(std::string_view cellRef)
{
    XMLNode docElement     = xmlDocument().document_element();
    XMLNode hyperlinksNode = docElement.child("hyperlinks");
    if (hyperlinksNode.empty()) return;
    for (auto link : hyperlinksNode.children("hyperlink")) {
        if (std::string_view(link.attribute("ref").value()) == cellRef) {
            hyperlinksNode.remove_child(link);
            break;
        }
    }
    if (hyperlinksNode.first_child().empty()) { docElement.remove_child(hyperlinksNode); }
}

bool XLWorksheet::hasPanes() const
{
    auto sheetViews = xmlDocument().document_element().child("sheetViews");
    if (sheetViews.empty()) return false;
    auto sheetView = sheetViews.child("sheetView");
    if (sheetView.empty()) return false;
    return !sheetView.child("pane").empty();
}

XMLNode XLWorksheet::prepareSheetViewForPanes()
{
    auto docElement = xmlDocument().document_element();
    auto sheetViews = docElement.child("sheetViews");
    if (sheetViews.empty()) { sheetViews = appendAndGetNode(docElement, "sheetViews", m_nodeOrder); }
    auto sheetView = sheetViews.child("sheetView");
    if (sheetView.empty()) {
        sheetView = sheetViews.append_child("sheetView");
        sheetView.append_attribute("workbookViewId").set_value(0);
    }
    sheetView.remove_child("pane");
    while (auto selection = sheetView.child("selection")) { sheetView.remove_child(selection); }
    return sheetView;
}

void XLWorksheet::freezePanes(uint16_t column, uint32_t row)
{
    XMLNode sheetView = prepareSheetViewForPanes();
    if (column == 0 && row == 0) return;
    auto pane = appendAndGetNode(sheetView, "pane", XLSheetViewNodeOrder);
    if (column > 0) pane.append_attribute("xSplit").set_value(column);
    if (row > 0) pane.append_attribute("ySplit").set_value(row);
    const std::string address = XLCellReference(row + 1, static_cast<uint16_t>(column + 1)).address();
    pane.append_attribute("topLeftCell").set_value(address.c_str());
    pane.append_attribute("state").set_value("frozen");
    const char* activePaneStr = (column > 0 && row > 0) ? "bottomRight" : (column > 0 ? "topRight" : "bottomLeft");
    pane.append_attribute("activePane").set_value(activePaneStr);
    if (column > 0 && row > 0) {
        auto sel1 = sheetView.insert_child_after("selection", pane);
        sel1.append_attribute("pane").set_value("topRight");
        auto sel2 = sheetView.insert_child_after("selection", sel1);
        sel2.append_attribute("pane").set_value("bottomLeft");
        auto sel3 = sheetView.insert_child_after("selection", sel2);
        sel3.append_attribute("pane").set_value("bottomRight");
        sel3.append_attribute("activeCell").set_value(address.c_str());
        sel3.append_attribute("sqref").set_value(address.c_str());
    }
    else {
        auto sel = sheetView.insert_child_after("selection", pane);
        sel.append_attribute("pane").set_value(activePaneStr);
        sel.append_attribute("activeCell").set_value(address.c_str());
        sel.append_attribute("sqref").set_value(address.c_str());
    }
}

void XLWorksheet::freezePanes(std::string_view cellRef)
{
    XLCellReference ref(std::string{cellRef});
    freezePanes(static_cast<uint16_t>(ref.column() - 1), ref.row() - 1);
}

void XLWorksheet::splitPanes(double xSplit, double ySplit, std::string_view topLeftCell, XLPane activePane)
{
    XMLNode sheetView = prepareSheetViewForPanes();
    if (xSplit == 0 && ySplit == 0) return;
    auto pane = appendAndGetNode(sheetView, "pane", XLSheetViewNodeOrder);
    if (xSplit > 0) pane.append_attribute("xSplit").set_value(xSplit);
    if (ySplit > 0) pane.append_attribute("ySplit").set_value(ySplit);
    if (!topLeftCell.empty()) pane.append_attribute("topLeftCell").set_value(std::string(topLeftCell).c_str());
    pane.append_attribute("activePane").set_value(XLPaneToString(activePane).c_str());
    pane.append_attribute("state").set_value("split");
    auto sel = sheetView.insert_child_after("selection", pane);
    sel.append_attribute("pane").set_value(XLPaneToString(activePane).c_str());
}

void XLWorksheet::clearPanes()
{
    auto docElement = xmlDocument().document_element();
    auto sheetViews = docElement.child("sheetViews");
    if (sheetViews.empty()) return;
    auto sheetView = sheetViews.child("sheetView");
    if (sheetView.empty()) return;
    sheetView.remove_child("pane");
    while (auto selection = sheetView.child("selection")) { sheetView.remove_child(selection); }
}

void XLWorksheet::setZoom(uint16_t scale)
{
    if (scale < 10 || scale > 400) throw XLInputError("XLWorksheet::setZoom: scale must be between 10 and 400");
    auto docElement = xmlDocument().document_element();
    auto sheetViews = docElement.child("sheetViews");
    if (sheetViews.empty()) { sheetViews = appendAndGetNode(docElement, "sheetViews", m_nodeOrder); }
    auto sheetView = sheetViews.child("sheetView");
    if (sheetView.empty()) {
        sheetView = sheetViews.append_child("sheetView");
        sheetView.append_attribute("workbookViewId").set_value(0);
    }
    appendAndSetAttribute(sheetView, "zoomScale", std::to_string(scale));
    appendAndSetAttribute(sheetView, "zoomScaleNormal", std::to_string(scale));
}

uint16_t XLWorksheet::zoom() const
{
    auto docElement = xmlDocument().document_element();
    auto sheetViews = docElement.child("sheetViews");
    if (sheetViews.empty()) return 100;
    auto sheetView = sheetViews.child("sheetView");
    if (sheetView.empty()) return 100;
    auto zoomAttr = sheetView.attribute("zoomScale");
    if (zoomAttr.empty()) return 100;
    return static_cast<uint16_t>(zoomAttr.as_uint());
}

void XLWorksheet::insertRowBreak(uint32_t row)
{
    auto docElement = xmlDocument().document_element();
    auto rowBreaks = docElement.child("rowBreaks");
    if (rowBreaks.empty()) {
        rowBreaks = appendAndGetNode(docElement, "rowBreaks", m_nodeOrder);
        rowBreaks.append_attribute("count").set_value(0);
        rowBreaks.append_attribute("manualBreakCount").set_value(0);
    }
    for (auto brk : rowBreaks.children("brk")) if (brk.attribute("id").as_uint() == row) return;
    auto brk = rowBreaks.append_child("brk");
    brk.append_attribute("id").set_value(row);
    brk.append_attribute("max").set_value(16383);
    brk.append_attribute("man").set_value("1");
    uint32_t count = rowBreaks.attribute("count").as_uint(0) + 1;
    rowBreaks.attribute("count").set_value(count);
    uint32_t manualCount = rowBreaks.attribute("manualBreakCount").as_uint(0) + 1;
    rowBreaks.attribute("manualBreakCount").set_value(manualCount);
}

void XLWorksheet::removeRowBreak(uint32_t row)
{
    auto docElement = xmlDocument().document_element();
    auto rowBreaks = docElement.child("rowBreaks");
    if (rowBreaks.empty()) return;
    for (auto brk : rowBreaks.children("brk")) {
        if (brk.attribute("id").as_uint() == row) {
            rowBreaks.remove_child(brk);
            uint32_t count = rowBreaks.attribute("count").as_uint(1) - 1;
            rowBreaks.attribute("count").set_value(count);
            uint32_t manualCount = rowBreaks.attribute("manualBreakCount").as_uint(1) - 1;
            rowBreaks.attribute("manualBreakCount").set_value(manualCount);
            if (count == 0) docElement.remove_child(rowBreaks);
            break;
        }
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

#include <string>

#include "XLWorksheet.hpp"
#include "XLDocument.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

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

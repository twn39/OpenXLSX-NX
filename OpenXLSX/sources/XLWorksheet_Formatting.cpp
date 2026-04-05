#include <string>

#include "XLDocument.hpp"
#include "XLUtilities.hpp"
#include "XLWorksheet.hpp"

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

XLCellRange XLWorksheet::mergeCells(XLCellRange const& rangeToMerge, bool emptyHiddenCells)
{
    if (static_cast<uint64_t>(rangeToMerge.numRows()) * rangeToMerge.numColumns() < 2) {
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
    return rangeToMerge;
}

XLCellRange XLWorksheet::mergeCells(const std::string& rangeReference, bool emptyHiddenCells)
{ return mergeCells(range(rangeReference), emptyHiddenCells); }

void XLWorksheet::unmergeCells(XLCellRange const& rangeToUnmerge)
{
    int32_t mergeIndex = merges().findMerge(rangeToUnmerge.address());
    if (mergeIndex != -1)
        merges().deleteMerge(mergeIndex);
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
    auto                cfList = conditionalFormats();
    XLConditionalFormat cfTarget;
    bool                found = false;
    for (size_t i = 0; i < cfList.count(); ++i) {
        if (cfList[i].sqref() == sqref) {
            cfTarget = cfList[i];
            found    = true;
            break;
        }
    }
    if (!found) {
        size_t idx = cfList.create();
        cfTarget   = cfList[idx];
        cfTarget.setSqref(sqref);
    }

    // Compute global max priority across all conditional formatting rules in this worksheet
    uint16_t globalMaxPrio = 0;
    for (size_t i = 0; i < cfList.count(); ++i) {
        auto cfmt = cfList[i];
        for (size_t j = 0; j < cfmt.cfRules().count(); ++j) {
            uint16_t prio = cfmt.cfRules()[j].priority();
            if (prio > globalMaxPrio) { globalMaxPrio = prio; }
        }
    }

    // We create a mutable copy to update priority before inserting it
    XLCfRule ruleWithPrio = rule;
    ruleWithPrio.setPriority(globalMaxPrio + 1);

    // The underlying XLCfRules::create also computes maxPrio internally but only within its parent CF node.
    // By setting it explicitly beforehand and relying on the `copyFrom` behavior, we can ensure it's copied properly.
    size_t newRuleIdx = cfTarget.cfRules().create(ruleWithPrio);
    XLCfRule newRule = cfTarget.cfRules()[newRuleIdx];

    // --- Handle Advanced ExtLst for DataBar (OOXML Excel 2010+ features) ---
    if (newRule.type() == XLCfType::DataBar) {
        auto dbNode = newRule.node().child("dataBar");
        if (auto fakeExtLst = dbNode.child("extLst")) {
            XMLNode extNode;
            for (auto ext = fakeExtLst.child("ext"); ext; ext = ext.next_sibling("ext")) {
                if (std::string(ext.attribute("uri").value()) == "{B469CE28-E4FAB-4e90-B891-B227B70C6CDA}") {
                    extNode = ext;
                    break;
                }
            }
            if (extNode) {
                // 1. Generate an ID (fake GUID, but matching valid UUID structure)
                char guidStr[64];
                snprintf(guidStr, sizeof(guidStr), "{%08X-0000-4000-8000-000000000000}", globalMaxPrio + 1);

                // 2. Add real extLst to the cfRule element (NOT dataBar)
                auto ruleExtLst = newRule.node().child("extLst");
                if (!ruleExtLst) ruleExtLst = newRule.node().append_child("extLst");
                auto ruleExt = ruleExtLst.append_child("ext");
                ruleExt.append_attribute("uri") = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";
                ruleExt.append_attribute("xmlns:x14") = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
                ruleExt.append_child("x14:id").text().set(guidStr);

                // 3. Create or find worksheet-level extLst
                auto wsNode = xmlDocument().document_element();
                auto wsExtLst = wsNode.child("extLst");
                if (!wsExtLst) wsExtLst = wsNode.append_child("extLst");
                XMLNode wsExt;
                for (auto e = wsExtLst.child("ext"); e; e = e.next_sibling("ext")) {
                    if (std::string(e.attribute("uri").value()) == "{78C0D931-6437-407d-A8EE-F0AAD7539E65}") {
                        wsExt = e;
                        break;
                    }
                }
                if (!wsExt) {
                    wsExt = wsExtLst.append_child("ext");
                    wsExt.append_attribute("uri") = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
                    wsExt.append_attribute("xmlns:x14") = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
                }
                auto condFmts = wsExt.child("x14:conditionalFormattings");
                if (!condFmts) condFmts = wsExt.append_child("x14:conditionalFormattings");

                // 4. Find or create the conditionalFormatting container inside the ext
                XMLNode x14condFmt;
                for (auto c = condFmts.child("x14:conditionalFormatting"); c; c = c.next_sibling("x14:conditionalFormatting")) {
                    if (std::string(c.child("xm:sqref").text().get()) == sqref) {
                        x14condFmt = c;
                        break;
                    }
                }
                if (!x14condFmt) {
                    x14condFmt = condFmts.append_child("x14:conditionalFormatting");
                    x14condFmt.append_attribute("xmlns:xm") = "http://schemas.microsoft.com/office/excel/2006/main";
                    x14condFmt.append_child("xm:sqref").text().set(sqref.c_str());
                }

                // 5. Append the rule BEFORE <xm:sqref>
                auto x14cfRule = x14condFmt.insert_child_before("x14:cfRule", x14condFmt.child("xm:sqref"));
                x14cfRule.append_attribute("type") = "dataBar";
                x14cfRule.append_attribute("id") = guidStr;

                // Move the actual <x14:dataBar> node over
                auto sourceDataBar = extNode.child("x14:dataBar");
                if (sourceDataBar) {
                    auto copiedDataBar = x14cfRule.append_copy(sourceDataBar);
                    // Add mandatory child nodes if missing (Excel requires cfvo autoMin/autoMax)
                    if (!copiedDataBar.child("x14:cfvo")) {
                        auto minCfvo = copiedDataBar.prepend_child("x14:cfvo");
                        minCfvo.append_attribute("type") = "autoMin";
                        auto maxCfvo = copiedDataBar.insert_child_after("x14:cfvo", minCfvo);
                        maxCfvo.append_attribute("type") = "autoMax";
                    }
                }

                // 6. Clean up the fake extension from the original dataBar node
                fakeExtLst.remove_child(extNode);
                if (fakeExtLst.first_child().empty()) {
                    dbNode.remove_child(fakeExtLst);
                }

                // 7. Ensure <extLst> is strictly the last element in the <worksheet> according to the OOXML Schema.
                wsNode.append_move(wsExtLst);
            }
        }
    }
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
{ addConditionalFormatting(range.address(), rule); }

void XLWorksheet::addConditionalFormatting(const XLCellRange& range, const XLCfRule& rule, const XLDxf& dxf)
{ addConditionalFormatting(range.address(), rule, dxf); }

void XLWorksheet::removeConditionalFormatting(const std::string& sqref)
{
    auto rootNode = xmlDocument().document_element();
    for (XMLNode node = rootNode.child("conditionalFormatting"); not node.empty();) {
        XMLNode next = node.next_sibling("conditionalFormatting");
        if (std::string(node.attribute("sqref").value()) == sqref) { rootNode.remove_child(node); }
        node = next;
    }
}

void XLWorksheet::removeConditionalFormatting(const XLCellRange& range) { removeConditionalFormatting(range.address()); }

void XLWorksheet::clearAllConditionalFormatting()
{
    auto rootNode = xmlDocument().document_element();
    while (XMLNode node = rootNode.child("conditionalFormatting")) { rootNode.remove_child(node); }
}

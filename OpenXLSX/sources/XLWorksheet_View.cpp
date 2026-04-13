#include <string>

#include "XLCellReference.hpp"
#include "XLDocument.hpp"
#include "XLUtilities.hpp"
#include "XLWorksheet.hpp"

#include "XLChart.hpp"
#include "XLComments.hpp"
#include "XLConditionalFormatting.hpp"
#include "XLDataValidation.hpp"
#include "XLDrawing.hpp"
#include "XLImageOptions.hpp"
#include "XLMergeCells.hpp"
#include "XLPageSetup.hpp"
#include "XLPivotTable.hpp"
#include "XLRelationships.hpp"
#include "XLSparkline.hpp"
#include "XLStreamReader.hpp"
#include "XLStreamWriter.hpp"
#include "XLTables.hpp"
#include "XLThreadedComments.hpp"

using namespace OpenXLSX;

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
    uint32_t localSheetId = index() - 1;    // index() is 1-based, localSheetId is 0-based
    parentDoc().workbook().definedNames().append("_xlnm.Print_Area", "'" + name() + "'!" + sqref, localSheetId);
}

void XLWorksheet::setPrintTitleRows(uint32_t firstRow, uint32_t lastRow)
{
    uint32_t    localSheetId = index() - 1;
    std::string rowRef       = "'" + name() + "'!$" + std::to_string(firstRow) + ":$" + std::to_string(lastRow);

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
    uint32_t    localSheetId = index() - 1;
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

void XLWorksheet::freezePanes(const std::string& topLeftCell)
{
    XLCellReference ref(topLeftCell);
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
    auto rowBreaks  = docElement.child("rowBreaks");
    if (rowBreaks.empty()) {
        rowBreaks = appendAndGetNode(docElement, "rowBreaks", m_nodeOrder);
        rowBreaks.append_attribute("count").set_value(0);
        rowBreaks.append_attribute("manualBreakCount").set_value(0);
    }
    for (auto brk : rowBreaks.children("brk"))
        if (brk.attribute("id").as_uint() == row) return;
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
    auto rowBreaks  = docElement.child("rowBreaks");
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

XLHeaderFooter XLWorksheet::headerFooter() const
{
    XMLNode rootNode = xmlDocument().document_element();
    XMLNode node     = rootNode.child("headerFooter");
    if (node.empty()) node = appendAndGetNode(rootNode, "headerFooter", m_nodeOrder);
    return XLHeaderFooter(node);
}

void XLWorksheet::insertColBreak(uint16_t col)
{
    auto docElement = xmlDocument().document_element();
    auto colBreaks  = docElement.child("colBreaks");
    if (colBreaks.empty()) {
        colBreaks = appendAndGetNode(docElement, "colBreaks", m_nodeOrder);
        colBreaks.append_attribute("count").set_value(0);
        colBreaks.append_attribute("manualBreakCount").set_value(0);
    }
    for (auto brk : colBreaks.children("brk"))
        if (brk.attribute("id").as_uint() == col) return;
    auto brk = colBreaks.append_child("brk");
    brk.append_attribute("id").set_value(col);
    brk.append_attribute("max").set_value(1048575);
    brk.append_attribute("man").set_value("1");
    uint32_t count = colBreaks.attribute("count").as_uint(0) + 1;
    colBreaks.attribute("count").set_value(count);
    uint32_t manualCount = colBreaks.attribute("manualBreakCount").as_uint(0) + 1;
    colBreaks.attribute("manualBreakCount").set_value(manualCount);
}

void XLWorksheet::removeColBreak(uint16_t col)
{
    auto docElement = xmlDocument().document_element();
    auto colBreaks  = docElement.child("colBreaks");
    if (colBreaks.empty()) return;
    for (auto brk : colBreaks.children("brk")) {
        if (brk.attribute("id").as_uint() == col) {
            colBreaks.remove_child(brk);
            uint32_t count = colBreaks.attribute("count").as_uint(1) - 1;
            colBreaks.attribute("count").set_value(count);
            uint32_t manualCount = colBreaks.attribute("manualBreakCount").as_uint(1) - 1;
            colBreaks.attribute("manualBreakCount").set_value(manualCount);
            if (count == 0) docElement.remove_child(colBreaks);
            return;
        }
    }
}

void XLWorksheet::setSheetViewMode(std::string_view mode)
{
    auto docElement = xmlDocument().document_element();
    auto sheetViews = docElement.child("sheetViews");
    if (sheetViews.empty()) { sheetViews = appendAndGetNode(docElement, "sheetViews", m_nodeOrder); }
    auto sheetView = sheetViews.child("sheetView");
    if (sheetView.empty()) {
        sheetView = sheetViews.append_child("sheetView");
        sheetView.append_attribute("workbookViewId").set_value(0);
    }
    appendAndSetAttribute(sheetView, "view", std::string(mode));
}

std::string XLWorksheet::sheetViewMode() const
{
    auto docElement = xmlDocument().document_element();
    auto sheetView = docElement.child("sheetViews").child("sheetView");
    auto attr = sheetView.attribute("view");
    return attr.empty() ? "normal" : attr.value();
}

void XLWorksheet::setShowGridLines(bool show)
{
    auto docElement = xmlDocument().document_element();
    auto sheetViews = docElement.child("sheetViews");
    if (sheetViews.empty()) { sheetViews = appendAndGetNode(docElement, "sheetViews", m_nodeOrder); }
    auto sheetView = sheetViews.child("sheetView");
    if (sheetView.empty()) {
        sheetView = sheetViews.append_child("sheetView");
        sheetView.append_attribute("workbookViewId").set_value(0);
    }
    appendAndSetAttribute(sheetView, "showGridLines", show ? "1" : "0");
}

bool XLWorksheet::showGridLines() const
{
    auto docElement = xmlDocument().document_element();
    auto attr = docElement.child("sheetViews").child("sheetView").attribute("showGridLines");
    return attr.empty() ? true : attr.as_bool();
}

void XLWorksheet::setShowRowColHeaders(bool show)
{
    auto docElement = xmlDocument().document_element();
    auto sheetViews = docElement.child("sheetViews");
    if (sheetViews.empty()) { sheetViews = appendAndGetNode(docElement, "sheetViews", m_nodeOrder); }
    auto sheetView = sheetViews.child("sheetView");
    if (sheetView.empty()) {
        sheetView = sheetViews.append_child("sheetView");
        sheetView.append_attribute("workbookViewId").set_value(0);
    }
    appendAndSetAttribute(sheetView, "showRowColHeaders", show ? "1" : "0");
}

bool XLWorksheet::showRowColHeaders() const
{
    auto docElement = xmlDocument().document_element();
    auto attr = docElement.child("sheetViews").child("sheetView").attribute("showRowColHeaders");
    return attr.empty() ? true : attr.as_bool();
}

void XLWorksheet::fitToPages(uint32_t fitToWidth, uint32_t fitToHeight)
{
    auto docElement = xmlDocument().document_element();
    
    // Enable fitToPage in sheetPr -> pageSetUpPr
    auto sheetPr = docElement.child("sheetPr");
    if (sheetPr.empty()) { sheetPr = docElement.prepend_child("sheetPr"); }
    auto pageSetUpPr = sheetPr.child("pageSetUpPr");
    if (pageSetUpPr.empty()) { pageSetUpPr = sheetPr.append_child("pageSetUpPr"); }
    appendAndSetAttribute(pageSetUpPr, "fitToPage", "1");
    
    // Set actual values in pageSetup and remove scale to prevent conflict
    auto ps = pageSetup();
    ps.setFitToWidth(fitToWidth);
    ps.setFitToHeight(fitToHeight);
    
    auto pageSetupNode = docElement.child("pageSetup");
    if (!pageSetupNode.empty()) {
        pageSetupNode.remove_attribute("scale");
    }
}

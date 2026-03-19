#include <algorithm>
#include <map>
#include <limits>
#include <fmt/format.h>
#include <pugixml.hpp>
#include "XLSheet.hpp"
#include "XLUtilities.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{
    /**
     * @brief Function for setting tab color.
     * @param xmlDocument XMLDocument object
     * @param color Thr color to set
     */
    void setTabColor(const XMLDocument& xmlDocument, const XLColor& color)
    {
        if (!xmlDocument.document_element().child("sheetPr")) xmlDocument.document_element().prepend_child("sheetPr");

        if (!xmlDocument.document_element().child("sheetPr").child("tabColor"))
            xmlDocument.document_element().child("sheetPr").prepend_child("tabColor");

        auto colorNode = xmlDocument.document_element().child("sheetPr").child("tabColor");
        for (auto attr : colorNode.attributes()) colorNode.remove_attribute(attr);

        colorNode.prepend_attribute("rgb").set_value(color.hex().c_str());
    }

    /**
     * @brief Set the tab selected property to desired value
     * @param xmlDocument
     * @param selected
     */
    void setTabSelected(const XMLDocument& xmlDocument, bool selected)
    {
        unsigned int value       = (selected ? 1 : 0);
        XMLNode      sheetView   = xmlDocument.document_element().child("sheetViews").first_child_of_type(pugi::node_element);
        XMLAttribute tabSelected = sheetView.attribute("tabSelected");
        if (tabSelected.empty())
            tabSelected = sheetView.prepend_attribute("tabSelected");
        tabSelected.set_value(value);
    }

    /**
     * @brief Function for checking if the tab is selected.
     * @param xmlDocument
     * @return
     */
    bool tabIsSelected(const XMLDocument& xmlDocument)
    {
        return xmlDocument.document_element()
            .child("sheetViews")
            .first_child_of_type(pugi::node_element)
            .attribute("tabSelected")
            .as_bool();
    }

    /**
     * @brief get the correct XLPaneState from the OOXML pane state attribute string
     * @param stateString the string as used in the OOXML
     * @return the corresponding XLPaneState enum value
     */
    std::string XLPaneStateToString(XLPaneState state)
    {
        switch (state) {
            case XLPaneState::Split:
                return "split";
            case XLPaneState::Frozen:
                return "frozen";
            case XLPaneState::FrozenSplit:
                return "frozenSplit";
            default:
                return "";
        }
    }

    /**
     * @brief inverse of XLPaneStateToString
     * @param stateString the string for which to get the XLPaneState
     */
    XLPaneState XLPaneStateFromString(std::string const& stateString)
    {
        if (stateString == "split") return XLPaneState::Split;
        if (stateString == "frozen") return XLPaneState::Frozen;
        if (stateString == "frozenSplit") return XLPaneState::FrozenSplit;
        return XLPaneState::Split;
    }

    /**
     * @brief get the correct XLPane from the OOXML pane identifier attribute string
     * @param paneString the string as used in the OOXML
     * @return the corresponding XLPane enum value
     */
    std::string XLPaneToString(XLPane pane)
    {
        switch (pane) {
            case XLPane::BottomRight:
                return "bottomRight";
            case XLPane::TopRight:
                return "topRight";
            case XLPane::BottomLeft:
                return "bottomLeft";
            case XLPane::TopLeft:
                return "topLeft";
            default:
                return "";
        }
    }

    /**
     * @brief inverse of XLPaneToString
     * @param paneString the string for which to get the XLPane
     */
    XLPane XLPaneFromString(std::string const& paneString)
    {
        if (paneString == "bottomRight") return XLPane::BottomRight;
        if (paneString == "topRight") return XLPane::TopRight;
        if (paneString == "bottomLeft") return XLPane::BottomLeft;
        if (paneString == "topLeft") return XLPane::TopLeft;
        return XLPane::BottomRight;
    }
}

XLSheet::XLSheet(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
    if (xmlData->getXmlType() == XLContentType::Worksheet)
        m_sheet = XLWorksheet(xmlData);
    else if (xmlData->getXmlType() == XLContentType::Chartsheet)
        m_sheet = XLChartsheet(xmlData);
}

XLSheetState XLSheet::visibility() const
{
    return std::visit([](auto& sheet) { return sheet.visibility(); }, m_sheet);
}

void XLSheet::setVisibility(XLSheetState state)
{
    std::visit([&](auto& sheet) { sheet.setVisibility(state); }, m_sheet);
}

XLColor XLSheet::color() const
{
    return std::visit([](auto& sheet) { return sheet.color(); }, m_sheet);
}

void XLSheet::setColor(const XLColor& color)
{
    std::visit([&](auto& sheet) { sheet.setColor(color); }, m_sheet);
}

uint16_t XLSheet::index() const
{
    return std::visit([](auto& sheet) { return sheet.index(); }, m_sheet);
}

void XLSheet::setIndex(uint16_t index)
{
    std::visit([&](auto& sheet) { sheet.setIndex(index); }, m_sheet);
}

std::string XLSheet::name() const
{
    return std::visit([](auto& sheet) { return sheet.name(); }, m_sheet);
}

void XLSheet::setName(const std::string& name)
{
    std::visit([&](auto& sheet) { sheet.setName(name); }, m_sheet);
}

bool XLSheet::isSelected() const
{
    return std::visit([](auto& sheet) { return sheet.isSelected(); }, m_sheet);
}

void XLSheet::setSelected(bool selected)
{
    std::visit([&](auto& sheet) { sheet.setSelected(selected); }, m_sheet);
}

bool XLSheet::isActive() const
{
    return std::visit([](auto& sheet) { return sheet.isActive(); }, m_sheet);
}

bool XLSheet::setActive()
{
    return std::visit([](auto& sheet) { return sheet.setActive(); }, m_sheet);
}

void XLSheet::clone(const std::string& newName)
{
    std::visit([&](auto& sheet) { sheet.clone(newName); }, m_sheet);
}

XLSheet::operator XLWorksheet() const { return std::get<XLWorksheet>(m_sheet); }
XLSheet::operator XLChartsheet() const { return std::get<XLChartsheet>(m_sheet); }

void XLSheet::print(std::basic_ostream<char>& ostr) const
{
    xmlDocument().print(ostr);
}

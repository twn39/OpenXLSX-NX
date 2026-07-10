#include "XLTableColumn.hpp"
#include "XLUtilities.hpp"
#include <string>
#include <vector>

namespace OpenXLSX
{

    XLTableColumn::XLTableColumn(XMLNode node) : m_node(node) {}

    XLTableColumn::operator bool() const { return m_node != nullptr; }

    uint32_t    XLTableColumn::id() const { return m_node.attribute("id").as_uint(); }
    std::string XLTableColumn::name() const { return m_node.attribute("name").value(); }
    void        XLTableColumn::setName(std::string_view name) { setAttr(m_node, "name", std::string(name)); }

    XLTotalsRowFunction XLTableColumn::totalsRowFunction() const
    {
        std::string funcStr = m_node.attribute("totalsRowFunction").value();
        if (funcStr == "sum") return XLTotalsRowFunction::Sum;
        if (funcStr == "min") return XLTotalsRowFunction::Min;
        if (funcStr == "max") return XLTotalsRowFunction::Max;
        if (funcStr == "average") return XLTotalsRowFunction::Average;
        if (funcStr == "count") return XLTotalsRowFunction::Count;
        return XLTotalsRowFunction::None;
    }

    void XLTableColumn::setTotalsRowFunction(XLTotalsRowFunction func)
    {
        std::string funcStr = "";
        switch (func) {
            case XLTotalsRowFunction::Sum:
                funcStr = "sum";
                break;
            case XLTotalsRowFunction::Min:
                funcStr = "min";
                break;
            case XLTotalsRowFunction::Max:
                funcStr = "max";
                break;
            case XLTotalsRowFunction::Average:
                funcStr = "average";
                break;
            case XLTotalsRowFunction::Count:
                funcStr = "count";
                break;
            default:
                break;
        }
        if (funcStr.empty())
            m_node.remove_attribute("totalsRowFunction");
        else
            setAttr(m_node, "totalsRowFunction", funcStr);
    }

    std::string XLTableColumn::totalsRowLabel() const { return m_node.attribute("totalsRowLabel").value(); }
    void        XLTableColumn::setTotalsRowLabel(std::string_view label)
    {
        if (label.empty())
            m_node.remove_attribute("totalsRowLabel");
        else
            setAttr(m_node, "totalsRowLabel", std::string(label));
    }

    const std::vector<std::string_view> ColumnNodeOrder = {"calculatedColumnFormula", "totalsRowFormula", "xmlColumnPr", "extLst"};

    std::string XLTableColumn::totalsRowFormula() const
    {
        auto formulaNode = m_node.child("totalsRowFormula");
        return formulaNode.empty() ? "" : formulaNode.text().get();
    }

    void XLTableColumn::setTotalsRowFormula(std::string_view formula)
    {
        if (formula.empty()) {
            auto node = m_node.child("totalsRowFormula");
            if (!node.empty()) m_node.remove_child(node);
        }
        else {
            auto node = ensureChild(m_node, "totalsRowFormula", ColumnNodeOrder);
            node.text().set(std::string(formula).c_str());
        }
    }

    std::string XLTableColumn::calculatedColumnFormula() const { return ""; }
    void        XLTableColumn::setCalculatedColumnFormula(std::string_view /*formula*/) {}

}    // namespace OpenXLSX

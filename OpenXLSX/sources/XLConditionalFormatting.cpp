#include "XLConditionalFormatting.hpp"
#include "XLException.hpp"
#include "XLUtilities.hpp"
#include "XLWorksheet.hpp"
#include <algorithm>
#include <fmt/format.h>
#include <limits>
#include <map>
#include <string>

using namespace OpenXLSX;

XLConditionalFormat::XLConditionalFormat() : m_conditionalFormattingNode(XMLNode()) {}
XLConditionalFormat::XLConditionalFormat(const XMLNode& node) : m_conditionalFormattingNode(node) {}
XLConditionalFormat::XLConditionalFormat(const XLConditionalFormat& other) : m_conditionalFormattingNode(other.m_conditionalFormattingNode)
{}
XLConditionalFormat::~XLConditionalFormat() = default;

XLConditionalFormat& XLConditionalFormat::operator=(const XLConditionalFormat& other)
{
    if (&other != this) m_conditionalFormattingNode = other.m_conditionalFormattingNode;
    return *this;
}

bool        XLConditionalFormat::empty() const { return m_conditionalFormattingNode.empty(); }
std::string XLConditionalFormat::sqref() const { return m_conditionalFormattingNode.attribute("sqref").value(); }
XLCfRules   XLConditionalFormat::cfRules() const { return XLCfRules(m_conditionalFormattingNode); }
bool        XLConditionalFormat::setSqref(std::string newSqref)
{ return setAttr(m_conditionalFormattingNode, "sqref", newSqref).empty() == false; }
bool XLConditionalFormat::setExtLst(XLUnsupportedElement const& newExtLst)
{
    OpenXLSX::ignore(newExtLst);
    return false;
}

std::string XLConditionalFormat::summary() const { return fmt::format("sqref is {}, cfRules: {}", sqref(), cfRules().summary()); }

XLConditionalFormats::XLConditionalFormats() : m_sheetNode(XMLNode()), m_nodeOrder(&XLWorksheetNodeOrder) {}
XLConditionalFormats::XLConditionalFormats(const XMLNode& sheet) : m_sheetNode(sheet), m_nodeOrder(&XLWorksheetNodeOrder) {}
XLConditionalFormats::~XLConditionalFormats() {}
XLConditionalFormats::XLConditionalFormats(const XLConditionalFormats& other)
    : m_sheetNode(other.m_sheetNode),
      m_nodeOrder(other.m_nodeOrder)
{}
XLConditionalFormats::XLConditionalFormats(XLConditionalFormats&& other) noexcept
    : m_sheetNode(std::move(other.m_sheetNode)),
      m_nodeOrder(other.m_nodeOrder)
{}

XLConditionalFormats& XLConditionalFormats::operator=(const XLConditionalFormats& other)
{
    if (&other != this) {
        m_sheetNode = other.m_sheetNode;
        m_nodeOrder = other.m_nodeOrder;
    }
    return *this;
}

bool XLConditionalFormats::empty() const { return m_sheetNode.empty(); }

size_t XLConditionalFormats::count() const
{
    XMLNode node = m_sheetNode.first_child_of_type(node_element);
    while (not node.empty() and std::string(node.name()) != "conditionalFormatting") node = node.next_sibling_of_type(node_element);
    size_t count = 0;
    while (not node.empty() and std::string(node.name()) == "conditionalFormatting") {
        ++count;
        node = node.next_sibling_of_type(node_element);
    }
    return count;
}

XLConditionalFormat XLConditionalFormats::conditionalFormatByIndex(size_t index) const
{
    XMLNode node = m_sheetNode.first_child_of_type(node_element);
    while (not node.empty() and std::string(node.name()) != "conditionalFormatting") node = node.next_sibling_of_type(node_element);
    if (not node.empty()) {
        size_t count = 0;
        while (not node.empty() and std::string(node.name()) == "conditionalFormatting" and count < index) {
            ++count;
            node = node.next_sibling_of_type(node_element);
        }
        if (count == index and std::string(node.name()) == "conditionalFormatting") return XLConditionalFormat(node);
    }
    using namespace std::literals::string_literals;
    throw XLException("XLConditionalFormats::"s + __func__ + ": conditional format with index "s + std::to_string(index) +
                      " does not exist");
}

size_t XLConditionalFormats::create(XLConditionalFormat copyFrom, std::string conditionalFormattingPrefix)
{
    size_t  index = count();
    XMLNode newNode{};
    if (index == 0)
        newNode = ensureChild(m_sheetNode, "conditionalFormatting", *m_nodeOrder);
    else {
        XMLNode lastConditionalFormat = conditionalFormatByIndex(index - 1).m_conditionalFormattingNode;
        if (not lastConditionalFormat.empty()) newNode = m_sheetNode.insert_child_after("conditionalFormatting", lastConditionalFormat);
    }
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLConditionalFormats::"s + __func__ + ": failed to create a new conditional formatting entry");
    }

    if (!copyFrom.empty()) {
        for (XMLAttribute attr = copyFrom.m_conditionalFormattingNode.first_attribute(); attr; attr = attr.next_attribute()) {
            newNode.append_attribute(attr.name()).set_value(attr.value());
        }
        for (XMLNode child = copyFrom.m_conditionalFormattingNode.first_child(); child; child = child.next_sibling()) {
            newNode.append_copy(child);
        }
    }

    m_sheetNode.insert_child_before(node_pcdata, newNode).set_value(conditionalFormattingPrefix.c_str());
    return index;
}

std::string XLConditionalFormats::summary() const
{
    size_t conditionalFormatsCount = count();
    if (conditionalFormatsCount == 0) return "(no conditionalFormatting entries)";
    std::string result = "";
    for (size_t idx = 0; idx < conditionalFormatsCount; ++idx) {
        result += fmt::format("conditionalFormatting[{}] {}", idx, conditionalFormatByIndex(idx).summary());
        if (idx + 1 < conditionalFormatsCount) result += ", ";
    }
    return result;
}

// ----- Helper Builder Functions for XLCfRule -----

namespace OpenXLSX
{

    XLCfRule XLColorScaleRule(const XLColor& minColor, const XLColor& maxColor)
    {
        XLCfRule rule;
        rule.setType(XLCfType::ColorScale);
        XLCfColorScale colorScale;
        colorScale.addValue(XLCfvoType::Min, "", minColor);
        colorScale.addValue(XLCfvoType::Max, "", maxColor);
        rule.setColorScale(colorScale);
        return rule;
    }

    XLCfRule XLColorScaleRule(const XLColor& minColor, const XLColor& midColor, const XLColor& maxColor)
    {
        XLCfRule rule;
        rule.setType(XLCfType::ColorScale);
        XLCfColorScale colorScale;
        colorScale.addValue(XLCfvoType::Min, "", minColor);
        colorScale.addValue(XLCfvoType::Percentile, "50", midColor);
        colorScale.addValue(XLCfvoType::Max, "", maxColor);
        rule.setColorScale(colorScale);
        return rule;
    }

    XLCfRule XLDataBarRule(const XLColor& color, bool showValue)
    {
        XLCfRule rule;
        rule.setType(XLCfType::DataBar);
        XLCfDataBar dataBar;
        dataBar.setMin(XLCfvoType::Min, "");
        dataBar.setMax(XLCfvoType::Max, "");
        dataBar.setColor(color);
        dataBar.setShowValue(showValue);
        rule.setDataBar(dataBar);
        return rule;
    }

    XLCfRule XLCellIsRule(XLCfOperator op, const std::string& value)
    {
        XLCfRule rule;
        rule.setType(XLCfType::CellIs);
        rule.setOperator(op);
        rule.addFormula(value);
        return rule;
    }

    XLCfRule XLCellIsRule(const std::string& op, const std::string& value)
    {
        XLCfOperator cfOp = XLCfOperator::Equal;
        if (op == ">")
            cfOp = XLCfOperator::GreaterThan;
        else if (op == ">=")
            cfOp = XLCfOperator::GreaterThanOrEqual;
        else if (op == "<")
            cfOp = XLCfOperator::LessThan;
        else if (op == "<=")
            cfOp = XLCfOperator::LessThanOrEqual;
        else if (op == "==" || op == "=")
            cfOp = XLCfOperator::Equal;
        else if (op == "!=" || op == "<>")
            cfOp = XLCfOperator::NotEqual;
        else
            cfOp = XLCfOperatorFromString(op);

        return XLCellIsRule(cfOp, value);
    }

    XLCfRule XLFormulaRule(const std::string& formula)
    {
        XLCfRule rule;
        rule.setType(XLCfType::Expression);
        rule.addFormula(formula);
        return rule;
    }

    // Advanced builders
    XLCfRule XLIconSetRule(const std::string& iconSetName, bool showValue, bool reverse)
    {
        XLCfRule rule;
        rule.setType(XLCfType::IconSet);
        XLCfIconSet iconSet;
        iconSet.setIconSet(iconSetName);

        if (!showValue) iconSet.setShowValue(false);
        if (reverse) iconSet.setReverse(true);

        // Determine number of icons by looking at the first character of the iconSetName (e.g., "3Arrows", "4TrafficLights", "5Rating")
        int numIcons = 3;    // default
        if (!iconSetName.empty() && std::isdigit(iconSetName[0])) { numIcons = iconSetName[0] - '0'; }

        // Generate evenly spaced percentile thresholds (e.g., for 3: 0, 33, 67)
        for (int i = 0; i < numIcons; ++i) {
            int percent = (i * 100) / numIcons;
            iconSet.addValue(XLCfvoType::Percent, std::to_string(percent));
        }

        rule.setIconSet(iconSet);
        return rule;
    }

    XLCfRule XLTop10Rule(uint16_t rank, bool percent, bool bottom)
    {
        XLCfRule rule;
        rule.setType(XLCfType::Top10);
        rule.setRank(rank);
        if (percent) rule.setPercent(true);
        if (bottom) rule.setBottom(true);
        return rule;
    }

    XLCfRule XLAboveAverageRule(bool aboveAverage, bool equalAverage, int16_t stdDev)
    {
        XLCfRule rule;
        rule.setType(XLCfType::AboveAverage);
        if (!aboveAverage) rule.setAboveAverage(false);
        if (equalAverage) rule.setEqualAverage(true);
        if (stdDev > 0) rule.setStdDev(stdDev);
        return rule;
    }

    XLCfRule XLDuplicateValuesRule(bool unique)
    {
        XLCfRule rule;
        rule.setType(unique ? XLCfType::UniqueValues : XLCfType::DuplicateValues);
        return rule;
    }

    XLCfRule XLContainsTextRule(const std::string& text)
    {
        XLCfRule rule;
        rule.setType(XLCfType::ContainsText);
        rule.setText(text);
        rule.setOperator(XLCfOperator::ContainsText);
        return rule;
    }

    XLCfRule XLNotContainsTextRule(const std::string& text)
    {
        XLCfRule rule;
        rule.setType(XLCfType::NotContainsText);
        rule.setText(text);
        rule.setOperator(XLCfOperator::NotContains);
        return rule;
    }

    XLCfRule XLContainsBlanksRule()
    {
        XLCfRule rule;
        rule.setType(XLCfType::ContainsBlanks);
        return rule;
    }

    XLCfRule XLNotContainsBlanksRule()
    {
        XLCfRule rule;
        rule.setType(XLCfType::NotContainsBlanks);
        return rule;
    }

    XLCfRule XLContainsErrorsRule()
    {
        XLCfRule rule;
        rule.setType(XLCfType::ContainsErrors);
        return rule;
    }

    XLCfRule XLNotContainsErrorsRule()
    {
        XLCfRule rule;
        rule.setType(XLCfType::NotContainsErrors);
        return rule;
    }

}    // namespace OpenXLSX

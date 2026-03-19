#include <algorithm>
#include <map>
#include <limits>
#include <fmt/format.h>
#include "XLConditionalFormatting.hpp"
#include "XLWorksheet.hpp"
#include "XLUtilities.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{
    /**
     * @brief get the correct XLCfType from the OOXML cfRule type attribute string
     * @param typeString the string as used in the OOXML
     * @return the corresponding XLCfType enum value
     */
    XLCfType XLCfTypeFromString(std::string const& typeString)
    {
        if (typeString == "expression") return XLCfType::Expression;
        if (typeString == "cellIs") return XLCfType::CellIs;
        if (typeString == "colorScale") return XLCfType::ColorScale;
        if (typeString == "dataBar") return XLCfType::DataBar;
        if (typeString == "iconSet") return XLCfType::IconSet;
        if (typeString == "top10") return XLCfType::Top10;
        if (typeString == "uniqueValues") return XLCfType::UniqueValues;
        if (typeString == "duplicateValues") return XLCfType::DuplicateValues;
        if (typeString == "containsText") return XLCfType::ContainsText;
        if (typeString == "notContainsText") return XLCfType::NotContainsText;
        if (typeString == "beginsWith") return XLCfType::BeginsWith;
        if (typeString == "endsWith") return XLCfType::EndsWith;
        if (typeString == "containsBlanks") return XLCfType::ContainsBlanks;
        if (typeString == "notContainsBlanks") return XLCfType::NotContainsBlanks;
        if (typeString == "containsErrors") return XLCfType::ContainsErrors;
        if (typeString == "notContainsErrors") return XLCfType::NotContainsErrors;
        if (typeString == "timePeriod") return XLCfType::TimePeriod;
        if (typeString == "aboveAverage") return XLCfType::AboveAverage;
        return XLCfType::Invalid;
    }

    /**
     * @brief inverse of XLCfTypeFromString
     * @param cfType the type for which to get the OOXML string
     */
    std::string XLCfTypeToString(XLCfType cfType)
    {
        switch (cfType) {
            case XLCfType::Expression:
                return "expression";
            case XLCfType::CellIs:
                return "cellIs";
            case XLCfType::ColorScale:
                return "colorScale";
            case XLCfType::DataBar:
                return "dataBar";
            case XLCfType::IconSet:
                return "iconSet";
            case XLCfType::Top10:
                return "top10";
            case XLCfType::UniqueValues:
                return "uniqueValues";
            case XLCfType::DuplicateValues:
                return "duplicateValues";
            case XLCfType::ContainsText:
                return "containsText";
            case XLCfType::NotContainsText:
                return "notContainsText";
            case XLCfType::BeginsWith:
                return "beginsWith";
            case XLCfType::EndsWith:
                return "endsWith";
            case XLCfType::ContainsBlanks:
                return "containsBlanks";
            case XLCfType::NotContainsBlanks:
                return "notContainsBlanks";
            case XLCfType::ContainsErrors:
                return "containsErrors";
            case XLCfType::NotContainsErrors:
                return "notContainsErrors";
            case XLCfType::TimePeriod:
                return "timePeriod";
            case XLCfType::AboveAverage:
                return "aboveAverage";
            case XLCfType::Invalid:
                [[fallthrough]];
            default:
                return "(invalid)";
        }
    }

    /**
     * @brief get the correct XLCfOperator from the OOXML cfRule operator attribute string
     * @param operatorString the string as used in the OOXML
     * @return the corresponding XLCfOperator enum value
     */
    XLCfOperator XLCfOperatorFromString(std::string const& operatorString)
    {
        if (operatorString == "lessThan") return XLCfOperator::LessThan;
        if (operatorString == "lessThanOrEqual") return XLCfOperator::LessThanOrEqual;
        if (operatorString == "equal") return XLCfOperator::Equal;
        if (operatorString == "notEqual") return XLCfOperator::NotEqual;
        if (operatorString == "greaterThanOrEqual") return XLCfOperator::GreaterThanOrEqual;
        if (operatorString == "greaterThan") return XLCfOperator::GreaterThan;
        if (operatorString == "between") return XLCfOperator::Between;
        if (operatorString == "notBetween") return XLCfOperator::NotBetween;
        if (operatorString == "containsText") return XLCfOperator::ContainsText;
        if (operatorString == "notContains") return XLCfOperator::NotContains;
        if (operatorString == "beginsWith") return XLCfOperator::BeginsWith;
        if (operatorString == "endsWith") return XLCfOperator::EndsWith;
        return XLCfOperator::Invalid;
    }

    /**
     * @brief inverse of XLCfOperatorFromString
     * @param cfOperator the XLCfOperator for which to get the OOXML string
     */
    std::string XLCfOperatorToString(XLCfOperator cfOperator)
    {
        switch (cfOperator) {
            case XLCfOperator::LessThan:
                return "lessThan";
            case XLCfOperator::LessThanOrEqual:
                return "lessThanOrEqual";
            case XLCfOperator::Equal:
                return "equal";
            case XLCfOperator::NotEqual:
                return "notEqual";
            case XLCfOperator::GreaterThanOrEqual:
                return "greaterThanOrEqual";
            case XLCfOperator::GreaterThan:
                return "greaterThan";
            case XLCfOperator::Between:
                return "between";
            case XLCfOperator::NotBetween:
                return "notBetween";
            case XLCfOperator::ContainsText:
                return "containsText";
            case XLCfOperator::NotContains:
                return "notContains";
            case XLCfOperator::BeginsWith:
                return "beginsWith";
            case XLCfOperator::EndsWith:
                return "endsWith";
            case XLCfOperator::Invalid:
                [[fallthrough]];
            default:
                return "(invalid)";
        }
    }

    /**
     * @brief get the correct XLCfTimePeriod from the OOXML cfRule timePeriod attribute string
     * @param timePeriodString the string as used in the OOXML
     * @return the corresponding XLCfTimePeriod enum value
     */
    XLCfTimePeriod XLCfTimePeriodFromString(std::string const& timePeriodString)
    {
        if (timePeriodString == "today") return XLCfTimePeriod::Today;
        if (timePeriodString == "yesterday") return XLCfTimePeriod::Yesterday;
        if (timePeriodString == "tomorrow") return XLCfTimePeriod::Tomorrow;
        if (timePeriodString == "last7Days") return XLCfTimePeriod::Last7Days;
        if (timePeriodString == "thisMonth") return XLCfTimePeriod::ThisMonth;
        if (timePeriodString == "lastMonth") return XLCfTimePeriod::LastMonth;
        if (timePeriodString == "nextMonth") return XLCfTimePeriod::NextMonth;
        if (timePeriodString == "thisWeek") return XLCfTimePeriod::ThisWeek;
        if (timePeriodString == "lastWeek") return XLCfTimePeriod::LastWeek;
        if (timePeriodString == "nextWeek") return XLCfTimePeriod::NextWeek;
        return XLCfTimePeriod::Invalid;
    }

    /**
     * @brief inverse of XLCfTimePeriodFromString
     * @param cfTimePeriod the XLCfTimePeriod for which to get the OOXML string
     */
    std::string XLCfTimePeriodToString(XLCfTimePeriod cfTimePeriod)
    {
        switch (cfTimePeriod) {
            case XLCfTimePeriod::Today:
                return "today";
            case XLCfTimePeriod::Yesterday:
                return "yesterday";
            case XLCfTimePeriod::Tomorrow:
                return "tomorrow";
            case XLCfTimePeriod::Last7Days:
                return "last7Days";
            case XLCfTimePeriod::ThisMonth:
                return "thisMonth";
            case XLCfTimePeriod::LastMonth:
                return "lastMonth";
            case XLCfTimePeriod::NextMonth:
                return "nextMonth";
            case XLCfTimePeriod::ThisWeek:
                return "thisWeek";
            case XLCfTimePeriod::LastWeek:
                return "lastWeek";
            case XLCfTimePeriod::NextWeek:
                return "nextWeek";
            case XLCfTimePeriod::Invalid:
                [[fallthrough]];
            default:
                return "(invalid)";
        }
    }
}

XLCfRule::XLCfRule() : m_cfRuleNode(XMLNode()) {}
XLCfRule::XLCfRule(const XMLNode& node) : m_cfRuleNode(node) {}
XLCfRule::XLCfRule(const XLCfRule& other) : m_cfRuleNode(other.m_cfRuleNode) {}
XLCfRule::~XLCfRule() = default;

XLCfRule& XLCfRule::operator=(const XLCfRule& other)
{
    if (&other != this) m_cfRuleNode = other.m_cfRuleNode;
    return *this;
}

bool XLCfRule::empty() const { return m_cfRuleNode.empty(); }

std::string XLCfRule::formula() const
{
    return m_cfRuleNode.child("formula").first_child_of_type(pugi::node_pcdata).value();
}

XLUnsupportedElement XLCfRule::colorScale() const { return XLUnsupportedElement(); }
XLUnsupportedElement XLCfRule::dataBar() const { return XLUnsupportedElement(); }
XLUnsupportedElement XLCfRule::iconSet() const { return XLUnsupportedElement(); }
XLUnsupportedElement XLCfRule::extLst() const { return XLUnsupportedElement{}; }

XLCfType     XLCfRule::type() const { return XLCfTypeFromString(m_cfRuleNode.attribute("type").value()); }
XLStyleIndex XLCfRule::dxfId() const { return m_cfRuleNode.attribute("dxfId").as_uint(XLInvalidStyleIndex); }
uint16_t     XLCfRule::priority() const { return static_cast<uint16_t>(m_cfRuleNode.attribute("priority").as_uint(XLPriorityNotSet)); }
bool         XLCfRule::stopIfTrue() const { return m_cfRuleNode.attribute("stopIfTrue").as_bool(false); }
bool         XLCfRule::aboveAverage() const { return m_cfRuleNode.attribute("aboveAverage").as_bool(false); }
bool         XLCfRule::percent() const { return m_cfRuleNode.attribute("percent").as_bool(false); }
bool         XLCfRule::bottom() const { return m_cfRuleNode.attribute("bottom").as_bool(false); }
XLCfOperator   XLCfRule::Operator() const { return XLCfOperatorFromString(m_cfRuleNode.attribute("operator").value()); }
std::string    XLCfRule::text() const { return m_cfRuleNode.attribute("text").value(); }
XLCfTimePeriod XLCfRule::timePeriod() const { return XLCfTimePeriodFromString(m_cfRuleNode.attribute("timePeriod").value()); }
uint16_t       XLCfRule::rank() const { return static_cast<uint16_t>(m_cfRuleNode.attribute("rank").as_uint()); }
int16_t        XLCfRule::stdDev() const { return static_cast<int16_t>(m_cfRuleNode.attribute("stdDev").as_int()); }
bool           XLCfRule::equalAverage() const { return m_cfRuleNode.attribute("equalAverage").as_bool(false); }

bool XLCfRule::setFormula(std::string const& newFormula)
{
    XMLNode formula = appendAndGetNode(m_cfRuleNode, "formula", m_nodeOrder);
    if (formula.empty()) return false;
    formula.remove_children();
    return formula.append_child(pugi::node_pcdata).set_value(newFormula.c_str());
}

bool XLCfRule::setColorScale(XLUnsupportedElement const& newColorScale) { OpenXLSX::ignore(newColorScale); return false; }
bool XLCfRule::setDataBar(XLUnsupportedElement const& newDataBar) { OpenXLSX::ignore(newDataBar); return false; }
bool XLCfRule::setIconSet(XLUnsupportedElement const& newIconSet) { OpenXLSX::ignore(newIconSet); return false; }
bool XLCfRule::setExtLst(XLUnsupportedElement const& newExtLst) { OpenXLSX::ignore(newExtLst); return false; }

bool XLCfRule::setType(XLCfType newType) { return appendAndSetAttribute(m_cfRuleNode, "type", XLCfTypeToString(newType)).empty() == false; }
bool XLCfRule::setDxfId(XLStyleIndex newDxfId) { return appendAndSetAttribute(m_cfRuleNode, "dxfId", std::to_string(newDxfId)).empty() == false; }
bool XLCfRule::setPriority(uint16_t newPriority) { return appendAndSetAttribute(m_cfRuleNode, "priority", std::to_string(newPriority)).empty() == false; }
bool XLCfRule::setStopIfTrue(bool set) { return appendAndSetAttribute(m_cfRuleNode, "stopIfTrue", (set ? "true" : "false")).empty() == false; }
bool XLCfRule::setAboveAverage(bool set) { return appendAndSetAttribute(m_cfRuleNode, "aboveAverage", (set ? "true" : "false")).empty() == false; }
bool XLCfRule::setPercent(bool set) { return appendAndSetAttribute(m_cfRuleNode, "percent", (set ? "true" : "false")).empty() == false; }
bool XLCfRule::setBottom(bool set) { return appendAndSetAttribute(m_cfRuleNode, "bottom", (set ? "true" : "false")).empty() == false; }
bool XLCfRule::setOperator(XLCfOperator newOperator) { return appendAndSetAttribute(m_cfRuleNode, "operator", XLCfOperatorToString(newOperator)).empty() == false; }
bool XLCfRule::setText(std::string const& newText) { return appendAndSetAttribute(m_cfRuleNode, "text", newText.c_str()).empty() == false; }
bool XLCfRule::setTimePeriod(XLCfTimePeriod newTimePeriod) { return appendAndSetAttribute(m_cfRuleNode, "timePeriod", XLCfTimePeriodToString(newTimePeriod)).empty() == false; }
bool XLCfRule::setRank(uint16_t newRank) { return appendAndSetAttribute(m_cfRuleNode, "rank", std::to_string(newRank)).empty() == false; }
bool XLCfRule::setStdDev(int16_t newStdDev) { return appendAndSetAttribute(m_cfRuleNode, "stdDev", std::to_string(newStdDev)).empty() == false; }
bool XLCfRule::setEqualAverage(bool set) { return appendAndSetAttribute(m_cfRuleNode, "equalAverage", (set ? "true" : "false")).empty() == false; }

std::string XLCfRule::summary() const
{
    XLStyleIndex dxfIndex = dxfId();
    return fmt::format("formula is {}, type is {}, dxfId is {}, priority is {}, stopIfTrue: {}, aboveAverage: {}, percent: {}, bottom: {}, operator is {}, text is \"{}\", timePeriod is {}, rank is {}, stdDev is {}, equalAverage: {}",
                       formula(),
                       XLCfTypeToString(type()),
                       dxfIndex == XLInvalidStyleIndex ? "(invalid)" : std::to_string(dxfId()),
                       priority(),
                       stopIfTrue() ? "true" : "false",
                       aboveAverage() ? "true" : "false",
                       percent() ? "true" : "false",
                       bottom() ? "true" : "false",
                       XLCfOperatorToString(Operator()),
                       text(),
                       XLCfTimePeriodToString(timePeriod()),
                       rank(),
                       stdDev(),
                       equalAverage() ? "true" : "false");
}

XLCfRules::XLCfRules() : m_conditionalFormattingNode(XMLNode()) {}
XLCfRules::XLCfRules(const XMLNode& node) : m_conditionalFormattingNode(node) {}
XLCfRules::XLCfRules(const XLCfRules& other) : m_conditionalFormattingNode(other.m_conditionalFormattingNode) {}
XLCfRules::~XLCfRules() = default;

XLCfRules& XLCfRules::operator=(const XLCfRules& other)
{
    if (&other != this) m_conditionalFormattingNode = other.m_conditionalFormattingNode;
    return *this;
}

bool XLCfRules::empty() const { return m_conditionalFormattingNode.empty(); }

uint16_t XLCfRules::maxPriorityValue() const
{
    XMLNode node = m_conditionalFormattingNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(pugi::node_element);
    uint16_t maxPriority = XLPriorityNotSet;
    while (not node.empty() and std::string(node.name()) == "cfRule") {
        maxPriority = std::max(maxPriority, XLCfRule(node).priority());
        node        = node.next_sibling_of_type(pugi::node_element);
    }
    return maxPriority;
}

bool XLCfRules::setPriority(size_t cfRuleIndex, uint16_t newPriority)
{
    XLCfRule affectedRule = cfRuleByIndex(cfRuleIndex);
    if (newPriority == affectedRule.priority()) return true;

    XMLNode node = m_conditionalFormattingNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule")
        node = node.next_sibling_of_type(pugi::node_element);

    XMLNode node2           = node;
    bool    newPriorityFree = true;
    while (newPriorityFree and not node2.empty() and std::string(node2.name()) == "cfRule") {
        if (XLCfRule(node2).priority() == newPriority) newPriorityFree = false;
        node2 = node2.next_sibling_of_type(pugi::node_element);
    }

    if (!newPriorityFree) {
        size_t index = 0;
        while (not node.empty() and std::string(node.name()) == "cfRule") {
            if (index != cfRuleIndex) {
                XLCfRule rule(node);
                uint16_t prio = rule.priority();
                if (prio >= newPriority) rule.setPriority(prio + 1);
            }
            node = node.next_sibling_of_type(pugi::node_element);
            ++index;
        }
    }
    return affectedRule.setPriority(newPriority);
}

void XLCfRules::renumberPriorities(uint16_t increment)
{
    if (increment == 0) throw XLException("XLCfRules::renumberPriorities: increment must not be 0");
    std::multimap<uint16_t, XLCfRule> rules;
    XMLNode node = m_conditionalFormattingNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule")
        node = node.next_sibling_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) == "cfRule") {
        XLCfRule rule(node);
        rules.insert(std::pair(rule.priority(), std::move(rule)));
        node = node.next_sibling_of_type(pugi::node_element);
    }
    if (rules.size() * increment > std::numeric_limits<uint16_t>::max()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCfRules::renumberPriorities: amount of rules "s + std::to_string(rules.size()) + " with given increment "s + std::to_string(increment) + " exceeds max range of uint16_t"s);
    }
    uint16_t prio = 0;
    for (auto& [key, rule] : rules) {
        prio += increment;
        rule.setPriority(prio);
    }
    rules.clear();
}

size_t XLCfRules::count() const
{
    XMLNode node = m_conditionalFormattingNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(pugi::node_element);
    size_t count = 0;
    while (not node.empty() and std::string(node.name()) == "cfRule") {
        ++count;
        node = node.next_sibling_of_type(pugi::node_element);
    }
    return count;
}

XLCfRule XLCfRules::cfRuleByIndex(size_t index) const
{
    XMLNode node = m_conditionalFormattingNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(pugi::node_element);
    if (not node.empty()) {
        size_t count = 0;
        while (not node.empty() and std::string(node.name()) == "cfRule" and count < index) {
            ++count;
            node = node.next_sibling_of_type(pugi::node_element);
        }
        if (count == index and std::string(node.name()) == "cfRule") return XLCfRule(node);
    }
    using namespace std::literals::string_literals;
    throw XLException("XLCfRules::"s + __func__ + ": cfRule with index "s + std::to_string(index) + " does not exist");
}

size_t XLCfRules::create(XLCfRule copyFrom, std::string cfRulePrefix)
{
    OpenXLSX::ignore(copyFrom);
    uint16_t maxPrio = maxPriorityValue();
    if (maxPrio == std::numeric_limits<uint16_t>::max()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCfRules::"s + __func__ + ": can not create a new cfRule entry: no available priority value - please renumberPriorities or otherwise free up the highest value"s);
    }
    size_t  index = count();
    XMLNode newNode{};
    if (index == 0)
        newNode = appendAndGetNode(m_conditionalFormattingNode, "cfRule", m_nodeOrder);
    else {
        XMLNode lastCfRule = cfRuleByIndex(index - 1).m_cfRuleNode;
        if (not lastCfRule.empty()) newNode = m_conditionalFormattingNode.insert_child_after("cfRule", lastCfRule);
    }
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCfRules::"s + __func__ + ": failed to create a new cfRule entry");
    }
    m_conditionalFormattingNode.insert_child_before(pugi::node_pcdata, newNode).set_value(cfRulePrefix.c_str());
    cfRuleByIndex(index).setPriority(maxPrio + 1);
    return index;
}

std::string XLCfRules::summary() const
{
    size_t rulesCount = count();
    if (rulesCount == 0) return "(no cfRule entries)";
    std::string result = "";
    for (size_t idx = 0; idx < rulesCount; ++idx) {
        result += fmt::format("cfRule[{}] {}", idx, cfRuleByIndex(idx).summary());
        if (idx + 1 < rulesCount) result += ", ";
    }
    return result;
}

XLConditionalFormat::XLConditionalFormat() : m_conditionalFormattingNode(XMLNode()) {}
XLConditionalFormat::XLConditionalFormat(const XMLNode& node) : m_conditionalFormattingNode(node) {}
XLConditionalFormat::XLConditionalFormat(const XLConditionalFormat& other) : m_conditionalFormattingNode(other.m_conditionalFormattingNode) {}
XLConditionalFormat::~XLConditionalFormat() = default;

XLConditionalFormat& XLConditionalFormat::operator=(const XLConditionalFormat& other)
{
    if (&other != this) m_conditionalFormattingNode = other.m_conditionalFormattingNode;
    return *this;
}

bool XLConditionalFormat::empty() const { return m_conditionalFormattingNode.empty(); }
std::string XLConditionalFormat::sqref() const { return m_conditionalFormattingNode.attribute("sqref").value(); }
XLCfRules   XLConditionalFormat::cfRules() const { return XLCfRules(m_conditionalFormattingNode); }
bool XLConditionalFormat::setSqref(std::string newSqref) { return appendAndSetAttribute(m_conditionalFormattingNode, "sqref", newSqref).empty() == false; }
bool XLConditionalFormat::setExtLst(XLUnsupportedElement const& newExtLst) { OpenXLSX::ignore(newExtLst); return false; }

std::string XLConditionalFormat::summary() const
{
    return fmt::format("sqref is {}, cfRules: {}", sqref(), cfRules().summary());
}

XLConditionalFormats::XLConditionalFormats() : m_sheetNode(XMLNode()), m_nodeOrder(&XLWorksheetNodeOrder) {}
XLConditionalFormats::XLConditionalFormats(const XMLNode& sheet) : m_sheetNode(sheet), m_nodeOrder(&XLWorksheetNodeOrder) {}
XLConditionalFormats::~XLConditionalFormats() {}
XLConditionalFormats::XLConditionalFormats(const XLConditionalFormats& other) : m_sheetNode(other.m_sheetNode), m_nodeOrder(other.m_nodeOrder) {}
XLConditionalFormats::XLConditionalFormats(XLConditionalFormats&& other) noexcept : m_sheetNode(std::move(other.m_sheetNode)), m_nodeOrder(other.m_nodeOrder) {}

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
    XMLNode node = m_sheetNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "conditionalFormatting") node = node.next_sibling_of_type(pugi::node_element);
    size_t count = 0;
    while (not node.empty() and std::string(node.name()) == "conditionalFormatting") {
        ++count;
        node = node.next_sibling_of_type(pugi::node_element);
    }
    return count;
}

XLConditionalFormat XLConditionalFormats::conditionalFormatByIndex(size_t index) const
{
    XMLNode node = m_sheetNode.first_child_of_type(pugi::node_element);
    while (not node.empty() and std::string(node.name()) != "conditionalFormatting") node = node.next_sibling_of_type(pugi::node_element);
    if (not node.empty()) {
        size_t count = 0;
        while (not node.empty() and std::string(node.name()) == "conditionalFormatting" and count < index) {
            ++count;
            node = node.next_sibling_of_type(pugi::node_element);
        }
        if (count == index and std::string(node.name()) == "conditionalFormatting") return XLConditionalFormat(node);
    }
    using namespace std::literals::string_literals;
    throw XLException("XLConditionalFormats::"s + __func__ + ": conditional format with index "s + std::to_string(index) + " does not exist");
}

size_t XLConditionalFormats::create(XLConditionalFormat copyFrom, std::string conditionalFormattingPrefix)
{
    OpenXLSX::ignore(copyFrom);
    size_t  index = count();
    XMLNode newNode{};
    if (index == 0)
        newNode = appendAndGetNode(m_sheetNode, "conditionalFormatting", *m_nodeOrder);
    else {
        XMLNode lastConditionalFormat = conditionalFormatByIndex(index - 1).m_conditionalFormattingNode;
        if (not lastConditionalFormat.empty()) newNode = m_sheetNode.insert_child_after("conditionalFormatting", lastConditionalFormat);
    }
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLConditionalFormats::"s + __func__ + ": failed to create a new conditional formatting entry");
    }
    m_sheetNode.insert_child_before(pugi::node_pcdata, newNode).set_value(conditionalFormattingPrefix.c_str());
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

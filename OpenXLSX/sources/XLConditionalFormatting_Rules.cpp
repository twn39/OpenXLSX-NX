#include "XLConditionalFormatting.hpp"
#include "XLException.hpp"
#include "XLUtilities.hpp"
#include <algorithm>
#include <fmt/format.h>
#include <limits>
#include <map>
#include <string>

using namespace OpenXLSX;

XLCfRule::XLCfRule() : m_xmlDocument(std::make_unique<XMLDocument>()), m_cfRuleNode(XMLNode())
{ m_cfRuleNode = m_xmlDocument->append_child("cfRule"); }
XLCfRule::XLCfRule(const XMLNode& node) : m_cfRuleNode(node) {}
XLCfRule::XLCfRule(const XLCfRule& other) : m_cfRuleNode(other.m_cfRuleNode)
{
    if (other.m_xmlDocument) {
        m_xmlDocument = std::make_unique<XMLDocument>();
        m_xmlDocument->reset(*other.m_xmlDocument);
        m_cfRuleNode = m_xmlDocument->document_element();
    }
}
XLCfRule::XLCfRule(XLCfRule&& other) noexcept : m_xmlDocument(std::move(other.m_xmlDocument)), m_cfRuleNode(other.m_cfRuleNode) {}
XLCfRule::~XLCfRule() = default;

XLCfRule& XLCfRule::operator=(const XLCfRule& other)
{
    if (&other != this) {
        m_cfRuleNode = other.m_cfRuleNode;
        if (other.m_xmlDocument) {
            m_xmlDocument = std::make_unique<XMLDocument>();
            m_xmlDocument->reset(*other.m_xmlDocument);
            m_cfRuleNode = m_xmlDocument->document_element();
        }
        else {
            m_xmlDocument.reset();
        }
    }
    return *this;
}

XLCfRule& XLCfRule::operator=(XLCfRule&& other) noexcept
{
    if (this != &other) {
        m_xmlDocument = std::move(other.m_xmlDocument);
        m_cfRuleNode  = other.m_cfRuleNode;
    }
    return *this;
}

bool XLCfRule::empty() const { return m_cfRuleNode.empty(); }

std::string XLCfRule::formula() const { return m_cfRuleNode.child("formula").first_child_of_type(node_pcdata).value(); }

std::vector<std::string> XLCfRule::formulas() const
{
    std::vector<std::string> result;
    for (auto& node : m_cfRuleNode.children("formula")) { result.emplace_back(node.text().get()); }
    return result;
}

XLCfColorScale       XLCfRule::colorScale() const { return XLCfColorScale(m_cfRuleNode.child("colorScale")); }
XLCfDataBar          XLCfRule::dataBar() const { return XLCfDataBar(m_cfRuleNode.child("dataBar")); }
XLCfIconSet          XLCfRule::iconSet() const { return XLCfIconSet(m_cfRuleNode.child("iconSet")); }
XLUnsupportedElement XLCfRule::extLst() const { return XLUnsupportedElement{}; }

XLCfType       XLCfRule::type() const { return XLCfTypeFromString(m_cfRuleNode.attribute("type").value()); }
XLStyleIndex   XLCfRule::dxfId() const { return m_cfRuleNode.attribute("dxfId").as_uint(XLInvalidStyleIndex); }
uint16_t       XLCfRule::priority() const { return static_cast<uint16_t>(m_cfRuleNode.attribute("priority").as_uint(XLPriorityNotSet)); }
bool           XLCfRule::stopIfTrue() const { return m_cfRuleNode.attribute("stopIfTrue").as_bool(false); }
bool           XLCfRule::aboveAverage() const { return m_cfRuleNode.attribute("aboveAverage").as_bool(false); }
bool           XLCfRule::percent() const { return m_cfRuleNode.attribute("percent").as_bool(false); }
bool           XLCfRule::bottom() const { return m_cfRuleNode.attribute("bottom").as_bool(false); }
XLCfOperator   XLCfRule::Operator() const { return XLCfOperatorFromString(m_cfRuleNode.attribute("operator").value()); }
std::string    XLCfRule::text() const { return m_cfRuleNode.attribute("text").value(); }
XLCfTimePeriod XLCfRule::timePeriod() const { return XLCfTimePeriodFromString(m_cfRuleNode.attribute("timePeriod").value()); }
uint16_t       XLCfRule::rank() const { return static_cast<uint16_t>(m_cfRuleNode.attribute("rank").as_uint()); }
int16_t        XLCfRule::stdDev() const { return static_cast<int16_t>(m_cfRuleNode.attribute("stdDev").as_int()); }
bool           XLCfRule::equalAverage() const { return m_cfRuleNode.attribute("equalAverage").as_bool(false); }

XLCfRule& XLCfRule::setFormula(std::string const& newFormula)
{
    clearFormulas();
    addFormula(newFormula);
    return *this;
}

bool XLCfRule::addFormula(std::string const& newFormula)
{
    // Find where to insert: before the first non-formula child that is in m_nodeOrder
    XMLNode insertBefore = XMLNode();
    for (auto child : m_cfRuleNode.children()) {
        std::string_view name = child.name();
        if (name != "formula") {
            insertBefore = child;
            break;
        }
    }

    XMLNode newNode =
        insertBefore.empty() ? m_cfRuleNode.append_child("formula") : m_cfRuleNode.insert_child_before("formula", insertBefore);
    return newNode.append_child(node_pcdata).set_value(newFormula.c_str());
}

void XLCfRule::clearFormulas()
{
    while (!m_cfRuleNode.child("formula").empty()) { m_cfRuleNode.remove_child("formula"); }
}

bool XLCfRule::setColorScale(XLCfColorScale const& newColorScale)
{
    auto node = ensureChild(m_cfRuleNode, "colorScale", m_nodeOrder);
    node.remove_children();
    for (auto& cfvo : newColorScale.cfvos()) node.append_copy(cfvo.node());
    for (auto& color : newColorScale.colors()) {
        auto colorNode = node.append_child("color");
        colorNode.append_attribute("rgb").set_value(color.hex().c_str());
    }
    return true;
}

bool XLCfRule::setDataBar(XLCfDataBar const& newDataBar)
{
    auto node = ensureChild(m_cfRuleNode, "dataBar", m_nodeOrder);
    node.remove_children();
    node.append_copy(newDataBar.min().node());
    node.append_copy(newDataBar.max().node());
    auto colorNode = node.append_child("color");
    colorNode.append_attribute("rgb").set_value(newDataBar.color().hex().c_str());
    return true;
}

bool XLCfRule::setIconSet(XLCfIconSet const& newIconSet)
{
    auto node = ensureChild(m_cfRuleNode, "iconSet", m_nodeOrder);
    node.remove_attributes();
    node.remove_children();

    // Copy all attributes
    for (XMLAttribute attr = newIconSet.node().first_attribute(); attr; attr = attr.next_attribute()) {
        node.append_attribute(attr.name()).set_value(attr.value());
    }

    // Copy all children
    for (XMLNode child = newIconSet.node().first_child(); child; child = child.next_sibling()) { node.append_copy(child); }
    return true;
}

bool XLCfRule::setExtLst(XLUnsupportedElement const& newExtLst)
{
    OpenXLSX::ignore(newExtLst);
    return false;
}

XLCfRule& XLCfRule::setType(XLCfType newType)
{
    setAttr(m_cfRuleNode, "type", XLCfTypeToString(newType)).empty();
    return *this;
}
XLCfRule& XLCfRule::setDxfId(XLStyleIndex newDxfId)
{
    setAttr(m_cfRuleNode, "dxfId", std::to_string(newDxfId)).empty();
    return *this;
}
XLCfRule& XLCfRule::setPriority(uint16_t newPriority)
{
    setAttr(m_cfRuleNode, "priority", std::to_string(newPriority)).empty();
    return *this;
}
XLCfRule& XLCfRule::setStopIfTrue(bool set)
{
    setAttr(m_cfRuleNode, "stopIfTrue", (set ? "1" : "0")).empty();
    return *this;
}
XLCfRule& XLCfRule::setAboveAverage(bool set)
{
    setAttr(m_cfRuleNode, "aboveAverage", (set ? "1" : "0")).empty();
    return *this;
}
XLCfRule& XLCfRule::setPercent(bool set)
{
    setAttr(m_cfRuleNode, "percent", (set ? "1" : "0")).empty();
    return *this;
}
XLCfRule& XLCfRule::setBottom(bool set)
{
    setAttr(m_cfRuleNode, "bottom", (set ? "1" : "0")).empty();
    return *this;
}
XLCfRule& XLCfRule::setOperator(XLCfOperator newOperator)
{
    setAttr(m_cfRuleNode, "operator", XLCfOperatorToString(newOperator)).empty();
    return *this;
}
XLCfRule& XLCfRule::setText(std::string const& newText)
{
    setAttr(m_cfRuleNode, "text", newText.c_str()).empty();
    return *this;
}
XLCfRule& XLCfRule::setTimePeriod(XLCfTimePeriod newTimePeriod)
{
    setAttr(m_cfRuleNode, "timePeriod", XLCfTimePeriodToString(newTimePeriod)).empty();
    return *this;
}
XLCfRule& XLCfRule::setRank(uint16_t newRank)
{
    setAttr(m_cfRuleNode, "rank", std::to_string(newRank)).empty();
    return *this;
}
XLCfRule& XLCfRule::setStdDev(int16_t newStdDev)
{
    setAttr(m_cfRuleNode, "stdDev", std::to_string(newStdDev)).empty();
    return *this;
}
XLCfRule& XLCfRule::setEqualAverage(bool set)
{
    setAttr(m_cfRuleNode, "equalAverage", (set ? "1" : "0")).empty();
    return *this;
}

std::string XLCfRule::summary() const
{
    XLStyleIndex dxfIndex = dxfId();
    return fmt::format("formula is {}, type is {}, dxfId is {}, priority is {}, stopIfTrue: {}, aboveAverage: {}, percent: {}, bottom: {}, "
                       "operator is {}, text is \"{}\", timePeriod is {}, rank is {}, stdDev is {}, equalAverage: {}",
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
    XMLNode node = m_conditionalFormattingNode.first_child_of_type(node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(node_element);
    uint16_t maxPriority = XLPriorityNotSet;
    while (not node.empty() and std::string(node.name()) == "cfRule") {
        maxPriority = std::max(maxPriority, XLCfRule(node).priority());
        node        = node.next_sibling_of_type(node_element);
    }
    return maxPriority;
}

bool XLCfRules::setPriority(size_t cfRuleIndex, uint16_t newPriority)
{
    XLCfRule affectedRule = cfRuleByIndex(cfRuleIndex);
    if (newPriority == affectedRule.priority()) return true;

    XMLNode node = m_conditionalFormattingNode.first_child_of_type(node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(node_element);

    XMLNode node2           = node;
    bool    newPriorityFree = true;
    while (newPriorityFree and not node2.empty() and std::string(node2.name()) == "cfRule") {
        if (XLCfRule(node2).priority() == newPriority) newPriorityFree = false;
        node2 = node2.next_sibling_of_type(node_element);
    }

    if (!newPriorityFree) {
        size_t index = 0;
        while (not node.empty() and std::string(node.name()) == "cfRule") {
            if (index != cfRuleIndex) {
                XLCfRule rule(node);
                uint16_t prio = rule.priority();
                if (prio >= newPriority) rule.setPriority(prio + 1);
            }
            node = node.next_sibling_of_type(node_element);
            ++index;
        }
    }
    affectedRule.setPriority(newPriority);
    return true;
}

void XLCfRules::renumberPriorities(uint16_t increment)
{
    if (increment == 0) throw XLException("XLCfRules::renumberPriorities: increment must not be 0");
    std::multimap<uint16_t, XLCfRule> rules;
    XMLNode                           node = m_conditionalFormattingNode.first_child_of_type(node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(node_element);
    while (not node.empty() and std::string(node.name()) == "cfRule") {
        XLCfRule rule(node);
        rules.insert(std::pair(rule.priority(), std::move(rule)));
        node = node.next_sibling_of_type(node_element);
    }
    if (rules.size() * increment > std::numeric_limits<uint16_t>::max()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCfRules::renumberPriorities: amount of rules "s + std::to_string(rules.size()) + " with given increment "s +
                          std::to_string(increment) + " exceeds max range of uint16_t"s);
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
    XMLNode node = m_conditionalFormattingNode.first_child_of_type(node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(node_element);
    size_t count = 0;
    while (not node.empty() and std::string(node.name()) == "cfRule") {
        ++count;
        node = node.next_sibling_of_type(node_element);
    }
    return count;
}

XLCfRule XLCfRules::cfRuleByIndex(size_t index) const
{
    XMLNode node = m_conditionalFormattingNode.first_child_of_type(node_element);
    while (not node.empty() and std::string(node.name()) != "cfRule") node = node.next_sibling_of_type(node_element);
    if (not node.empty()) {
        size_t count = 0;
        while (not node.empty() and std::string(node.name()) == "cfRule" and count < index) {
            ++count;
            node = node.next_sibling_of_type(node_element);
        }
        if (count == index and std::string(node.name()) == "cfRule") return XLCfRule(node);
    }
    using namespace std::literals::string_literals;
    throw XLException("XLCfRules::"s + __func__ + ": cfRule with index "s + std::to_string(index) + " does not exist");
}

size_t XLCfRules::create(XLCfRule copyFrom, std::string cfRulePrefix)
{
    uint16_t explicitPrio = XLPriorityNotSet;
    if (!copyFrom.empty() && copyFrom.priority() != XLPriorityNotSet) { explicitPrio = copyFrom.priority(); }

    uint16_t maxPrio = maxPriorityValue();
    if (explicitPrio == XLPriorityNotSet && maxPrio == std::numeric_limits<uint16_t>::max()) {
        using namespace std::literals::string_literals;
        throw XLException(
            "XLCfRules::"s + __func__ +
            ": can not create a new cfRule entry: no available priority value - please renumberPriorities or otherwise free up the highest value"s);
    }

    size_t  index = count();
    XMLNode newNode{};
    if (index == 0)
        newNode = ensureChild(m_conditionalFormattingNode, "cfRule", m_nodeOrder);
    else {
        XMLNode lastCfRule = cfRuleByIndex(index - 1).m_cfRuleNode;
        if (not lastCfRule.empty()) newNode = m_conditionalFormattingNode.insert_child_after("cfRule", lastCfRule);
    }
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCfRules::"s + __func__ + ": failed to create a new cfRule entry");
    }

    if (!copyFrom.empty()) {
        for (XMLAttribute attr = copyFrom.node().first_attribute(); attr; attr = attr.next_attribute()) {
            newNode.append_attribute(attr.name()).set_value(attr.value());
        }
        for (XMLNode child = copyFrom.node().first_child(); child; child = child.next_sibling()) { newNode.append_copy(child); }
    }

    m_conditionalFormattingNode.insert_child_before(node_pcdata, newNode).set_value(cfRulePrefix.c_str());

    if (explicitPrio != XLPriorityNotSet) { cfRuleByIndex(index).setPriority(explicitPrio); }
    else {
        cfRuleByIndex(index).setPriority(maxPrio + 1);
    }

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


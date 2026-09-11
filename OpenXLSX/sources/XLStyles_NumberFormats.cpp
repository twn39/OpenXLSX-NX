// ===== External Includes ===== //
#include <cstdint>
#include <fmt/format.h>
#include <gsl/gsl>
#include <memory>
#include <pugixml.hpp>
#include <stdexcept>
#include <string>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLStyles.hpp"
#include "XLStyles_Internal.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

XLNumberFormat::XLNumberFormat() : m_numberFormatNode(std::make_unique<XMLNode>()) {}

XLNumberFormat::XLNumberFormat(const XMLNode& node) : m_numberFormatNode(std::make_unique<XMLNode>(node)) {}

XLNumberFormat::~XLNumberFormat() = default;

XLNumberFormat::XLNumberFormat(const XLNumberFormat& other) : m_numberFormatNode(std::make_unique<XMLNode>(*other.m_numberFormatNode)) {}

XLNumberFormat& XLNumberFormat::operator=(const XLNumberFormat& other)
{
    if (&other != this) *m_numberFormatNode = *other.m_numberFormatNode;
    return *this;
}

uint32_t XLNumberFormat::numberFormatId() const { return m_numberFormatNode->attribute("numFmtId").as_uint(XLInvalidUInt32); }

std::string XLNumberFormat::formatCode() const { return m_numberFormatNode->attribute("formatCode").value(); }

bool XLNumberFormat::setNumberFormatId(uint32_t newNumberFormatId)
{ return setAttr(*m_numberFormatNode, "numFmtId", std::to_string(newNumberFormatId)).empty() == false; }
bool XLNumberFormat::setFormatCode(std::string_view newFormatCode)
{ return setAttr(*m_numberFormatNode, "formatCode", std::string(newFormatCode).c_str()).empty() == false; }

std::string XLNumberFormat::summary() const { return fmt::format("numFmtId={}, formatCode={}", numberFormatId(), formatCode()); }

// ===== XLNumberFormats, parent of XLNumberFormat

XLNumberFormats::XLNumberFormats() : m_numberFormatsNode(std::make_unique<XMLNode>()) {}

XLNumberFormats::XLNumberFormats(const XMLNode& numberFormats) : m_numberFormatsNode(std::make_unique<XMLNode>(numberFormats))
{
    // initialize XLNumberFormat entries and m_numberFormats here
    XMLNode node = numberFormats.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string_view nodeName(node.name());
        if (nodeName == "numFmt")
            m_numberFormats.push_back(XLNumberFormat(node));
        else
            std::cerr << "WARNING: XLNumberFormats constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLNumberFormats::~XLNumberFormats() { m_numberFormats.clear(); }

XLNumberFormats::XLNumberFormats(const XLNumberFormats& other)
    : m_numberFormatsNode(std::make_unique<XMLNode>(*other.m_numberFormatsNode)),
      m_numberFormats(other.m_numberFormats)
{}

XLNumberFormats::XLNumberFormats(XLNumberFormats&& other)
    : m_numberFormatsNode(std::move(other.m_numberFormatsNode)),
      m_numberFormats(std::move(other.m_numberFormats))
{}

XLNumberFormats& XLNumberFormats::operator=(const XLNumberFormats& other)
{
    if (&other != this) {
        *m_numberFormatsNode = *other.m_numberFormatsNode;
        m_numberFormats.clear();
        m_numberFormats = other.m_numberFormats;
    }
    return *this;
}

size_t XLNumberFormats::count() const { return m_numberFormats.size(); }

XLNumberFormat XLNumberFormats::numberFormatByIndex(XLStyleIndex index) const
{
    Expects(index < m_numberFormats.size());
    return m_numberFormats.at(index);
}

std::optional<uint32_t> XLNumberFormats::standardNumberFormatId(std::string_view formatCode) noexcept
{
    if (formatCode == "General" || formatCode == "general") return 0;
    if (formatCode == "0") return 1;
    if (formatCode == "0.00") return 2;
    if (formatCode == "#,##0") return 3;
    if (formatCode == "#,##0.00") return 4;
    if (formatCode == "0%") return 9;
    if (formatCode == "0.00%") return 10;
    if (formatCode == "0.00E+00" || formatCode == "0.00e+00") return 11;
    if (formatCode == "# ?/?" || formatCode == "?/?") return 12;
    if (formatCode == "# ?\?/?\?" || formatCode == "?\?/?\?") return 13;
    if (formatCode == "mm-dd-yy" || formatCode == "m/d/yy" || formatCode == "m/d/yyyy" || formatCode == "yyyy-mm-dd") return 14;
    if (formatCode == "d-mmm-yy") return 15;
    if (formatCode == "d-mmm") return 16;
    if (formatCode == "mmm-yy") return 17;
    if (formatCode == "h:mm AM/PM") return 18;
    if (formatCode == "h:mm:ss AM/PM") return 19;
    if (formatCode == "h:mm") return 20;
    if (formatCode == "h:mm:ss") return 21;
    if (formatCode == "m/d/yy h:mm" || formatCode == "m/d/yyyy h:mm") return 22;
    if (formatCode == "#,##0 ;(#,##0)" || formatCode == "#,##0;(#,##0)") return 37;
    if (formatCode == "#,##0 ;[Red](#,##0)" || formatCode == "#,##0;[Red](#,##0)") return 38;
    if (formatCode == "#,##0.00;(#,##0.00)" || formatCode == "#,##0.00 ;(#,##0.00)") return 39;
    if (formatCode == "#,##0.00;[Red](#,##0.00)" || formatCode == "#,##0.00 ;[Red](#,##0.00)") return 40;
    if (formatCode == "mm:ss") return 45;
    if (formatCode == "[h]:mm:ss") return 46;
    if (formatCode == "mmss.0") return 47;
    if (formatCode == "##0.0E+0") return 48;
    if (formatCode == "@") return 49;
    return std::nullopt;
}

std::string_view XLNumberFormats::standardFormatCode(uint32_t id) noexcept
{
    switch (id) {
        case 0: return "General";
        case 1: return "0";
        case 2: return "0.00";
        case 3: return "#,##0";
        case 4: return "#,##0.00";
        case 9: return "0%";
        case 10: return "0.00%";
        case 11: return "0.00E+00";
        case 12: return "# ?/?";
        case 13: return "# ?\?/?\?";
        case 14: return "mm-dd-yy";
        case 15: return "d-mmm-yy";
        case 16: return "d-mmm";
        case 17: return "mmm-yy";
        case 18: return "h:mm AM/PM";
        case 19: return "h:mm:ss AM/PM";
        case 20: return "h:mm";
        case 21: return "h:mm:ss";
        case 22: return "m/d/yy h:mm";
        case 37: return "#,##0 ;(#,##0)";
        case 38: return "#,##0 ;[Red](#,##0)";
        case 39: return "#,##0.00;(#,##0.00)";
        case 40: return "#,##0.00;[Red](#,##0.00)";
        case 45: return "mm:ss";
        case 46: return "[h]:mm:ss";
        case 47: return "mmss.0";
        case 48: return "##0.0E+0";
        case 49: return "@";
        default: return "";
    }
}

XLNumberFormat XLNumberFormats::numberFormatById(uint32_t numberFormatId) const
{
    for (const auto& fmt : m_numberFormats)
        if (fmt.numberFormatId() == numberFormatId) return fmt;

    auto stdCode = standardFormatCode(numberFormatId);
    if (!stdCode.empty()) {
        static const auto s_stdDoc = []() {
            auto doc = std::make_unique<pugi::xml_document>();
            return doc;
        }();
        static std::unordered_map<uint32_t, XMLNode> s_nodes;
        static std::mutex s_mutex;
        std::lock_guard<std::mutex> lock(s_mutex);
        auto it = s_nodes.find(numberFormatId);
        if (it != s_nodes.end()) return XLNumberFormat(it->second);

        XMLNode n = s_stdDoc->append_child("numFmt");
        n.append_attribute("numFmtId").set_value(numberFormatId);
        n.append_attribute("formatCode").set_value(std::string(stdCode).c_str());
        s_nodes.emplace(numberFormatId, n);
        return XLNumberFormat(n);
    }

    using namespace std::literals::string_literals;
    throw XLException("XLNumberFormats::"s + __func__ + ": numberFormatId "s + std::to_string(numberFormatId) + " not found"s);
}

uint32_t XLNumberFormats::numberFormatIdFromIndex(XLStyleIndex index) const
{
    Expects(index < m_numberFormats.size());
    return m_numberFormats[index].numberFormatId();
}

uint32_t XLNumberFormats::getFreeNumberFormatId() const
{
    uint32_t maxId = 163;    // Excel reserved format IDs end at 163
    for (const auto& fmt : m_numberFormats) {
        if (fmt.numberFormatId() > maxId) { maxId = fmt.numberFormatId(); }
    }
    return maxId + 1;
}

uint32_t XLNumberFormats::createNumberFormat(std::string_view formatCode)
{
    auto stdId = standardNumberFormatId(formatCode);
    if (stdId.has_value()) {
        return *stdId;
    }

    // If exact format code already exists, just return its ID
    for (const auto& fmt : m_numberFormats) {
        if (fmt.formatCode() == formatCode) { return fmt.numberFormatId(); }
    }

    uint32_t newId = getFreeNumberFormatId();

    XMLNode newNode{};
    XMLNode lastStyle = m_numberFormatsNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty())
        newNode = m_numberFormatsNode->prepend_child("numFmt");
    else
        newNode = m_numberFormatsNode->insert_child_after("numFmt", lastStyle);

    newNode.append_attribute("numFmtId").set_value(newId);
    newNode.append_attribute("formatCode").set_value(std::string(formatCode).c_str());

    // Add to our tracked vector
    m_numberFormats.emplace_back(newNode);

    // Update count attribute on <numFmts>
    setAttr(*m_numberFormatsNode, "count", static_cast<unsigned long>(m_numberFormats.size()));

    return newId;
}

XLStyleIndex XLNumberFormats::create(XLNumberFormat copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();
    XMLNode      newNode{};

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_numberFormatsNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty())
        newNode = m_numberFormatsNode->prepend_child("numFmt");
    else
        newNode = m_numberFormatsNode->insert_child_after("numFmt", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLNumberFormats::"s + __func__ + ": failed to append a new numFmt node"s);
    }
    if (styleEntriesPrefix.length() > 0)
        m_numberFormatsNode->insert_child_before(pugi::node_pcdata, newNode).set_value(std::string(styleEntriesPrefix).c_str());

    XLNumberFormat newNumberFormat(newNode);
    if (copyFrom.m_numberFormatNode->empty()) {
        // ===== Create a number format with default values
        newNumberFormat.setNumberFormatId(0);
        newNumberFormat.setFormatCode("General");
    }
    else {
        copyXMLNode(newNode, *copyFrom.m_numberFormatNode);
        newNumberFormat.setNumberFormatId(getFreeNumberFormatId());    // Ensure copied formats get a fresh ID
    }

    m_numberFormats.push_back(newNumberFormat);
    setAttr(*m_numberFormatsNode, "count", std::to_string(m_numberFormats.size()));
    return index;
}

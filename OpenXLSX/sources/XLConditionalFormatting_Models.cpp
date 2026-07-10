#include "XLConditionalFormatting.hpp"
#include "XLException.hpp"
#include "XLUtilities.hpp"
#include <algorithm>
#include <fmt/format.h>
#include <limits>
#include <map>

using namespace OpenXLSX;

XLCfvo::XLCfvo() : m_xmlDocument(std::make_unique<XMLDocument>()) { m_cfvoNode = m_xmlDocument->append_child("cfvo"); }
XLCfvo::XLCfvo(const XMLNode& node) : m_cfvoNode(node) {}
XLCfvo::XLCfvo(const XLCfvo& other) : m_cfvoNode(other.m_cfvoNode)
{
    if (other.m_xmlDocument) {
        m_xmlDocument = std::make_unique<XMLDocument>();
        m_xmlDocument->reset(*other.m_xmlDocument);
        m_cfvoNode = m_xmlDocument->document_element();
    }
}
XLCfvo::XLCfvo(XLCfvo&& other) noexcept : m_xmlDocument(std::move(other.m_xmlDocument)), m_cfvoNode(other.m_cfvoNode) {}
XLCfvo::~XLCfvo() = default;
XLCfvo& XLCfvo::operator=(const XLCfvo& other)
{
    if (this != &other) {
        m_cfvoNode = other.m_cfvoNode;
        if (other.m_xmlDocument) {
            m_xmlDocument = std::make_unique<XMLDocument>();
            m_xmlDocument->reset(*other.m_xmlDocument);
            m_cfvoNode = m_xmlDocument->document_element();
        }
        else {
            m_xmlDocument.reset();
        }
    }
    return *this;
}
XLCfvo& XLCfvo::operator=(XLCfvo&& other) noexcept
{
    if (this != &other) {
        m_xmlDocument = std::move(other.m_xmlDocument);
        m_cfvoNode    = other.m_cfvoNode;
    }
    return *this;
}

XLCfvoType  XLCfvo::type() const { return XLCfvoTypeFromString(m_cfvoNode.attribute("type").value()); }
std::string XLCfvo::value() const { return m_cfvoNode.attribute("val").value(); }
bool        XLCfvo::gte() const { return m_cfvoNode.attribute("gte").as_bool(true); }
void        XLCfvo::setType(XLCfvoType type)
{
    setAttr(m_cfvoNode, "type", XLCfvoTypeToString(type).c_str());
}
void XLCfvo::setValue(const std::string& value)
{
    if (value.empty()) { m_cfvoNode.remove_attribute("val"); }
    else {
        setAttr(m_cfvoNode, "val", value.c_str());
    }
}
void XLCfvo::setGte(bool gte) { setAttrBool01(m_cfvoNode, "gte", gte); }

XLCfColorScale::XLCfColorScale() : m_xmlDocument(std::make_unique<XMLDocument>())
{ m_colorScaleNode = m_xmlDocument->append_child("colorScale"); }
XLCfColorScale::XLCfColorScale(const XMLNode& node) : m_colorScaleNode(node) {}
XLCfColorScale::XLCfColorScale(const XLCfColorScale& other) : m_colorScaleNode(other.m_colorScaleNode)
{
    if (other.m_xmlDocument) {
        m_xmlDocument = std::make_unique<XMLDocument>();
        m_xmlDocument->reset(*other.m_xmlDocument);
        m_colorScaleNode = m_xmlDocument->document_element();
    }
}
XLCfColorScale::XLCfColorScale(XLCfColorScale&& other) noexcept
    : m_xmlDocument(std::move(other.m_xmlDocument)),
      m_colorScaleNode(other.m_colorScaleNode)
{}
XLCfColorScale::~XLCfColorScale() = default;
XLCfColorScale& XLCfColorScale::operator=(const XLCfColorScale& other)
{
    if (this != &other) {
        m_colorScaleNode = other.m_colorScaleNode;
        if (other.m_xmlDocument) {
            m_xmlDocument = std::make_unique<XMLDocument>();
            m_xmlDocument->reset(*other.m_xmlDocument);
            m_colorScaleNode = m_xmlDocument->document_element();
        }
        else {
            m_xmlDocument.reset();
        }
    }
    return *this;
}
XLCfColorScale& XLCfColorScale::operator=(XLCfColorScale&& other) noexcept
{
    if (this != &other) {
        m_xmlDocument    = std::move(other.m_xmlDocument);
        m_colorScaleNode = other.m_colorScaleNode;
    }
    return *this;
}

std::vector<XLCfvo> XLCfColorScale::cfvos() const
{
    std::vector<XLCfvo> result;
    for (auto& node : m_colorScaleNode.children("cfvo")) result.emplace_back(node);
    return result;
}
std::vector<XLColor> XLCfColorScale::colors() const
{
    std::vector<XLColor> result;
    for (auto& node : m_colorScaleNode.children("color")) {
        auto attr = node.attribute("rgb");
        if (!attr.empty() && std::string_view(attr.value()).size() >= 6)
            result.emplace_back(attr.value());
        else
            result.emplace_back();
    }
    return result;
}
void XLCfColorScale::addValue(XLCfvoType type, const std::string& value, const XLColor& color)
{
    auto cfvo = m_colorScaleNode.append_child("cfvo");
    cfvo.append_attribute("type").set_value(XLCfvoTypeToString(type).c_str());
    if (!value.empty()) cfvo.append_attribute("val").set_value(value.c_str());

    auto colorNode = m_colorScaleNode.append_child("color");
    colorNode.append_attribute("rgb").set_value(color.hex().c_str());
}
void XLCfColorScale::clear() { m_colorScaleNode.remove_children(); }

XLCfDataBar::XLCfDataBar() : m_xmlDocument(std::make_unique<XMLDocument>()) { m_dataBarNode = m_xmlDocument->append_child("dataBar"); }
XLCfDataBar::XLCfDataBar(const XMLNode& node) : m_dataBarNode(node) {}
XLCfDataBar::XLCfDataBar(const XLCfDataBar& other) : m_dataBarNode(other.m_dataBarNode)
{
    if (other.m_xmlDocument) {
        m_xmlDocument = std::make_unique<XMLDocument>();
        m_xmlDocument->reset(*other.m_xmlDocument);
        m_dataBarNode = m_xmlDocument->document_element();
    }
}
XLCfDataBar::XLCfDataBar(XLCfDataBar&& other) noexcept : m_xmlDocument(std::move(other.m_xmlDocument)), m_dataBarNode(other.m_dataBarNode)
{}
XLCfDataBar::~XLCfDataBar() = default;
XLCfDataBar& XLCfDataBar::operator=(const XLCfDataBar& other)
{
    if (this != &other) {
        m_dataBarNode = other.m_dataBarNode;
        if (other.m_xmlDocument) {
            m_xmlDocument = std::make_unique<XMLDocument>();
            m_xmlDocument->reset(*other.m_xmlDocument);
            m_dataBarNode = m_xmlDocument->document_element();
        }
        else {
            m_xmlDocument.reset();
        }
    }
    return *this;
}
XLCfDataBar& XLCfDataBar::operator=(XLCfDataBar&& other) noexcept
{
    if (this != &other) {
        m_xmlDocument = std::move(other.m_xmlDocument);
        m_dataBarNode = other.m_dataBarNode;
    }
    return *this;
}

XLCfvo  XLCfDataBar::min() const { return XLCfvo(m_dataBarNode.child("cfvo")); }
XLCfvo  XLCfDataBar::max() const { return XLCfvo(m_dataBarNode.child("cfvo").next_sibling("cfvo")); }
XLColor XLCfDataBar::color() const
{
    auto attr = m_dataBarNode.child("color").attribute("rgb");
    return (!attr.empty() && std::string_view(attr.value()).size() >= 6) ? XLColor(attr.value()) : XLColor();
}
void XLCfDataBar::setMin(XLCfvoType type, const std::string& value)
{
    auto node = m_dataBarNode.child("cfvo");
    if (node.empty()) node = m_dataBarNode.prepend_child("cfvo");
    XLCfvo(node).setType(type);
    XLCfvo(node).setValue(value);
}
void XLCfDataBar::setMax(XLCfvoType type, const std::string& value)
{
    auto node = m_dataBarNode.child("cfvo").next_sibling("cfvo");
    if (node.empty()) node = m_dataBarNode.insert_child_after("cfvo", m_dataBarNode.child("cfvo"));
    XLCfvo(node).setType(type);
    XLCfvo(node).setValue(value);
}
void XLCfDataBar::setColor(const XLColor& color)
{
    setChildAttr(m_dataBarNode, "color", "rgb", color.hex().c_str());
}
bool XLCfDataBar::showValue() const { return m_dataBarNode.attribute("showValue").as_bool(true); }
void XLCfDataBar::setShowValue(bool show) { setAttrBool01(m_dataBarNode, "showValue", show); }

XMLNode XLCfDataBar::getExtNode() const
{
    auto extLst = m_dataBarNode.child("extLst");
    if (extLst) {
        for (auto ext = extLst.child("ext"); ext; ext = ext.next_sibling("ext")) {
            if (std::string(ext.attribute("uri").value()) == "{B469CE28-E4FAB-4e90-B891-B227B70C6CDA}") {
                return ext.child("x14:dataBar");
            }
        }
    }
    return XMLNode();
}

XMLNode XLCfDataBar::getOrCreateExtNode() const
{
    auto extLst = m_dataBarNode.child("extLst");
    if (!extLst) {
        extLst = m_dataBarNode.append_child("extLst");
    }
    XMLNode extNode;
    for (auto ext = extLst.child("ext"); ext; ext = ext.next_sibling("ext")) {
        if (std::string(ext.attribute("uri").value()) == "{B469CE28-E4FAB-4e90-B891-B227B70C6CDA}") {
            extNode = ext;
            break;
        }
    }
    if (!extNode) {
        extNode = extLst.append_child("ext");
        extNode.append_attribute("uri") = "{B469CE28-E4FAB-4e90-B891-B227B70C6CDA}";
        extNode.append_attribute("xmlns:x14") = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    }
    auto x14DataBar = extNode.child("x14:dataBar");
    if (!x14DataBar) {
        x14DataBar = extNode.append_child("x14:dataBar");
        x14DataBar.append_attribute("minLength") = "0";
        x14DataBar.append_attribute("maxLength") = "100";
    }
    return x14DataBar;
}

uint32_t XLCfDataBar::minLength() const
{
    auto extNode = getExtNode();
    return extNode ? extNode.attribute("minLength").as_uint(10) : 10;
}

void XLCfDataBar::setMinLength(uint32_t length)
{
    auto extNode = getOrCreateExtNode();
    setAttr(extNode, "minLength", length);
}

uint32_t XLCfDataBar::maxLength() const
{
    auto extNode = getExtNode();
    return extNode ? extNode.attribute("maxLength").as_uint(90) : 90;
}

void XLCfDataBar::setMaxLength(uint32_t length)
{
    auto extNode = getOrCreateExtNode();
    setAttr(extNode, "maxLength", length);
}

bool XLCfDataBar::border() const
{
    auto extNode = getExtNode();
    return extNode ? extNode.attribute("border").as_bool(false) : false;
}

void XLCfDataBar::setBorder(bool border)
{
    auto extNode = getOrCreateExtNode();
    setAttrBool01(extNode, "border", border);
}

bool XLCfDataBar::gradient() const
{
    auto extNode = getExtNode();
    return extNode ? extNode.attribute("gradient").as_bool(true) : true;
}

void XLCfDataBar::setGradient(bool gradient)
{
    auto extNode = getOrCreateExtNode();
    setAttrBool01(extNode, "gradient", gradient);
}

XLDataBarDirection XLCfDataBar::direction() const
{
    auto extNode = getExtNode();
    return extNode ? XLDataBarDirectionFromString(extNode.attribute("direction").value()) : XLDataBarDirection::Context;
}

void XLCfDataBar::setDirection(XLDataBarDirection direction)
{
    auto extNode = getOrCreateExtNode();
    setAttr(extNode, "direction", XLDataBarDirectionToString(direction).c_str());
}

XLColor XLCfDataBar::borderColor() const
{
    auto extNode = getExtNode();
    if (!extNode) return XLColor();
    auto attr = extNode.child("x14:borderColor").attribute("rgb");
    return (!attr.empty() && std::string_view(attr.value()).size() >= 6) ? XLColor(attr.value()) : XLColor();
}

void XLCfDataBar::setBorderColor(const XLColor& color)
{
    auto extNode = getOrCreateExtNode();
    setChildAttr(extNode, "x14:borderColor", "rgb", color.hex().c_str());
}

XLColor XLCfDataBar::negativeFillColor() const
{
    auto extNode = getExtNode();
    if (!extNode) return XLColor();
    auto attr = extNode.child("x14:negativeFillColor").attribute("rgb");
    return (!attr.empty() && std::string_view(attr.value()).size() >= 6) ? XLColor(attr.value()) : XLColor();
}

void XLCfDataBar::setNegativeFillColor(const XLColor& color)
{
    auto extNode = getOrCreateExtNode();
    setChildAttr(extNode, "x14:negativeFillColor", "rgb", color.hex().c_str());
}

XLColor XLCfDataBar::negativeBorderColor() const
{
    auto extNode = getExtNode();
    if (!extNode) return XLColor();
    auto attr = extNode.child("x14:negativeBorderColor").attribute("rgb");
    return (!attr.empty() && std::string_view(attr.value()).size() >= 6) ? XLColor(attr.value()) : XLColor();
}

void XLCfDataBar::setNegativeBorderColor(const XLColor& color)
{
    auto extNode = getOrCreateExtNode();
    setChildAttr(extNode, "x14:negativeBorderColor", "rgb", color.hex().c_str());
}

XLDataBarAxisPosition XLCfDataBar::axisPosition() const
{
    auto extNode = getExtNode();
    return extNode ? XLDataBarAxisPositionFromString(extNode.attribute("axisPosition").value()) : XLDataBarAxisPosition::Automatic;
}

void XLCfDataBar::setAxisPosition(XLDataBarAxisPosition position)
{
    auto extNode = getOrCreateExtNode();
    setAttr(extNode, "axisPosition", XLDataBarAxisPositionToString(position).c_str());
}

XLColor XLCfDataBar::axisColor() const
{
    auto extNode = getExtNode();
    if (!extNode) return XLColor();
    auto attr = extNode.child("x14:axisColor").attribute("rgb");
    return (!attr.empty() && std::string_view(attr.value()).size() >= 6) ? XLColor(attr.value()) : XLColor();
}

void XLCfDataBar::setAxisColor(const XLColor& color)
{
    auto extNode = getOrCreateExtNode();
    setChildAttr(extNode, "x14:axisColor", "rgb", color.hex().c_str());
}


XLCfIconSet::XLCfIconSet() : m_xmlDocument(std::make_unique<XMLDocument>()) { m_iconSetNode = m_xmlDocument->append_child("iconSet"); }
XLCfIconSet::XLCfIconSet(const XMLNode& node) : m_iconSetNode(node) {}
XLCfIconSet::XLCfIconSet(const XLCfIconSet& other) : m_iconSetNode(other.m_iconSetNode)
{
    if (other.m_xmlDocument) {
        m_xmlDocument = std::make_unique<XMLDocument>();
        m_xmlDocument->reset(*other.m_xmlDocument);
        m_iconSetNode = m_xmlDocument->document_element();
    }
}
XLCfIconSet::XLCfIconSet(XLCfIconSet&& other) noexcept : m_xmlDocument(std::move(other.m_xmlDocument)), m_iconSetNode(other.m_iconSetNode)
{}
XLCfIconSet::~XLCfIconSet() = default;
XLCfIconSet& XLCfIconSet::operator=(const XLCfIconSet& other)
{
    if (this != &other) {
        m_iconSetNode = other.m_iconSetNode;
        if (other.m_xmlDocument) {
            m_xmlDocument = std::make_unique<XMLDocument>();
            m_xmlDocument->reset(*other.m_xmlDocument);
            m_iconSetNode = m_xmlDocument->document_element();
        }
        else {
            m_xmlDocument.reset();
        }
    }
    return *this;
}
XLCfIconSet& XLCfIconSet::operator=(XLCfIconSet&& other) noexcept
{
    if (this != &other) {
        m_xmlDocument = std::move(other.m_xmlDocument);
        m_iconSetNode = other.m_iconSetNode;
    }
    return *this;
}

std::string XLCfIconSet::iconSet() const { return m_iconSetNode.attribute("iconSet").value(); }
void        XLCfIconSet::setIconSet(const std::string& iconSetName)
{
    setAttr(m_iconSetNode, "iconSet", iconSetName.c_str());
}
std::vector<XLCfvo> XLCfIconSet::cfvos() const
{
    std::vector<XLCfvo> result;
    for (auto& node : m_iconSetNode.children("cfvo")) result.emplace_back(node);
    return result;
}
void XLCfIconSet::addValue(XLCfvoType type, const std::string& value)
{
    auto cfvo = m_iconSetNode.append_child("cfvo");
    cfvo.append_attribute("type").set_value(XLCfvoTypeToString(type).c_str());
    if (!value.empty()) cfvo.append_attribute("val").set_value(value.c_str());
}
void XLCfIconSet::clear() { m_iconSetNode.remove_children(); }
bool XLCfIconSet::showValue() const { return m_iconSetNode.attribute("showValue").as_bool(true); }
void XLCfIconSet::setShowValue(bool show) { setAttrBool01(m_iconSetNode, "showValue", show); }
bool XLCfIconSet::percent() const { return m_iconSetNode.attribute("percent").as_bool(true); }
void XLCfIconSet::setPercent(bool percent) { setAttrBool01(m_iconSetNode, "percent", percent); }
bool XLCfIconSet::reverse() const { return m_iconSetNode.attribute("reverse").as_bool(false); }
void XLCfIconSet::setReverse(bool reverse) { setAttrBool01(m_iconSetNode, "reverse", reverse); }


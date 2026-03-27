// ===== External Includes ===== //
#include <cstdint>
#include <gsl/gsl>
#include <fmt/format.h>
#include <memory>     
#include <pugixml.hpp>
#include <stdexcept>   
#include <string>      
#include <vector>      

// ===== OpenXLSX Includes ===== //
#include "XLColor.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLStyles.hpp"
#include "XLUtilities.hpp"
#include "XLStyles_Internal.hpp"

using namespace OpenXLSX;

namespace {
    constexpr EnumStringMap<XLLineType> XLLineTypeMap[] = {
        {"left", XLLineLeft},
        {"right", XLLineRight},
        {"top", XLLineTop},
        {"bottom", XLLineBottom},
        {"diagonal", XLLineDiagonal},
        {"vertical", XLLineVertical},
        {"horizontal", XLLineHorizontal},
    };

    [[maybe_unused]] XLLineType XLLineTypeFromString(std::string_view str) {
        return EnumFromString(str, XLLineTypeMap, XLLineInvalid);
    }

    const char* XLLineTypeToString(XLLineType val) {
        return EnumToString(val, XLLineTypeMap, "(invalid)");
    }

    constexpr EnumStringMap<XLLineStyle> XLLineStyleMap[] = {
        {"", XLLineStyleNone},
        {"thin", XLLineStyleThin},
        {"medium", XLLineStyleMedium},
        {"dashed", XLLineStyleDashed},
        {"dotted", XLLineStyleDotted},
        {"thick", XLLineStyleThick},
        {"double", XLLineStyleDouble},
        {"hair", XLLineStyleHair},
        {"mediumDashed", XLLineStyleMediumDashed},
        {"dashDot", XLLineStyleDashDot},
        {"mediumDashDot", XLLineStyleMediumDashDot},
        {"dashDotDot", XLLineStyleDashDotDot},
        {"mediumDashDotDot", XLLineStyleMediumDashDotDot},
        {"slantDashDot", XLLineStyleSlantDashDot},
    };

    XLLineStyle XLLineStyleFromString(std::string_view str) {
        return EnumFromString(str, XLLineStyleMap, XLLineStyleInvalid);
    }

    const char* XLLineStyleToString(XLLineStyle val) {
        return EnumToString(val, XLLineStyleMap, "(invalid)");
    }
}

XLLine::XLLine() : m_lineNode(std::make_unique<XMLNode>()) {}

XLLine::XLLine(const XMLNode& node) : m_lineNode(std::make_unique<XMLNode>(node)) {}

XLLine::~XLLine() = default;

XLLine::XLLine(const XLLine& other) : m_lineNode(std::make_unique<XMLNode>(*other.m_lineNode)) {}

XLLine& XLLine::operator=(const XLLine& other)
{
    if (&other != this) *m_lineNode = *other.m_lineNode;
    return *this;
}

XLLineStyle XLLine::style() const
{
    if (m_lineNode->empty()) return XLLineStyleNone;
    XMLAttribute attr = appendAndGetAttribute(*m_lineNode, "style", OpenXLSX::XLDefaultLineStyle);
    return XLLineStyleFromString(attr.value());
}

XLLine::operator bool() const { return (style() != XLLineStyleNone); }

XLDataBarColor XLLine::color() const
{
    XMLNode color = appendAndGetNode(*m_lineNode, "color");
    if (color.empty()) return XLDataBarColor{};
    return XLDataBarColor(color);
}

std::string XLLine::summary() const
{
    return fmt::format("line style is {}, {}", XLLineStyleToString(style()), color().summary());
}

XLBorder::XLBorder() : m_borderNode(std::make_unique<XMLNode>()) {}

XLBorder::XLBorder(const XMLNode& node) : m_borderNode(std::make_unique<XMLNode>(node)) {}

XLBorder::~XLBorder() = default;

XLBorder::XLBorder(const XLBorder& other) : m_borderNode(std::make_unique<XMLNode>(*other.m_borderNode)) {}

XLBorder& XLBorder::operator=(const XLBorder& other)
{
    if (&other != this) *m_borderNode = *other.m_borderNode;
    return *this;
}

bool XLBorder::diagonalUp() const { return m_borderNode->attribute("diagonalUp").as_bool(); }

bool XLBorder::diagonalDown() const { return m_borderNode->attribute("diagonalDown").as_bool(); }

bool XLBorder::outline() const { return m_borderNode->attribute("outline").as_bool(); }

XLLine XLBorder::left() const { return XLLine(m_borderNode->child("left")); }
XLLine XLBorder::right() const { return XLLine(m_borderNode->child("right")); }
XLLine XLBorder::top() const { return XLLine(m_borderNode->child("top")); }
XLLine XLBorder::bottom() const { return XLLine(m_borderNode->child("bottom")); }
XLLine XLBorder::diagonal() const { return XLLine(m_borderNode->child("diagonal")); }
XLLine XLBorder::vertical() const { return XLLine(m_borderNode->child("vertical")); }
XLLine XLBorder::horizontal() const { return XLLine(m_borderNode->child("horizontal")); }

XLBorder& XLBorder::setDiagonalUp(bool set)
{ appendAndSetAttribute(*m_borderNode, "diagonalUp", (set ? "true" : "false")).empty(); return *this; }
XLBorder& XLBorder::setDiagonalDown(bool set)
{ appendAndSetAttribute(*m_borderNode, "diagonalDown", (set ? "true" : "false")).empty(); return *this; }
XLBorder& XLBorder::setOutline(bool set) { appendAndSetAttribute(*m_borderNode, "outline", (set ? "true" : "false")).empty(); return *this; }
bool XLBorder::setLine(XLLineType lineType, XLLineStyle lineStyle, XLColor lineColor, double lineTint)
{
    XMLNode lineNode = appendAndGetNode(*m_borderNode, XLLineTypeToString(lineType), m_nodeOrder);   
    // 2024-12-19: non-existing lines are added using an ordered insert to address issue #304
    bool success = (lineNode.empty() == false);
    if (success) {
        const char* styleStr = XLLineStyleToString(lineStyle);
        if (styleStr && *styleStr)
            success = (appendAndSetAttribute(lineNode, "style", styleStr).empty() == false);   
        else
            lineNode.remove_attribute("style");
    }
    XMLNode colorNode{};                                                                                         
    if (success) colorNode = appendAndGetNode(lineNode, "color");   
    XLDataBarColor colorObject{colorNode};
    success = (colorNode.empty() == false);
    if (success) success = colorObject.setRgb(lineColor);
    if (success) success = colorObject.setTint(lineTint);
    return success;
}
bool XLBorder::setLeft(XLLineStyle lineStyle, XLColor lineColor, double lineTint)
{ return setLine(XLLineLeft, lineStyle, lineColor, lineTint); }
bool XLBorder::setRight(XLLineStyle lineStyle, XLColor lineColor, double lineTint)
{ return setLine(XLLineRight, lineStyle, lineColor, lineTint); }
bool XLBorder::setTop(XLLineStyle lineStyle, XLColor lineColor, double lineTint)
{ return setLine(XLLineTop, lineStyle, lineColor, lineTint); }
bool XLBorder::setBottom(XLLineStyle lineStyle, XLColor lineColor, double lineTint)
{ return setLine(XLLineBottom, lineStyle, lineColor, lineTint); }
bool XLBorder::setDiagonal(XLLineStyle lineStyle, XLColor lineColor, double lineTint)
{ return setLine(XLLineDiagonal, lineStyle, lineColor, lineTint); }
bool XLBorder::setVertical(XLLineStyle lineStyle, XLColor lineColor, double lineTint)
{ return setLine(XLLineVertical, lineStyle, lineColor, lineTint); }
bool XLBorder::setHorizontal(XLLineStyle lineStyle, XLColor lineColor, double lineTint)
{ return setLine(XLLineHorizontal, lineStyle, lineColor, lineTint); }

std::string XLBorder::summary() const
{
    std::string lineInfo = fmt::format(", left: {}, right: {}, top: {}, bottom: {}, diagonal: {}, vertical: {}, horizontal: {}",
                                       left().summary(),
                                       right().summary(),
                                       top().summary(),
                                       bottom().summary(),
                                       diagonal().summary(),
                                       vertical().summary(),
                                       horizontal().summary());
    return fmt::format("diagonalUp: {}, diagonalDown: {}, outline: {}{}",
                       diagonalUp() ? "true" : "false",
                       diagonalDown() ? "true" : "false",
                       outline() ? "true" : "false",
                       lineInfo);
}

// ===== XLBorders, parent of XLBorder

XLBorders::XLBorders() : m_bordersNode(std::make_unique<XMLNode>()) {}

XLBorders::XLBorders(const XMLNode& borders) : m_bordersNode(std::make_unique<XMLNode>(borders))
{
    // initialize XLBorders entries and m_borders here
    XMLNode node = borders.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string_view nodeName(node.name());
        if (nodeName == "border")
            m_borders.push_back(XLBorder(node));
        else
            std::cerr << "WARNING: XLBorders constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLBorders::~XLBorders()
{
    m_borders.clear();   
}

XLBorders::XLBorders(const XLBorders& other) : m_bordersNode(std::make_unique<XMLNode>(*other.m_bordersNode)), m_borders(other.m_borders) {}

XLBorders::XLBorders(XLBorders&& other) : m_bordersNode(std::move(other.m_bordersNode)), m_borders(std::move(other.m_borders)) {}

XLBorders& XLBorders::operator=(const XLBorders& other)
{
    if (&other != this) {
        *m_bordersNode = *other.m_bordersNode;
        m_borders.clear();
        m_borders = other.m_borders;
    }
    return *this;
}

size_t XLBorders::count() const { return m_borders.size(); }

XLBorder XLBorders::borderByIndex(XLStyleIndex index) const
{
    Expects(index < m_borders.size());
    return m_borders.at(index);
}

XLStyleIndex XLBorders::create(XLBorder copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();   
    XMLNode      newNode{};         

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_bordersNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty())
        newNode = m_bordersNode->prepend_child("border");
    else
        newNode = m_bordersNode->insert_child_after("border", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLBorders::"s + __func__ + ": failed to append a new border node"s);
    }
    if (styleEntriesPrefix.length() > 0)   
        m_bordersNode->insert_child_before(pugi::node_pcdata, newNode)
            .set_value(std::string(styleEntriesPrefix).c_str());   

    XLBorder newBorder(newNode);
    if (copyFrom.m_borderNode->empty()) {   
        // ===== Create a border with default values
        newBorder.setLeft(OpenXLSX::XLLineStyleNone, XLColor(OpenXLSX::XLDefaultFontColor));
        newBorder.setRight(OpenXLSX::XLLineStyleNone, XLColor(OpenXLSX::XLDefaultFontColor));
        newBorder.setTop(OpenXLSX::XLLineStyleNone, XLColor(OpenXLSX::XLDefaultFontColor));
        newBorder.setBottom(OpenXLSX::XLLineStyleNone, XLColor(OpenXLSX::XLDefaultFontColor));
        newBorder.setDiagonal(OpenXLSX::XLLineStyleNone, XLColor(OpenXLSX::XLDefaultFontColor));
    }
    else
        copyXMLNode(newNode, *copyFrom.m_borderNode);   

    m_borders.push_back(newBorder);
    appendAndSetAttribute(*m_bordersNode, "count", std::to_string(m_borders.size()));   
    return index;
}

XLStyleIndex XLBorders::findOrCreate(XLBorder copyFrom, std::string_view styleEntriesPrefix)
{
    // Compute the fingerprint of the requested border
    std::string key = xmlNodeFingerprint(*copyFrom.m_borderNode);

    // Fast path: cache hit
    auto it = m_fingerprintCache.find(key);
    if (it != m_fingerprintCache.end()) return it->second;

    // Cold path: scan all existing borders (covers borders loaded from an existing file)
    for (size_t i = 0; i < m_borders.size(); ++i) {
        if (xmlNodeFingerprint(*m_borders[i].m_borderNode) == key) {
            m_fingerprintCache.emplace(key, i);
            return i;
        }
    }

    // No match found — create and cache
    XLStyleIndex idx = create(copyFrom, styleEntriesPrefix);
    m_fingerprintCache.emplace(key, idx);
    return idx;
}

// ===== External Includes ===== //
#include <cstdint>
#include <gsl/gsl>
#include <fmt/format.h>
#include <memory>      // std::make_unique
#include <pugixml.hpp>
#include <stdexcept>    // std::invalid_argument
#include <string>       // std::stoi, std::literals::string_literals
#include <vector>       // std::vector

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

/**
 * @details Constructor. Initializes an empty XLLine object
 */
XLLine::XLLine() : m_lineNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLLine object.
 */
XLLine::XLLine(const XMLNode& node) : m_lineNode(std::make_unique<XMLNode>(node)) {}

XLLine::~XLLine() = default;

XLLine::XLLine(const XLLine& other) : m_lineNode(std::make_unique<XMLNode>(*other.m_lineNode)) {}

XLLine& XLLine::operator=(const XLLine& other)
{
    if (&other != this) *m_lineNode = *other.m_lineNode;
    return *this;
}

/**
 * @details Returns the line style (XLLineStyleNone if line is not set)
 */
XLLineStyle XLLine::style() const
{
    if (m_lineNode->empty()) return XLLineStyleNone;
    XMLAttribute attr = appendAndGetAttribute(*m_lineNode, "style", OpenXLSX::XLDefaultLineStyle);
    return XLLineStyleFromString(attr.value());
}

/**
 * @details check if line is used (set) or not
 */
XLLine::operator bool() const { return (style() != XLLineStyleNone); }

/**
 * @details Returns the line data bar color object
 */
XLDataBarColor XLLine::color() const
{
    XMLNode color = appendAndGetNode(*m_lineNode, "color");
    if (color.empty()) return XLDataBarColor{};
    return XLDataBarColor(color);
}
// XLColor XLLine::color() const
// {
//     XMLNode colorDetails = m_lineNode->child("color");
//     if (colorDetails.empty()) return XLColor("ffffffff");
//     XMLAttribute colorRGB = colorDetails.attribute("rgb");
//     if (colorRGB.empty()) return XLColor("ffffffff");
//     return XLColor(colorRGB.value());
// }
//
// /**
//  * @details Returns the line color tint
//  */
// double XLLine::colorTint() const
// {
//     XMLNode colorDetails = m_lineNode->child("color");
//     if (not colorDetails.empty()) {
//         XMLAttribute colorTint = colorDetails.attribute("tint");
//         if (not colorTint.empty())
//             return colorTint.as_double();
//     }
//     return 0.0;
// }

/**
 * @details assemble a string summary about the fill
 */
std::string XLLine::summary() const
{
    return fmt::format("line style is {}, {}", XLLineStyleToString(style()), color().summary());
    // double tint = colorTint();
    // std::string tintSummary = fmt::format("colorTint is {}", (tint == 0.0 ? "(none)" : std::to_string(tint)));
    // size_t tintDecimalPos = tintSummary.find_last_of('.');
    // if (tintDecimalPos != std::string::npos) tintSummary = tintSummary.substr(0, tintDecimalPos + 3); // truncate colorTint double output
    // 2 digits after the decimal separator return fmt::format("line style is {}, color is {}, {}", XLLineStyleToString(style()), color().hex(), tintSummary);
}

/**
 * @details Constructor. Initializes an empty XLBorder object
 */
XLBorder::XLBorder() : m_borderNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLBorder object.
 */
XLBorder::XLBorder(const XMLNode& node) : m_borderNode(std::make_unique<XMLNode>(node)) {}

XLBorder::~XLBorder() = default;

XLBorder::XLBorder(const XLBorder& other) : m_borderNode(std::make_unique<XMLNode>(*other.m_borderNode)) {}

XLBorder& XLBorder::operator=(const XLBorder& other)
{
    if (&other != this) *m_borderNode = *other.m_borderNode;
    return *this;
}

/**
 * @details determines whether the diagonalUp property is set
 */
bool XLBorder::diagonalUp() const { return m_borderNode->attribute("diagonalUp").as_bool(); }

/**
 * @details determines whether the diagonalDown property is set
 */
bool XLBorder::diagonalDown() const { return m_borderNode->attribute("diagonalDown").as_bool(); }

/**
 * @details determines whether the outline property is set
 */
bool XLBorder::outline() const { return m_borderNode->attribute("outline").as_bool(); }

/**
 * @details fetch lines
 */
XLLine XLBorder::left() const { return XLLine(m_borderNode->child("left")); }
XLLine XLBorder::right() const { return XLLine(m_borderNode->child("right")); }
XLLine XLBorder::top() const { return XLLine(m_borderNode->child("top")); }
XLLine XLBorder::bottom() const { return XLLine(m_borderNode->child("bottom")); }
XLLine XLBorder::diagonal() const { return XLLine(m_borderNode->child("diagonal")); }
XLLine XLBorder::vertical() const { return XLLine(m_borderNode->child("vertical")); }
XLLine XLBorder::horizontal() const { return XLLine(m_borderNode->child("horizontal")); }

/**
 * @details Setter functions
 */
bool XLBorder::setDiagonalUp(bool set)
{ return appendAndSetAttribute(*m_borderNode, "diagonalUp", (set ? "true" : "false")).empty() == false; }
bool XLBorder::setDiagonalDown(bool set)
{ return appendAndSetAttribute(*m_borderNode, "diagonalDown", (set ? "true" : "false")).empty() == false; }
bool XLBorder::setOutline(bool set) { return appendAndSetAttribute(*m_borderNode, "outline", (set ? "true" : "false")).empty() == false; }
bool XLBorder::setLine(XLLineType lineType, XLLineStyle lineStyle, XLColor lineColor, double lineTint)
{
    XMLNode lineNode = appendAndGetNode(*m_borderNode, XLLineTypeToString(lineType), m_nodeOrder);    // generate line node if not present
    // 2024-12-19: non-existing lines are added using an ordered insert to address issue #304
    bool success = (lineNode.empty() == false);
    if (success) {
        const char* styleStr = XLLineStyleToString(lineStyle);
        if (styleStr && *styleStr)
            success = (appendAndSetAttribute(lineNode, "style", styleStr).empty() == false);    // set style attribute
        else
            lineNode.remove_attribute("style");
    }
    XMLNode colorNode{};                                                                                          // empty node
    if (success) colorNode = appendAndGetNode(lineNode, "color");    // generate color node if not present
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

/**
 * @details assemble a string summary about the fill
 */
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

/**
 * @details Constructor. Initializes an empty XLBorders object
 */
XLBorders::XLBorders() : m_bordersNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLBorders object.
 */
XLBorders::XLBorders(const XMLNode& borders) : m_bordersNode(std::make_unique<XMLNode>(borders))
{
    // initialize XLBorders entries and m_borders here
    XMLNode node = borders.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        if (nodeName == "border")
            m_borders.push_back(XLBorder(node));
        else
            std::cerr << "WARNING: XLBorders constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLBorders::~XLBorders()
{
    m_borders.clear();    // delete vector with all children
}

XLBorders::XLBorders(const XLBorders& other) : m_bordersNode(std::make_unique<XMLNode>(*other.m_bordersNode)), m_borders(other.m_borders) {}

XLBorders::XLBorders(XLBorders&& other) : m_bordersNode(std::move(other.m_bordersNode)), m_borders(std::move(other.m_borders)) {}

/**
 * @details Copy assignment operator
 */
XLBorders& XLBorders::operator=(const XLBorders& other)
{
    if (&other != this) {
        *m_bordersNode = *other.m_bordersNode;
        m_borders.clear();
        m_borders = other.m_borders;
    }
    return *this;
}

/**
 * @details Returns the amount of border descriptions held by the class
 */
size_t XLBorders::count() const { return m_borders.size(); }

/**
 * @details fetch XLBorder from m_borders by index
 */
XLBorder XLBorders::borderByIndex(XLStyleIndex index) const
{
    Expects(index < m_borders.size());
    return m_borders.at(index);
}

/**
 * @details append a new XLBorder to m_borders and m_bordersNode, based on copyFrom
 */
XLStyleIndex XLBorders::create(XLBorder copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the border to be created
    XMLNode      newNode{};          // scope declaration

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
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_bordersNode->insert_child_before(pugi::node_pcdata, newNode)
            .set_value(std::string(styleEntriesPrefix).c_str());    // prefix the new node with styleEntriesPrefix

    XLBorder newBorder(newNode);
    if (copyFrom.m_borderNode->empty()) {    // if no template is given
        // ===== Create a border with default values
        newBorder.setLeft(OpenXLSX::XLLineStyleNone, XLColor(OpenXLSX::XLDefaultFontColor));
        newBorder.setRight(OpenXLSX::XLLineStyleNone, XLColor(OpenXLSX::XLDefaultFontColor));
        newBorder.setTop(OpenXLSX::XLLineStyleNone, XLColor(OpenXLSX::XLDefaultFontColor));
        newBorder.setBottom(OpenXLSX::XLLineStyleNone, XLColor(OpenXLSX::XLDefaultFontColor));
        newBorder.setDiagonal(OpenXLSX::XLLineStyleNone, XLColor(OpenXLSX::XLDefaultFontColor));
    }
    else
        copyXMLNode(newNode, *copyFrom.m_borderNode);    // will use copyFrom as template, does nothing if copyFrom is empty

    m_borders.push_back(newBorder);
    appendAndSetAttribute(*m_bordersNode, "count", std::to_string(m_borders.size()));    // update array count in XML
    return index;
}


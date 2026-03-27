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
    constexpr EnumStringMap<XLFillType> XLFillTypeMap[] = {
        {"gradientFill", XLGradientFill},
        {"patternFill", XLPatternFill},
    };

    XLFillType XLFillTypeFromString(std::string_view str) {
        return EnumFromString(str, XLFillTypeMap, XLFillTypeInvalid);
    }

    const char* XLFillTypeToString(XLFillType val) {
        return EnumToString(val, XLFillTypeMap, "(invalid)");
    }

    constexpr EnumStringMap<XLGradientType> XLGradientTypeMap[] = {
        {"linear", XLGradientLinear},
        {"path", XLGradientPath},
    };

    XLGradientType XLGradientTypeFromString(std::string_view str) {
        return EnumFromString(str, XLGradientTypeMap, XLGradientTypeInvalid);
    }

    const char* XLGradientTypeToString(XLGradientType val) {
        return EnumToString(val, XLGradientTypeMap, "(invalid)");
    }

    constexpr EnumStringMap<XLPatternType> XLPatternTypeMap[] = {
        {"", XLPatternNone},
        {"solid", XLPatternSolid},
        {"mediumGray", XLPatternMediumGray},
        {"darkGray", XLPatternDarkGray},
        {"lightGray", XLPatternLightGray},
        {"darkHorizontal", XLPatternDarkHorizontal},
        {"darkVertical", XLPatternDarkVertical},
        {"darkDown", XLPatternDarkDown},
        {"darkUp", XLPatternDarkUp},
        {"darkGrid", XLPatternDarkGrid},
        {"darkTrellis", XLPatternDarkTrellis},
        {"lightHorizontal", XLPatternLightHorizontal},
        {"lightVertical", XLPatternLightVertical},
        {"lightDown", XLPatternLightDown},
        {"lightUp", XLPatternLightUp},
        {"lightGrid", XLPatternLightGrid},
        {"lightTrellis", XLPatternLightTrellis},
        {"gray125", XLPatternGray125},
        {"gray0625", XLPatternGray0625},
        {"none", XLPatternNone},
    };

    XLPatternType XLPatternTypeFromString(std::string_view str) {
        return EnumFromString(str, XLPatternTypeMap, XLPatternTypeInvalid);
    }

    const char* XLPatternTypeToString(XLPatternType val) {
        return EnumToString(val, XLPatternTypeMap, "(invalid)");
    }
}


// ===== XLDataBarColor, used by XLFills gradientFill and by XLLine (to be implemented)

XLDataBarColor::XLDataBarColor() : m_colorNode(std::make_unique<XMLNode>()) {}

XLDataBarColor::XLDataBarColor(const XMLNode& node) : m_colorNode(std::make_unique<XMLNode>(node)) {}

XLDataBarColor::XLDataBarColor(const XLDataBarColor& other) : m_colorNode(std::make_unique<XMLNode>(*other.m_colorNode)) {}

XLDataBarColor& XLDataBarColor::operator=(const XLDataBarColor& other)
{
    if (&other != this) *m_colorNode = *other.m_colorNode;
    return *this;
}

XLColor  XLDataBarColor::rgb() const { return XLColor(m_colorNode->attribute("rgb").as_string("ffffffff")); }
double   XLDataBarColor::tint() const { return m_colorNode->attribute("tint").as_double(0.0); }
bool     XLDataBarColor::automatic() const { return m_colorNode->attribute("auto").as_bool(); }
uint32_t XLDataBarColor::indexed() const { return m_colorNode->attribute("indexed").as_uint(); }
uint32_t XLDataBarColor::theme() const { return m_colorNode->attribute("theme").as_uint(); }
bool XLDataBarColor::setRgb(XLColor newColor) { return appendAndSetAttribute(*m_colorNode, "rgb", newColor.hex()).empty() == false; }
bool XLDataBarColor::setTint(double newTint)
{
    std::string tintString = "";
    if (newTint != 0.0) {
        tintString = checkAndFormatDoubleAsString(newTint, -1.0, +1.0, 0.01);
        if (tintString.length() == 0) {
            using namespace std::literals::string_literals;
            throw XLException("XLDataBarColor::setTint: color tint "s + fmt::format("{}", newTint) + " is not in range [-1.0;+1.0]"s);
        }
    }
    if (tintString.length() == 0) return m_colorNode->remove_attribute("tint");    // remove tint attribute for a value 0

    return (appendAndSetAttribute(*m_colorNode, "tint", tintString).empty() == false);    // else: set tint attribute
}
bool XLDataBarColor::setAutomatic(bool set)
{ return appendAndSetAttribute(*m_colorNode, "auto", (set ? "true" : "false")).empty() == false; }
bool XLDataBarColor::setIndexed(uint32_t newIndex)
{ return appendAndSetAttribute(*m_colorNode, "indexed", std::to_string(newIndex)).empty() == false; }
bool XLDataBarColor::setTheme(uint32_t newTheme)
{
    if (newTheme == XLDeleteProperty) return m_colorNode->remove_attribute("theme");
    return appendAndSetAttribute(*m_colorNode, "theme", std::to_string(newTheme)).empty() == false;
}

std::string XLDataBarColor::summary() const
{
    return fmt::format("rgb is {}, tint is {}{}{}{}",
                       rgb().hex(),
                       formatDoubleAsString(tint()),
                       automatic() ? ", +automatic" : "",
                       indexed() ? fmt::format(", index is {}", indexed()) : "",
                       theme() ? fmt::format(", theme is {}", theme()) : "");
}

XLGradientStop::XLGradientStop() : m_stopNode(std::make_unique<XMLNode>()) {}

XLGradientStop::XLGradientStop(const XMLNode& node) : m_stopNode(std::make_unique<XMLNode>(node)) {}

XLGradientStop::XLGradientStop(const XLGradientStop& other) : m_stopNode(std::make_unique<XMLNode>(*other.m_stopNode)) {}

XLGradientStop& XLGradientStop::operator=(const XLGradientStop& other)
{
    if (&other != this) *m_stopNode = *other.m_stopNode;
    return *this;
}

XLDataBarColor XLGradientStop::color() const
{
    XMLNode color = appendAndGetNode(*m_stopNode, "color");
    if (color.empty()) return XLDataBarColor{};
    return XLDataBarColor(color);
}
double XLGradientStop::position() const
{
    XMLAttribute attr = m_stopNode->attribute("position");
    return attr.as_double(0.0);
}

bool XLGradientStop::setPosition(double newPosition)
{ return appendAndSetAttribute(*m_stopNode, "position", formatDoubleAsString(newPosition)).empty() == false; }

std::string XLGradientStop::summary() const
{
    return fmt::format("stop position is {}, {}", formatDoubleAsString(position()), color().summary());
}

// ===== XLGradientStops, parent of XLGradientStop

XLGradientStops::XLGradientStops() : m_gradientNode(std::make_unique<XMLNode>()) {}

XLGradientStops::XLGradientStops(const XMLNode& gradient) : m_gradientNode(std::make_unique<XMLNode>(gradient))
{
    // initialize XLGradientStops entries and m_gradientStops here
    XMLNode node = m_gradientNode->first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string_view nodeName(node.name());
        if (nodeName == "stop")
            m_gradientStops.push_back(XLGradientStop(node));
        else
            std::cerr << "WARNING: XLGradientStops constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLGradientStops::~XLGradientStops()
{
    m_gradientStops.clear();   
}

XLGradientStops::XLGradientStops(const XLGradientStops& other)
    : m_gradientNode(std::make_unique<XMLNode>(*other.m_gradientNode)),
      m_gradientStops(other.m_gradientStops)
{}

XLGradientStops::XLGradientStops(XLGradientStops&& other)
    : m_gradientNode(std::move(other.m_gradientNode)),
      m_gradientStops(std::move(other.m_gradientStops))
{}

XLGradientStops& XLGradientStops::operator=(const XLGradientStops& other)
{
    if (&other != this) {
        *m_gradientNode = *other.m_gradientNode;
        m_gradientStops.clear();
        m_gradientStops = other.m_gradientStops;
    }
    return *this;
}

size_t XLGradientStops::count() const { return m_gradientStops.size(); }

XLGradientStop XLGradientStops::stopByIndex(XLStyleIndex index) const
{
    Expects(index < m_gradientStops.size());
    return m_gradientStops.at(index);
}

XLStyleIndex XLGradientStops::create(XLGradientStop copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();   
    XMLNode      newNode{};         

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_gradientNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty())
        newNode = m_gradientNode->prepend_child("stop");
    else
        newNode = m_gradientNode->insert_child_after("stop", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLGradientStops::"s + __func__ + ": failed to append a new stop node"s);
    }
    if (styleEntriesPrefix.length() > 0)   
        m_gradientNode->insert_child_before(pugi::node_pcdata, newNode)
            .set_value(std::string(styleEntriesPrefix).c_str());   

    XLGradientStop newStop(newNode);
    if (copyFrom.m_stopNode->empty()) {   
        // ===== Create a stop node with default values
        newStop.setPosition(0.0);
        newStop.color().setRgb(XLColor(OpenXLSX::XLDefaultPatternFgColor));
    }
    else
        copyXMLNode(newNode, *copyFrom.m_stopNode);   

    m_gradientStops.push_back(newStop);
    appendAndSetAttribute(*m_gradientNode, "count", std::to_string(m_gradientStops.size()));   
    return index;
}

std::string XLGradientStops::summary() const
{
    std::string result{};
    for (const auto& stop : m_gradientStops) result += stop.summary() + ", ";
    if (result.length() < 2) return "";    // if no stop summary was created - return empty string
    return result.substr(0, result.length() - 2);
}

XLFill::XLFill() : m_fillNode(std::make_unique<XMLNode>()) {}

XLFill::XLFill(const XMLNode& node) : m_fillNode(std::make_unique<XMLNode>(node)) {}

XLFill::~XLFill() = default;

XLFill::XLFill(const XLFill& other) : m_fillNode(std::make_unique<XMLNode>(*other.m_fillNode)) {}

XLFill& XLFill::operator=(const XLFill& other)
{
    if (&other != this) *m_fillNode = *other.m_fillNode;
    return *this;
}

/**
 * @details Returns the name of the first element child of fill
 * @note an empty node ::name() returns an empty string "", leading to XLFillTypeInvalid
 */
XLFillType XLFill::fillType() const { return XLFillTypeFromString(m_fillNode->first_child_of_type(pugi::node_element).name()); }

bool XLFill::setFillType(XLFillType newFillType, bool force)
{
    XLFillType ft = fillType();    // determine once, use twice

    // ===== If desired filltype is already set
    if (ft == newFillType) return true;    // nothing to do

    // ===== If force == true or fillType is just not set at all, delete existing child nodes, otherwise throw
    if (!force and ft != XLFillTypeInvalid) {
        using namespace std::literals::string_literals;
        throw XLException("XLFill::setFillType: can not change the fillType from "s + XLFillTypeToString(fillType()) + " to "s +
                          XLFillTypeToString(newFillType) + " - invoke with force == true if override is desired"s);
    }
    // ===== At this point, m_fillNode needs to be emptied for a setting / force-change of the fill type

    // ===== Delete all fill node children and insert a new node for the newFillType
    m_fillNode->remove_children();
    return (m_fillNode->append_child(XLFillTypeToString(newFillType)).empty() == false);
}

void XLFill::throwOnFillType(XLFillType typeToThrowOn, const char* functionName) const
{
    using namespace std::literals::string_literals;
    if (fillType() == typeToThrowOn)
        throw XLException("XLFill::"s + functionName + " must not be invoked for a "s + XLFillTypeToString(typeToThrowOn));
}

XMLNode XLFill::getValidFillDescription(XLFillType fillTypeIfEmpty, const char* functionName)
{
    XLFillType throwOnThis = XLFillTypeInvalid;
    switch (fillTypeIfEmpty) {
        case XLGradientFill:
            throwOnThis = XLPatternFill;
            break;    // throw on non-matching fill type
        case XLPatternFill:
            throwOnThis = XLGradientFill;
            break;    //   "
        default:
            throw XLException("XLFill::getValidFillDescription was not invoked with XLPatternFill or XLGradientFill");
    }
    throwOnFillType(throwOnThis, functionName);
    XMLNode fillDescription = m_fillNode->first_child_of_type(pugi::node_element);    // fetch an existing fill description
    if (fillDescription.empty() and setFillType(fillTypeIfEmpty))                      // if none exists, attempt to create a description
        fillDescription = m_fillNode->first_child_of_type(pugi::node_element);        // fetch newly inserted description
    return fillDescription;
}

XLGradientType XLFill::gradientType()
{ return XLGradientTypeFromString(getValidFillDescription(XLGradientFill, __func__).attribute("type").value()); }
double          XLFill::degree() { return getValidFillDescription(XLGradientFill, __func__).attribute("degree").as_double(0); }
double          XLFill::left() { return getValidFillDescription(XLGradientFill, __func__).attribute("left").as_double(0); }
double          XLFill::right() { return getValidFillDescription(XLGradientFill, __func__).attribute("right").as_double(0); }
double          XLFill::top() { return getValidFillDescription(XLGradientFill, __func__).attribute("top").as_double(0); }
double          XLFill::bottom() { return getValidFillDescription(XLGradientFill, __func__).attribute("bottom").as_double(0); }
XLGradientStops XLFill::stops() { return XLGradientStops(getValidFillDescription(XLGradientFill, __func__)); }

XLPatternType XLFill::patternType()
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return XLDefaultPatternType;    // if no description could be fetched: fail
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fillNode,
                                                  XLFillTypeToString(XLPatternFill),
                                                  "patternType",
                                                  XLPatternTypeToString(XLDefaultPatternType));
    return XLPatternTypeFromString(attr.value());
}

XLColor XLFill::color()
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return XLColor{};    // if no description could be fetched: fail
    XMLAttribute fgColorRGB = appendAndGetNodeAttribute(fillDescription, "fgColor", "rgb", XLDefaultPatternFgColor);
    return XLColor(fgColorRGB.value());
}
XLColor XLFill::backgroundColor()
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return XLColor{};    // if no description could be fetched: fail
    XMLAttribute bgColorRGB = appendAndGetNodeAttribute(fillDescription, "bgColor", "rgb", XLDefaultPatternBgColor);
    return XLColor(bgColorRGB.value());
}

bool XLFill::setGradientType(XLGradientType newType)
{
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "type", XLGradientTypeToString(newType)).empty() == false;
}
bool XLFill::setDegree(double newDegree)
{
    std::string degreeString = checkAndFormatDoubleAsString(newDegree, 0.0, 360.0, 0.01);
    if (degreeString.length() == 0) {
        using namespace std::literals::string_literals;
        throw XLException("XLFill::setDegree: gradientFill degree value "s + fmt::format("{}", newDegree) + " is not in range [0.0;360.0]"s);
    }
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "degree", degreeString).empty() == false;
}
bool XLFill::setLeft(double newLeft)
{
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "left", formatDoubleAsString(newLeft).c_str()).empty() == false;
}
bool XLFill::setRight(double newRight)
{
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "right", formatDoubleAsString(newRight).c_str()).empty() == false;
}
bool XLFill::setTop(double newTop)
{
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "top", formatDoubleAsString(newTop).c_str()).empty() == false;
}
bool XLFill::setBottom(double newBottom)
{
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "bottom", formatDoubleAsString(newBottom).c_str()).empty() == false;
}

XLFill& XLFill::setPatternType(XLPatternType newFillPattern)
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return *this;
    appendAndSetAttribute(fillDescription, "patternType", XLPatternTypeToString(newFillPattern));
    return *this;
}
XLFill& XLFill::setColor(XLColor newColor)
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return *this;
    std::vector<std::string_view> nodeOrder = {"fgColor", "bgColor"};
    appendAndSetNodeAttribute(fillDescription, "fgColor", "rgb", newColor.hex(), XLRemoveAttributes, nodeOrder);
    if (patternType() == XLPatternSolid)
        appendAndSetNodeAttribute(fillDescription, "bgColor", "rgb", newColor.hex(), XLRemoveAttributes, nodeOrder);
    return *this;
}
XLFill& XLFill::setBackgroundColor(XLColor newBgColor)
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return *this;
    std::vector<std::string_view> nodeOrder = {"fgColor", "bgColor"};
    appendAndSetNodeAttribute(fillDescription, "bgColor", "rgb", newBgColor.hex(), XLRemoveAttributes, nodeOrder);
    return *this;
}

std::string XLFill::summary()
{
    switch (fillType()) {
        case XLGradientFill:
            return fmt::format("fill type is {}, gradient type is {}, degree is: {}, left is: {}, right is: {}, top is: {}, bottom is: {}, stops: {}",
                               XLFillTypeToString(fillType()),
                               XLGradientTypeToString(gradientType()),
                               formatDoubleAsString(degree()),
                               formatDoubleAsString(left()),
                               formatDoubleAsString(right()),
                               formatDoubleAsString(top()),
                               formatDoubleAsString(bottom()),
                               stops().summary());
        case XLPatternFill:
            return fmt::format("fill type is {}, pattern type is {}, fgcolor is: {}, bgcolor: {}",
                               XLFillTypeToString(fillType()),
                               XLPatternTypeToString(patternType()),
                               color().hex(),
                               backgroundColor().hex());
        case XLFillTypeInvalid:
            [[fallthrough]];
        default:
            return "fill type is invalid!";
    }
}

// ===== XLFills, parent of XLFill

XLFills::XLFills() : m_fillsNode(std::make_unique<XMLNode>()) {}

XLFills::XLFills(const XMLNode& fills) : m_fillsNode(std::make_unique<XMLNode>(fills))
{
    // initialize XLFills entries and m_fills here
    XMLNode node = fills.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string_view nodeName(node.name());
        if (nodeName == "fill")
            m_fills.push_back(XLFill(node));
        else
            std::cerr << "WARNING: XLFills constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLFills::~XLFills()
{
    m_fills.clear();   
}

XLFills::XLFills(const XLFills& other) : m_fillsNode(std::make_unique<XMLNode>(*other.m_fillsNode)), m_fills(other.m_fills) {}

XLFills::XLFills(XLFills&& other) : m_fillsNode(std::move(other.m_fillsNode)), m_fills(std::move(other.m_fills)) {}

XLFills& XLFills::operator=(const XLFills& other)
{
    if (&other != this) {
        *m_fillsNode = *other.m_fillsNode;
        m_fills.clear();
        m_fills = other.m_fills;
    }
    return *this;
}

size_t XLFills::count() const { return m_fills.size(); }

XLFill XLFills::fillByIndex(XLStyleIndex index) const
{
    Expects(index < m_fills.size());
    return m_fills.at(index);
}

XLStyleIndex XLFills::create(XLFill copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();   
    XMLNode      newNode{};         

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_fillsNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty())
        newNode = m_fillsNode->prepend_child("fill");
    else
        newNode = m_fillsNode->insert_child_after("fill", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLFills::"s + __func__ + ": failed to append a new fill node"s);
    }
    if (styleEntriesPrefix.length() > 0)   
        m_fillsNode->insert_child_before(pugi::node_pcdata, newNode)
            .set_value(std::string(styleEntriesPrefix).c_str());   

    XLFill newFill(newNode);
    if (copyFrom.m_fillNode->empty()) {   
        // ===== Create a fill with default values
        newFill.setPatternType(OpenXLSX::XLDefaultPatternType);
    }
    else
        copyXMLNode(newNode, *copyFrom.m_fillNode);   

    m_fills.push_back(newFill);
    appendAndSetAttribute(*m_fillsNode, "count", std::to_string(m_fills.size()));   
    return index;
}

XLStyleIndex XLFills::findOrCreate(XLFill copyFrom, std::string_view styleEntriesPrefix)
{
    // Compute the fingerprint of the requested fill
    std::string key = xmlNodeFingerprint(*copyFrom.m_fillNode);

    // Fast path: cache hit
    auto it = m_fingerprintCache.find(key);
    if (it != m_fingerprintCache.end()) return it->second;

    // Cold path: scan all existing fills (covers fills loaded from an existing file)
    for (size_t i = 0; i < m_fills.size(); ++i) {
        if (xmlNodeFingerprint(*m_fills[i].m_fillNode) == key) {
            m_fingerprintCache.emplace(key, i);
            return i;
        }
    }

    // No match found — create and cache
    XLStyleIndex idx = create(copyFrom, styleEntriesPrefix);
    m_fingerprintCache.emplace(key, idx);
    return idx;
}

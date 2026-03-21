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
    constexpr EnumStringMap<XLUnderlineStyle> XLUnderlineStyleMap[] = {
        {"", XLUnderlineNone},
        {"single", XLUnderlineSingle},
        {"double", XLUnderlineDouble},
        {"none", XLUnderlineNone},
    };

    XLUnderlineStyle XLUnderlineStyleFromString(std::string_view str) {
        return EnumFromString(str, XLUnderlineStyleMap, XLUnderlineInvalid);
    }

    const char* XLUnderlineStyleToString(XLUnderlineStyle val) {
        return EnumToString(val, XLUnderlineStyleMap, "(invalid)");
    }

    constexpr EnumStringMap<XLFontSchemeStyle> XLFontSchemeStyleMap[] = {
        {"", XLFontSchemeNone},
        {"major", XLFontSchemeMajor},
        {"minor", XLFontSchemeMinor},
        {"none", XLFontSchemeNone},
    };

    XLFontSchemeStyle XLFontSchemeStyleFromString(std::string_view str) {
        return EnumFromString(str, XLFontSchemeStyleMap, XLFontSchemeInvalid);
    }

    const char* XLFontSchemeStyleToString(XLFontSchemeStyle val) {
        return EnumToString(val, XLFontSchemeStyleMap, "(invalid)");
    }

    constexpr EnumStringMap<XLVerticalAlignRunStyle> XLVerticalAlignRunStyleMap[] = {
        {"", XLBaseline},
        {"subscript", XLSubscript},
        {"superscript", XLSuperscript},
        {"baseline", XLBaseline},
    };

    XLVerticalAlignRunStyle XLVerticalAlignRunStyleFromString(std::string_view str) {
        return EnumFromString(str, XLVerticalAlignRunStyleMap, XLVerticalAlignRunInvalid);
    }

    const char* XLVerticalAlignRunStyleToString(XLVerticalAlignRunStyle val) {
        return EnumToString(val, XLVerticalAlignRunStyleMap, "(invalid)");
    }
}

/**
 * @details Constructor. Initializes an empty XLFont object
 */
XLFont::XLFont() : m_fontNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLFont object.
 */
XLFont::XLFont(const XMLNode& node) : m_fontNode(std::make_unique<XMLNode>(node)) {}

XLFont::~XLFont() = default;

XLFont::XLFont(const XLFont& other) : m_fontNode(std::make_unique<XMLNode>(*other.m_fontNode)) {}

XLFont& XLFont::operator=(const XLFont& other)
{
    if (&other != this) *m_fontNode = *other.m_fontNode;
    return *this;
}

/**
 * @details Returns the font name property
 */
std::string XLFont::fontName() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "name", "val", OpenXLSX::XLDefaultFontName);
    return attr.value();
}

/**
 * @details Returns the font charset property
 */
size_t XLFont::fontCharset() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "charset", "val", std::to_string(OpenXLSX::XLDefaultFontCharset));
    return attr.as_uint();
}

/**
 * @details Returns the font family property
 */
size_t XLFont::fontFamily() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "family", "val", std::to_string(OpenXLSX::XLDefaultFontFamily));
    return attr.as_uint();
}

/**
 * @details Returns the font size property
 */
size_t XLFont::fontSize() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "sz", "val", std::to_string(OpenXLSX::XLDefaultFontSize));
    return attr.as_uint();
}

/**
 * @details Returns the font color property
 */
XLColor XLFont::fontColor() const
{
    using namespace std::literals::string_literals;
    // XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "color", "theme", OpenXLSX::XLDefaultFontColorTheme);
    // TBD what "theme" is and whether it should be supported at all
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "color", "rgb", XLDefaultFontColor);
    return XLColor(attr.value());
}

/**
 * @details getter functions: return the font's bold, italic, underline, strikethrough status
 */
bool             XLFont::bold() const { return getBoolAttributeWhenOmittedMeansTrue(*m_fontNode, "b", "val"); }
bool             XLFont::italic() const { return getBoolAttributeWhenOmittedMeansTrue(*m_fontNode, "i", "val"); }
bool             XLFont::strikethrough() const { return getBoolAttributeWhenOmittedMeansTrue(*m_fontNode, "strike", "val"); }
XLUnderlineStyle XLFont::underline() const
{ return XLUnderlineStyleFromString(appendAndGetNodeAttribute(*m_fontNode, "u", "val", "none").value()); }
XLFontSchemeStyle XLFont::scheme() const
{ return XLFontSchemeStyleFromString(appendAndGetNodeAttribute(*m_fontNode, "scheme", "val", "none").value()); }
XLVerticalAlignRunStyle XLFont::vertAlign() const
{ return XLVerticalAlignRunStyleFromString(appendAndGetNodeAttribute(*m_fontNode, "vertAlign", "val", "baseline").value()); }
bool XLFont::outline() const { return getBoolAttributeWhenOmittedMeansTrue(*m_fontNode, "outline", "val"); }
bool XLFont::shadow() const { return getBoolAttributeWhenOmittedMeansTrue(*m_fontNode, "shadow", "val"); }
bool XLFont::condense() const { return getBoolAttributeWhenOmittedMeansTrue(*m_fontNode, "condense", "val"); }
bool XLFont::extend() const { return getBoolAttributeWhenOmittedMeansTrue(*m_fontNode, "extend", "val"); }

/**
 * @details Setter functions
 */
bool XLFont::setFontName(std::string_view newName)
{ return appendAndSetNodeAttribute(*m_fontNode, "name", "val", std::string(newName).c_str()).empty() == false; }
bool XLFont::setFontCharset(size_t newCharset)
{ return appendAndSetNodeAttribute(*m_fontNode, "charset", "val", std::to_string(newCharset)).empty() == false; }
bool XLFont::setFontFamily(size_t newFamily)
{ return appendAndSetNodeAttribute(*m_fontNode, "family", "val", std::to_string(newFamily)).empty() == false; }
bool XLFont::setFontSize(size_t newSize)
{ return appendAndSetNodeAttribute(*m_fontNode, "sz", "val", std::to_string(newSize)).empty() == false; }
bool XLFont::setFontColor(XLColor newColor)
{ return appendAndSetNodeAttribute(*m_fontNode, "color", "rgb", newColor.hex(), XLRemoveAttributes).empty() == false; }
bool XLFont::setBold(bool set) { return appendAndSetNodeAttribute(*m_fontNode, "b", "val", (set ? "1" : "0")).empty() == false; }
bool XLFont::setItalic(bool set) { return appendAndSetNodeAttribute(*m_fontNode, "i", "val", (set ? "1" : "0")).empty() == false; }
bool XLFont::setStrikethrough(bool set)
{ return appendAndSetNodeAttribute(*m_fontNode, "strike", "val", (set ? "1" : "0")).empty() == false; }
bool XLFont::setUnderline(XLUnderlineStyle style)
{ return appendAndSetNodeAttribute(*m_fontNode, "u", "val", XLUnderlineStyleToString(style)).empty() == false; }
bool XLFont::setScheme(XLFontSchemeStyle newScheme)
{ return appendAndSetNodeAttribute(*m_fontNode, "scheme", "val", XLFontSchemeStyleToString(newScheme)).empty() == false; }
bool XLFont::setVertAlign(XLVerticalAlignRunStyle newVertAlign)
{
    return appendAndSetNodeAttribute(*m_fontNode, "vertAlign", "val", XLVerticalAlignRunStyleToString(newVertAlign)).empty() ==
           false;
}
bool XLFont::setOutline(bool set)
{ return appendAndSetNodeAttribute(*m_fontNode, "outline", "val", (set ? "true" : "false")).empty() == false; }
bool XLFont::setShadow(bool set)
{ return appendAndSetNodeAttribute(*m_fontNode, "shadow", "val", (set ? "true" : "false")).empty() == false; }
bool XLFont::setCondense(bool set)
{ return appendAndSetNodeAttribute(*m_fontNode, "condense", "val", (set ? "true" : "false")).empty() == false; }
bool XLFont::setExtend(bool set)
{ return appendAndSetNodeAttribute(*m_fontNode, "extend", "val", (set ? "true" : "false")).empty() == false; }

/**
 * @details assemble a string summary about the font
 */
std::string XLFont::summary() const
{
    return fmt::format("font name is {}, charset: {}, font family: {}, size: {}, color: {}{}{}{}{}{}{}{}{}{}{}",
                       fontName(),
                       fontCharset(),
                       fontFamily(),
                       fontSize(),
                       fontColor().hex(),
                       bold() ? ", +bold" : "",
                       italic() ? ", +italic" : "",
                       strikethrough() ? ", +strikethrough" : "",
                       underline() != XLUnderlineNone ? fmt::format(", underline: {}", XLUnderlineStyleToString(underline())) : "",
                       scheme() != XLFontSchemeNone ? fmt::format(", scheme: {}", XLFontSchemeStyleToString(scheme())) : "",
                       vertAlign() != XLBaseline ? fmt::format(", vertAlign: {}", XLVerticalAlignRunStyleToString(vertAlign())) : "",
                       outline() ? ", +outline" : "",
                       shadow() ? ", +shadow" : "",
                       condense() ? ", +condense" : "",
                       extend() ? ", +extend" : "");
}

// ===== XLFonts, parent of XLFont

/**
 * @details Constructor. Initializes an empty XLFonts object
 */
XLFonts::XLFonts() : m_fontsNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLFonts object.
 */
XLFonts::XLFonts(const XMLNode& fonts) : m_fontsNode(std::make_unique<XMLNode>(fonts))
{
    // initialize XLFonts entries and m_fonts here
    XMLNode node = m_fontsNode->first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        if (nodeName == "font")
            m_fonts.push_back(XLFont(node));
        else
            std::cerr << "WARNING: XLFonts constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLFonts::~XLFonts()
{
    m_fonts.clear();    // delete vector with all children
}

XLFonts::XLFonts(const XLFonts& other) : m_fontsNode(std::make_unique<XMLNode>(*other.m_fontsNode)), m_fonts(other.m_fonts) {}

XLFonts::XLFonts(XLFonts&& other) : m_fontsNode(std::move(other.m_fontsNode)), m_fonts(std::move(other.m_fonts)) {}

/**
 * @details Copy assignment operator
 */
XLFonts& XLFonts::operator=(const XLFonts& other)
{
    if (&other != this) {
        *m_fontsNode = *other.m_fontsNode;
        m_fonts.clear();
        m_fonts = other.m_fonts;
    }
    return *this;
}

/**
 * @details Returns the amount of fonts held by the class
 */
size_t XLFonts::count() const { return m_fonts.size(); }

/**
 * @details fetch XLFont from m_Fonts by index
 */
XLFont XLFonts::fontByIndex(XLStyleIndex index) const
{
    Expects(index < m_fonts.size());
    return m_fonts.at(index);
}

/**
 * @details append a new XLFont to m_fonts and m_fontsNode, based on copyFrom
 */
XLStyleIndex XLFonts::create(XLFont copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the font to be created
    XMLNode      newNode{};          // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_fontsNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty())
        newNode = m_fontsNode->prepend_child("font");
    else
        newNode = m_fontsNode->insert_child_after("font", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLFonts::"s + __func__ + ": failed to append a new fonts node"s);
    }
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_fontsNode->insert_child_before(pugi::node_pcdata, newNode)
            .set_value(std::string(styleEntriesPrefix).c_str());    // prefix the new node with styleEntriesPrefix

    XLFont newFont(newNode);
    if (copyFrom.m_fontNode->empty()) {    // if no template is given
        // ===== Create a font with default values
        newFont.setFontName(OpenXLSX::XLDefaultFontName);
        newFont.setFontSize(OpenXLSX::XLDefaultFontSize);
        newFont.setFontColor(XLColor(OpenXLSX::XLDefaultFontColor));
        newFont.setFontFamily(OpenXLSX::XLDefaultFontFamily);
        newFont.setFontCharset(OpenXLSX::XLDefaultFontCharset);
    }
    else
        copyXMLNode(newNode, *copyFrom.m_fontNode);    // will use copyFrom as template, does nothing if copyFrom is empty

    m_fonts.push_back(newFont);
    appendAndSetAttribute(*m_fontsNode, "count", std::to_string(m_fonts.size()));    // update array count in XML
    return index;
}

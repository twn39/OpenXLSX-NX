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

XLFont::XLFont() : m_fontNode(std::make_unique<XMLNode>()) {}

XLFont::XLFont(const XMLNode& node) : m_fontNode(std::make_unique<XMLNode>(node)) {}

XLFont::~XLFont() = default;

XLFont::XLFont(const XLFont& other) : m_fontNode(std::make_unique<XMLNode>(*other.m_fontNode)) {}

XLFont& XLFont::operator=(const XLFont& other)
{
    if (&other != this) *m_fontNode = *other.m_fontNode;
    return *this;
}

std::string XLFont::fontName() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "name", "val", OpenXLSX::XLDefaultFontName);
    return attr.value();
}

size_t XLFont::fontCharset() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "charset", "val", std::to_string(OpenXLSX::XLDefaultFontCharset));
    return attr.as_uint();
}

size_t XLFont::fontFamily() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "family", "val", std::to_string(OpenXLSX::XLDefaultFontFamily));
    return attr.as_uint();
}

size_t XLFont::fontSize() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "sz", "val", std::to_string(OpenXLSX::XLDefaultFontSize));
    return attr.as_uint();
}

XLColor XLFont::fontColor() const
{
    using namespace std::literals::string_literals;
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "color", "rgb", XLDefaultFontColor);
    return XLColor(attr.value());
}

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

XLFont& XLFont::setFontName(std::string_view newName)
{ appendAndSetNodeAttribute(*m_fontNode, "name", "val", std::string(newName).c_str()).empty(); return *this; }
XLFont& XLFont::setFontCharset(size_t newCharset)
{ appendAndSetNodeAttribute(*m_fontNode, "charset", "val", std::to_string(newCharset)).empty(); return *this; }
XLFont& XLFont::setFontFamily(size_t newFamily)
{ appendAndSetNodeAttribute(*m_fontNode, "family", "val", std::to_string(newFamily)).empty(); return *this; }
XLFont& XLFont::setFontSize(size_t newSize)
{ appendAndSetNodeAttribute(*m_fontNode, "sz", "val", std::to_string(newSize)).empty(); return *this; }
XLFont& XLFont::setFontColor(XLColor newColor)
{ appendAndSetNodeAttribute(*m_fontNode, "color", "rgb", newColor.hex(), XLRemoveAttributes).empty(); return *this; }
XLFont& XLFont::setBold(bool set) { appendAndSetNodeAttribute(*m_fontNode, "b", "val", (set ? "1" : "0")).empty(); return *this; }
XLFont& XLFont::setItalic(bool set) { appendAndSetNodeAttribute(*m_fontNode, "i", "val", (set ? "1" : "0")).empty(); return *this; }
XLFont& XLFont::setStrikethrough(bool set)
{ appendAndSetNodeAttribute(*m_fontNode, "strike", "val", (set ? "1" : "0")).empty(); return *this; }
XLFont& XLFont::setUnderline(XLUnderlineStyle style)
{ appendAndSetNodeAttribute(*m_fontNode, "u", "val", XLUnderlineStyleToString(style)).empty(); return *this; }
XLFont& XLFont::setScheme(XLFontSchemeStyle newScheme)
{ appendAndSetNodeAttribute(*m_fontNode, "scheme", "val", XLFontSchemeStyleToString(newScheme)).empty(); return *this; }
XLFont& XLFont::setVertAlign(XLVerticalAlignRunStyle newVertAlign)
{
    appendAndSetNodeAttribute(*m_fontNode, "vertAlign", "val", XLVerticalAlignRunStyleToString(newVertAlign)).empty(); return *this;
}
XLFont& XLFont::setOutline(bool set)
{ appendAndSetNodeAttribute(*m_fontNode, "outline", "val", (set ? "true" : "false")).empty(); return *this; }
XLFont& XLFont::setShadow(bool set)
{ appendAndSetNodeAttribute(*m_fontNode, "shadow", "val", (set ? "true" : "false")).empty(); return *this; }
XLFont& XLFont::setCondense(bool set)
{ appendAndSetNodeAttribute(*m_fontNode, "condense", "val", (set ? "true" : "false")).empty(); return *this; }
XLFont& XLFont::setExtend(bool set)
{ appendAndSetNodeAttribute(*m_fontNode, "extend", "val", (set ? "true" : "false")).empty(); return *this; }

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

XLFonts::XLFonts() : m_fontsNode(std::make_unique<XMLNode>()) {}

XLFonts::XLFonts(const XMLNode& fonts) : m_fontsNode(std::make_unique<XMLNode>(fonts))
{
    // initialize XLFonts entries and m_fonts here
    XMLNode node = m_fontsNode->first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string_view nodeName(node.name());
        if (nodeName == "font")
            m_fonts.push_back(XLFont(node));
        else
            std::cerr << "WARNING: XLFonts constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLFonts::~XLFonts()
{
    m_fonts.clear();   
}

XLFonts::XLFonts(const XLFonts& other) : m_fontsNode(std::make_unique<XMLNode>(*other.m_fontsNode)), m_fonts(other.m_fonts) {}

XLFonts::XLFonts(XLFonts&& other) : m_fontsNode(std::move(other.m_fontsNode)), m_fonts(std::move(other.m_fonts)) {}

XLFonts& XLFonts::operator=(const XLFonts& other)
{
    if (&other != this) {
        *m_fontsNode = *other.m_fontsNode;
        m_fonts.clear();
        m_fonts = other.m_fonts;
    }
    return *this;
}

size_t XLFonts::count() const { return m_fonts.size(); }

XLFont XLFonts::fontByIndex(XLStyleIndex index) const
{
    Expects(index < m_fonts.size());
    return m_fonts.at(index);
}

XLStyleIndex XLFonts::create(XLFont copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();   
    XMLNode      newNode{};         

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
    if (styleEntriesPrefix.length() > 0)   
        m_fontsNode->insert_child_before(pugi::node_pcdata, newNode)
            .set_value(std::string(styleEntriesPrefix).c_str());   

    XLFont newFont(newNode);
    if (copyFrom.m_fontNode->empty()) {   
        // ===== Create a font with default values
        newFont.setFontName(OpenXLSX::XLDefaultFontName);
        newFont.setFontSize(OpenXLSX::XLDefaultFontSize);
        newFont.setFontColor(XLColor(OpenXLSX::XLDefaultFontColor));
        newFont.setFontFamily(OpenXLSX::XLDefaultFontFamily);
        newFont.setFontCharset(OpenXLSX::XLDefaultFontCharset);
    }
    else
        copyXMLNode(newNode, *copyFrom.m_fontNode);   

    m_fonts.push_back(newFont);
    appendAndSetAttribute(*m_fontsNode, "count", std::to_string(m_fonts.size()));   
    return index;
}

XLStyleIndex XLFonts::findOrCreate(XLFont copyFrom, std::string_view styleEntriesPrefix)
{
    // Compute the fingerprint of the requested font
    std::string key = xmlNodeFingerprint(*copyFrom.m_fontNode);

    // Fast path: cache hit
    auto it = m_fingerprintCache.find(key);
    if (it != m_fingerprintCache.end()) return it->second;

    // Cold path: scan all existing fonts (covers fonts loaded from an existing file)
    for (size_t i = 0; i < m_fonts.size(); ++i) {
        if (xmlNodeFingerprint(*m_fonts[i].m_fontNode) == key) {
            m_fingerprintCache.emplace(key, i);
            return i;
        }
    }

    // No match found — create and cache
    XLStyleIndex idx = create(copyFrom, styleEntriesPrefix);
    m_fingerprintCache.emplace(key, idx);
    return idx;
}

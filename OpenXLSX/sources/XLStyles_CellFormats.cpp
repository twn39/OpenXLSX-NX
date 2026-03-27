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
    constexpr EnumStringMap<XLAlignmentStyle> XLAlignmentStyleMap[] = {
        {"", XLAlignGeneral},
        {"left", XLAlignLeft},
        {"right", XLAlignRight},
        {"center", XLAlignCenter},
        {"top", XLAlignTop},
        {"bottom", XLAlignBottom},
        {"fill", XLAlignFill},
        {"justify", XLAlignJustify},
        {"centerContinuous", XLAlignCenterContinuous},
        {"distributed", XLAlignDistributed},
        {"general", XLAlignGeneral},
    };

    XLAlignmentStyle XLAlignmentStyleFromString(std::string_view str) {
        return EnumFromString(str, XLAlignmentStyleMap, XLAlignInvalid);
    }

    const char* XLAlignmentStyleToString(XLAlignmentStyle val) {
        return EnumToString(val, XLAlignmentStyleMap, "(unknown)");
    }

    const char* XLReadingOrderToString(uint32_t readingOrder) {
        switch (readingOrder) {
            case XLReadingOrderContextual:
                return "contextual";
            case XLReadingOrderLeftToRight:
                return "left-to-right";
            case XLReadingOrderRightToLeft:
                return "right-to-left";
            default:
                return "(unknown)";
        }
    }
}

XLAlignment::XLAlignment() : m_alignmentNode(std::make_unique<XMLNode>()) {}

XLAlignment::XLAlignment(const XMLNode& node) : m_alignmentNode(std::make_unique<XMLNode>(node)) {}

XLAlignment::~XLAlignment() = default;

XLAlignment::XLAlignment(const XLAlignment& other) : m_alignmentNode(std::make_unique<XMLNode>(*other.m_alignmentNode)) {}

XLAlignment& XLAlignment::operator=(const XLAlignment& other)
{
    if (&other != this) *m_alignmentNode = *other.m_alignmentNode;
    return *this;
}

XLAlignmentStyle XLAlignment::horizontal() const { return XLAlignmentStyleFromString(m_alignmentNode->attribute("horizontal").value()); }

XLAlignmentStyle XLAlignment::vertical() const { return XLAlignmentStyleFromString(m_alignmentNode->attribute("vertical").value()); }

uint16_t XLAlignment::textRotation() const { return gsl::narrow_cast<uint16_t>(m_alignmentNode->attribute("textRotation").as_uint()); }

bool XLAlignment::wrapText() const { return m_alignmentNode->attribute("wrapText").as_bool(); }

uint32_t XLAlignment::indent() const { return m_alignmentNode->attribute("indent").as_uint(); }

int32_t XLAlignment::relativeIndent() const { return m_alignmentNode->attribute("relativeIndent").as_int(); }

bool XLAlignment::justifyLastLine() const { return m_alignmentNode->attribute("justifyLastLine").as_bool(); }

bool XLAlignment::shrinkToFit() const { return m_alignmentNode->attribute("shrinkToFit").as_bool(); }

uint32_t XLAlignment::readingOrder() const { return m_alignmentNode->attribute("readingOrder").as_uint(); }

bool XLAlignment::setHorizontal(XLAlignmentStyle newStyle)
{ return appendAndSetAttribute(*m_alignmentNode, "horizontal", XLAlignmentStyleToString(newStyle)).empty() == false; }
bool XLAlignment::setVertical(XLAlignmentStyle newStyle)
{ return appendAndSetAttribute(*m_alignmentNode, "vertical", XLAlignmentStyleToString(newStyle)).empty() == false; }
bool XLAlignment::setTextRotation(uint16_t newRotation)
{ return appendAndSetAttribute(*m_alignmentNode, "textRotation", std::to_string(newRotation)).empty() == false; }
bool XLAlignment::setWrapText(bool set)
{ return appendAndSetAttribute(*m_alignmentNode, "wrapText", (set ? "true" : "false")).empty() == false; }
bool XLAlignment::setIndent(uint32_t newIndent)
{ return appendAndSetAttribute(*m_alignmentNode, "indent", std::to_string(newIndent)).empty() == false; }
bool XLAlignment::setRelativeIndent(int32_t newRelativeIndent)
{ return appendAndSetAttribute(*m_alignmentNode, "relativeIndent", std::to_string(newRelativeIndent)).empty() == false; }
bool XLAlignment::setJustifyLastLine(bool set)
{ return appendAndSetAttribute(*m_alignmentNode, "justifyLastLine", (set ? "true" : "false")).empty() == false; }
bool XLAlignment::setShrinkToFit(bool set)
{ return appendAndSetAttribute(*m_alignmentNode, "shrinkToFit", (set ? "true" : "false")).empty() == false; }
bool XLAlignment::setReadingOrder(uint32_t newReadingOrder)
{ return appendAndSetAttribute(*m_alignmentNode, "readingOrder", std::to_string(newReadingOrder)).empty() == false; }

std::string XLAlignment::summary() const
{
    return fmt::format("alignment horizontal={}, vertical={}, textRotation={}, wrapText={}, indent={}, relativeIndent={}, justifyLastLine={}, shrinkToFit={}, readingOrder={}",
                       XLAlignmentStyleToString(horizontal()),
                       XLAlignmentStyleToString(vertical()),
                       textRotation(),
                       wrapText() ? "true" : "false",
                       indent(),
                       relativeIndent(),
                       justifyLastLine() ? "true" : "false",
                       shrinkToFit() ? "true" : "false",
                       XLReadingOrderToString(readingOrder()));
}

XLCellFormat::XLCellFormat() : m_cellFormatNode(std::make_unique<XMLNode>()) {}

XLCellFormat::XLCellFormat(const XMLNode& node, bool permitXfId)
    : m_cellFormatNode(std::make_unique<XMLNode>(node)),
      m_permitXfId(permitXfId)
{}

XLCellFormat::~XLCellFormat() = default;

XLCellFormat::XLCellFormat(const XLCellFormat& other)
    : m_cellFormatNode(std::make_unique<XMLNode>(*other.m_cellFormatNode)),
      m_permitXfId(other.m_permitXfId)
{}

XLCellFormat& XLCellFormat::operator=(const XLCellFormat& other)
{
    if (&other != this) {
        *m_cellFormatNode = *other.m_cellFormatNode;
        m_permitXfId      = other.m_permitXfId;
    }
    return *this;
}

/**
 * @details determines the numberFormatId
 * @note returns XLInvalidUInt32 if attribute is not defined / set / empty
 */
uint32_t XLCellFormat::numberFormatId() const { return m_cellFormatNode->attribute("numFmtId").as_uint(XLInvalidUInt32); }

/**
 * @details determines the fontIndex
 * @note returns XLInvalidStyleIndex if attribute is not defined / set / empty
 */
XLStyleIndex XLCellFormat::fontIndex() const { return m_cellFormatNode->attribute("fontId").as_uint(XLInvalidStyleIndex); }

/**
 * @details determines the fillIndex
 * @note returns XLInvalidStyleIndex if attribute is not defined / set / empty
 */
XLStyleIndex XLCellFormat::fillIndex() const { return m_cellFormatNode->attribute("fillId").as_uint(XLInvalidStyleIndex); }

/**
 * @details determines the borderIndex
 * @note returns XLInvalidStyleIndex if attribute is not defined / set / empty
 */
XLStyleIndex XLCellFormat::borderIndex() const { return m_cellFormatNode->attribute("borderId").as_uint(XLInvalidStyleIndex); }

XLStyleIndex XLCellFormat::xfId() const
{
    if (m_permitXfId) return m_cellFormatNode->attribute("xfId").as_uint(XLInvalidStyleIndex);
    throw XLException("XLCellFormat::xfId not permitted when m_permitXfId is false");
}

bool XLCellFormat::applyNumberFormat() const { return m_cellFormatNode->attribute("applyNumberFormat").as_bool(); }
bool XLCellFormat::applyFont() const { return m_cellFormatNode->attribute("applyFont").as_bool(); }
bool XLCellFormat::applyFill() const { return m_cellFormatNode->attribute("applyFill").as_bool(); }
bool XLCellFormat::applyBorder() const { return m_cellFormatNode->attribute("applyBorder").as_bool(); }
bool XLCellFormat::applyAlignment() const { return m_cellFormatNode->attribute("applyAlignment").as_bool(); }
bool XLCellFormat::applyProtection() const { return m_cellFormatNode->attribute("applyProtection").as_bool(); }

bool XLCellFormat::quotePrefix() const { return m_cellFormatNode->attribute("quotePrefix").as_bool(); }
bool XLCellFormat::pivotButton() const { return m_cellFormatNode->attribute("pivotButton").as_bool(); }

bool XLCellFormat::locked() const { return m_cellFormatNode->child("protection").attribute("locked").as_bool(); }
bool XLCellFormat::hidden() const { return m_cellFormatNode->child("protection").attribute("hidden").as_bool(); }

XLAlignment XLCellFormat::alignment(bool createIfMissing) const
{
    XMLNode nodeAlignment = m_cellFormatNode->child("alignment");
    if (nodeAlignment.empty() and createIfMissing)
        nodeAlignment =
            appendAndGetNode(*m_cellFormatNode, "alignment", m_nodeOrder);    // 2024-12-19: ordered insert to address issue #305
    return XLAlignment(nodeAlignment);
}

XLCellFormat& XLCellFormat::setNumberFormatId(uint32_t newNumFmtId)
{ appendAndSetAttribute(*m_cellFormatNode, "numFmtId", std::to_string(newNumFmtId)).empty(); return *this; }
XLCellFormat& XLCellFormat::setFontIndex(XLStyleIndex newXfIndex)
{ appendAndSetAttribute(*m_cellFormatNode, "fontId", std::to_string(newXfIndex)).empty(); return *this; }
XLCellFormat& XLCellFormat::setFillIndex(XLStyleIndex newFillIndex)
{ appendAndSetAttribute(*m_cellFormatNode, "fillId", std::to_string(newFillIndex)).empty(); return *this; }
XLCellFormat& XLCellFormat::setBorderIndex(XLStyleIndex newBorderIndex)
{ appendAndSetAttribute(*m_cellFormatNode, "borderId", std::to_string(newBorderIndex)).empty(); return *this; }
XLCellFormat& XLCellFormat::setXfId(XLStyleIndex newXfId)
{
    if (m_permitXfId) {
        appendAndSetAttribute(*m_cellFormatNode, "xfId", std::to_string(newXfId)).empty();
        return *this;
    }
    throw XLException("XLCellFormat::setXfId not permitted when m_permitXfId is false");
}
XLCellFormat& XLCellFormat::setApplyNumberFormat(bool set)
{ appendAndSetAttribute(*m_cellFormatNode, "applyNumberFormat", (set ? "true" : "false")).empty(); return *this; }
XLCellFormat& XLCellFormat::setApplyFont(bool set)
{ appendAndSetAttribute(*m_cellFormatNode, "applyFont", (set ? "true" : "false")).empty(); return *this; }
XLCellFormat& XLCellFormat::setApplyFill(bool set)
{ appendAndSetAttribute(*m_cellFormatNode, "applyFill", (set ? "true" : "false")).empty(); return *this; }
XLCellFormat& XLCellFormat::setApplyBorder(bool set)
{ appendAndSetAttribute(*m_cellFormatNode, "applyBorder", (set ? "true" : "false")).empty(); return *this; }
XLCellFormat& XLCellFormat::setApplyAlignment(bool set)
{ appendAndSetAttribute(*m_cellFormatNode, "applyAlignment", (set ? "true" : "false")).empty(); return *this; }
XLCellFormat& XLCellFormat::setApplyProtection(bool set)
{ appendAndSetAttribute(*m_cellFormatNode, "applyProtection", (set ? "true" : "false")).empty(); return *this; }
XLCellFormat& XLCellFormat::setQuotePrefix(bool set)
{ appendAndSetAttribute(*m_cellFormatNode, "quotePrefix", (set ? "true" : "false")).empty(); return *this; }
XLCellFormat& XLCellFormat::setPivotButton(bool set)
{ appendAndSetAttribute(*m_cellFormatNode, "pivotButton", (set ? "true" : "false")).empty(); return *this; }
XLCellFormat& XLCellFormat::setLocked(bool set)
{
    appendAndSetNodeAttribute(*m_cellFormatNode,
                                     "protection",
                                     "locked",
                                     (set ? "true" : "false"),
                                     /**/ XLKeepAttributes,
                                     m_nodeOrder)
               .empty(); return *this;    // 2024-12-19: ordered insert to address issue #305
}
XLCellFormat& XLCellFormat::setHidden(bool set)
{
    appendAndSetNodeAttribute(*m_cellFormatNode,
                                     "protection",
                                     "hidden",
                                     (set ? "true" : "false"),
                                     /**/ XLKeepAttributes,
                                     m_nodeOrder)
               .empty(); return *this;    // 2024-12-19: ordered insert to address issue #305
}

/**
 * @brief Unsupported setter function
 */
XLCellFormat& XLCellFormat::setExtLst(XLUnsupportedElement const& newExtLst)
{
    OpenXLSX::ignore(newExtLst);
    return *this;
}

std::string XLCellFormat::summary() const
{
    return fmt::format("numberFormatId={}, fontIndex={}, fillIndex={}, borderIndex={}{}, applyNumberFormat: {}, applyFont: {}, applyFill: {}, applyBorder: {}, applyAlignment: {}, applyProtection: {}, quotePrefix: {}, pivotButton: {}, locked: {}, hidden: {}, {}",
                       numberFormatId(),
                       fontIndex(),
                       fillIndex(),
                       borderIndex(),
                       m_permitXfId ? fmt::format(", xfId: {}", xfId()) : "",
                       applyNumberFormat() ? "true" : "false",
                       applyFont() ? "true" : "false",
                       applyFill() ? "true" : "false",
                       applyBorder() ? "true" : "false",
                       applyAlignment() ? "true" : "false",
                       applyProtection() ? "true" : "false",
                       quotePrefix() ? "true" : "false",
                       pivotButton() ? "true" : "false",
                       locked() ? "true" : "false",
                       hidden() ? "true" : "false",
                       alignment().summary());
}

// ===== XLCellFormats, one parent of XLCellFormat (the other being XLCellFormats)

XLCellFormats::XLCellFormats() : m_cellFormatsNode(std::make_unique<XMLNode>()) {}

XLCellFormats::XLCellFormats(const XMLNode& cellStyleFormats, bool permitXfId)
    : m_cellFormatsNode(std::make_unique<XMLNode>(cellStyleFormats)),
      m_permitXfId(permitXfId)
{
    // initialize XLCellFormats entries and m_cellFormats here
    XMLNode node = cellStyleFormats.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string_view nodeName(node.name());
        if (nodeName == "xf")
            m_cellFormats.push_back(XLCellFormat(node, m_permitXfId));
        else
            std::cerr << "WARNING: XLCellFormats constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLCellFormats::~XLCellFormats()
{
    m_cellFormats.clear();   
}

XLCellFormats::XLCellFormats(const XLCellFormats& other)
    : m_cellFormatsNode(std::make_unique<XMLNode>(*other.m_cellFormatsNode)),
      m_cellFormats(other.m_cellFormats),
      m_permitXfId(other.m_permitXfId)
{}

XLCellFormats::XLCellFormats(XLCellFormats&& other)
    : m_cellFormatsNode(std::move(other.m_cellFormatsNode)),
      m_cellFormats(std::move(other.m_cellFormats)),
      m_permitXfId(other.m_permitXfId)
{}

XLCellFormats& XLCellFormats::operator=(const XLCellFormats& other)
{
    if (&other != this) {
        *m_cellFormatsNode = *other.m_cellFormatsNode;
        m_cellFormats.clear();
        m_cellFormats = other.m_cellFormats;
        m_permitXfId  = other.m_permitXfId;
    }
    return *this;
}

size_t XLCellFormats::count() const { return m_cellFormats.size(); }

XLCellFormat XLCellFormats::cellFormatByIndex(XLStyleIndex index) const
{
    Expects(index < m_cellFormats.size());
    return m_cellFormats.at(index);
}

XLStyleIndex XLCellFormats::create(XLCellFormat copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();   
    XMLNode      newNode{};         

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_cellFormatsNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty())
        newNode = m_cellFormatsNode->prepend_child("xf");
    else
        newNode = m_cellFormatsNode->insert_child_after("xf", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCellFormats::"s + __func__ + ": failed to append a new xf node"s);
    }
    if (styleEntriesPrefix.length() > 0)   
        m_cellFormatsNode->insert_child_before(pugi::node_pcdata, newNode)
            .set_value(std::string(styleEntriesPrefix).c_str());   

    XLCellFormat newCellFormat(newNode, m_permitXfId);
    if (copyFrom.m_cellFormatNode->empty()) {   
        // ===== Create a cell format with default values
        // default index 0 for other style elements should protect from exceptions
        newCellFormat.setNumberFormatId(0);
        newCellFormat.setFontIndex(0);
        newCellFormat.setFillIndex(0);
        newCellFormat.setBorderIndex(0);
        if (m_permitXfId) {
            newCellFormat.setXfId(0);
            newCellFormat.setQuotePrefix(false);
            newCellFormat.setPivotButton(false);
        }
        newCellFormat.setApplyNumberFormat(false);
        newCellFormat.setApplyFont(false);
        newCellFormat.setApplyFill(false);
        newCellFormat.setApplyBorder(false);
        newCellFormat.setApplyAlignment(false);
        newCellFormat.setApplyProtection(false);
        newCellFormat.setLocked(false);
        newCellFormat.setHidden(false);
        // Unsupported setter
        newCellFormat.setExtLst(XLUnsupportedElement{});
    }
    else
        copyXMLNode(newNode, *copyFrom.m_cellFormatNode);   

    m_cellFormats.push_back(newCellFormat);
    appendAndSetAttribute(*m_cellFormatsNode, "count", std::to_string(m_cellFormats.size()));   
    return index;
}

XLStyleIndex XLCellFormats::findOrCreate(XLCellFormat copyFrom, std::string_view styleEntriesPrefix)
{
    // Compute the fingerprint of the requested cell format
    std::string key = xmlNodeFingerprint(*copyFrom.m_cellFormatNode);

    // Fast path: cache hit
    auto it = m_fingerprintCache.find(key);
    if (it != m_fingerprintCache.end()) return it->second;

    // Cold path: scan all existing formats (covers formats loaded from an existing file).
    // Start from index 1: index 0 is the reserved default format, never reused.
    for (size_t i = 1; i < m_cellFormats.size(); ++i) {
        if (xmlNodeFingerprint(*m_cellFormats[i].m_cellFormatNode) == key) {
            m_fingerprintCache.emplace(key, i);
            return i;
        }
    }

    // No match found — create and cache
    XLStyleIndex idx = create(copyFrom, styleEntriesPrefix);
    m_fingerprintCache.emplace(key, idx);
    return idx;
}

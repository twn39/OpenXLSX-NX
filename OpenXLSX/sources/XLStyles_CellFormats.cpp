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

/**
 * @details Constructor. Initializes an empty XLAlignment object
 */
XLAlignment::XLAlignment() : m_alignmentNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLAlignment object.
 */
XLAlignment::XLAlignment(const XMLNode& node) : m_alignmentNode(std::make_unique<XMLNode>(node)) {}

XLAlignment::~XLAlignment() = default;

XLAlignment::XLAlignment(const XLAlignment& other) : m_alignmentNode(std::make_unique<XMLNode>(*other.m_alignmentNode)) {}

XLAlignment& XLAlignment::operator=(const XLAlignment& other)
{
    if (&other != this) *m_alignmentNode = *other.m_alignmentNode;
    return *this;
}

/**
 * @details Returns the horizontal alignment style
 */
XLAlignmentStyle XLAlignment::horizontal() const { return XLAlignmentStyleFromString(m_alignmentNode->attribute("horizontal").value()); }

/**
 * @details Returns the vertical alignment style
 */
XLAlignmentStyle XLAlignment::vertical() const { return XLAlignmentStyleFromString(m_alignmentNode->attribute("vertical").value()); }

/**
 * @details Returns the text rotation
 */
uint16_t XLAlignment::textRotation() const { return gsl::narrow_cast<uint16_t>(m_alignmentNode->attribute("textRotation").as_uint()); }

/**
 * @details check if text wrapping is enabled
 */
bool XLAlignment::wrapText() const { return m_alignmentNode->attribute("wrapText").as_bool(); }

/**
 * @details Returns the indent setting
 */
uint32_t XLAlignment::indent() const { return m_alignmentNode->attribute("indent").as_uint(); }

/**
 * @details Returns the relative indent setting
 */
int32_t XLAlignment::relativeIndent() const { return m_alignmentNode->attribute("relativeIndent").as_int(); }

/**
 * @details check if justification of last line is enabled
 */
bool XLAlignment::justifyLastLine() const { return m_alignmentNode->attribute("justifyLastLine").as_bool(); }

/**
 * @details check if shrink to fit is enabled
 */
bool XLAlignment::shrinkToFit() const { return m_alignmentNode->attribute("shrinkToFit").as_bool(); }

/**
 * @details Returns the reading order setting
 */
uint32_t XLAlignment::readingOrder() const { return m_alignmentNode->attribute("readingOrder").as_uint(); }

/**
 * @details Setter functions
 */
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

/**
 * @details assemble a string summary about the fill
 */
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

/**
 * @details Constructor. Initializes an empty XLCellFormat object
 */
XLCellFormat::XLCellFormat() : m_cellFormatNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLCellFormat object.
 */
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

/**
 * @details Returns the xfId value
 */
XLStyleIndex XLCellFormat::xfId() const
{
    if (m_permitXfId) return m_cellFormatNode->attribute("xfId").as_uint(XLInvalidStyleIndex);
    throw XLException("XLCellFormat::xfId not permitted when m_permitXfId is false");
}

/**
 * @details determine the applyNumberFormat,applyFont,applyFill,applyBorder,applyAlignment,applyProtection status
 */
bool XLCellFormat::applyNumberFormat() const { return m_cellFormatNode->attribute("applyNumberFormat").as_bool(); }
bool XLCellFormat::applyFont() const { return m_cellFormatNode->attribute("applyFont").as_bool(); }
bool XLCellFormat::applyFill() const { return m_cellFormatNode->attribute("applyFill").as_bool(); }
bool XLCellFormat::applyBorder() const { return m_cellFormatNode->attribute("applyBorder").as_bool(); }
bool XLCellFormat::applyAlignment() const { return m_cellFormatNode->attribute("applyAlignment").as_bool(); }
bool XLCellFormat::applyProtection() const { return m_cellFormatNode->attribute("applyProtection").as_bool(); }

/**
 * @details determine the quotePrefix, pivotButton status
 */
bool XLCellFormat::quotePrefix() const { return m_cellFormatNode->attribute("quotePrefix").as_bool(); }
bool XLCellFormat::pivotButton() const { return m_cellFormatNode->attribute("pivotButton").as_bool(); }

/**
 * @details determine the protection:locked,protection:hidden status
 */
bool XLCellFormat::locked() const { return m_cellFormatNode->child("protection").attribute("locked").as_bool(); }
bool XLCellFormat::hidden() const { return m_cellFormatNode->child("protection").attribute("hidden").as_bool(); }

/**
 * @details fetch alignment object
 */
XLAlignment XLCellFormat::alignment(bool createIfMissing) const
{
    XMLNode nodeAlignment = m_cellFormatNode->child("alignment");
    if (nodeAlignment.empty() and createIfMissing)
        nodeAlignment =
            appendAndGetNode(*m_cellFormatNode, "alignment", m_nodeOrder);    // 2024-12-19: ordered insert to address issue #305
    return XLAlignment(nodeAlignment);
}

/**
 * @details Setter functions
 */
bool XLCellFormat::setNumberFormatId(uint32_t newNumFmtId)
{ return appendAndSetAttribute(*m_cellFormatNode, "numFmtId", std::to_string(newNumFmtId)).empty() == false; }
bool XLCellFormat::setFontIndex(XLStyleIndex newXfIndex)
{ return appendAndSetAttribute(*m_cellFormatNode, "fontId", std::to_string(newXfIndex)).empty() == false; }
bool XLCellFormat::setFillIndex(XLStyleIndex newFillIndex)
{ return appendAndSetAttribute(*m_cellFormatNode, "fillId", std::to_string(newFillIndex)).empty() == false; }
bool XLCellFormat::setBorderIndex(XLStyleIndex newBorderIndex)
{ return appendAndSetAttribute(*m_cellFormatNode, "borderId", std::to_string(newBorderIndex)).empty() == false; }
bool XLCellFormat::setXfId(XLStyleIndex newXfId)
{
    if (m_permitXfId) return appendAndSetAttribute(*m_cellFormatNode, "xfId", std::to_string(newXfId)).empty() == false;
    throw XLException("XLCellFormat::setXfId not permitted when m_permitXfId is false");
}
bool XLCellFormat::setApplyNumberFormat(bool set)
{ return appendAndSetAttribute(*m_cellFormatNode, "applyNumberFormat", (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setApplyFont(bool set)
{ return appendAndSetAttribute(*m_cellFormatNode, "applyFont", (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setApplyFill(bool set)
{ return appendAndSetAttribute(*m_cellFormatNode, "applyFill", (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setApplyBorder(bool set)
{ return appendAndSetAttribute(*m_cellFormatNode, "applyBorder", (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setApplyAlignment(bool set)
{ return appendAndSetAttribute(*m_cellFormatNode, "applyAlignment", (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setApplyProtection(bool set)
{ return appendAndSetAttribute(*m_cellFormatNode, "applyProtection", (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setQuotePrefix(bool set)
{ return appendAndSetAttribute(*m_cellFormatNode, "quotePrefix", (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setPivotButton(bool set)
{ return appendAndSetAttribute(*m_cellFormatNode, "pivotButton", (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setLocked(bool set)
{
    return appendAndSetNodeAttribute(*m_cellFormatNode,
                                     "protection",
                                     "locked",
                                     (set ? "true" : "false"),
                                     /**/ XLKeepAttributes,
                                     m_nodeOrder)
               .empty() == false;    // 2024-12-19: ordered insert to address issue #305
}
bool XLCellFormat::setHidden(bool set)
{
    return appendAndSetNodeAttribute(*m_cellFormatNode,
                                     "protection",
                                     "hidden",
                                     (set ? "true" : "false"),
                                     /**/ XLKeepAttributes,
                                     m_nodeOrder)
               .empty() == false;    // 2024-12-19: ordered insert to address issue #305
}

/**
 * @brief Unsupported setter function
 */
bool XLCellFormat::setExtLst(XLUnsupportedElement const& newExtLst)
{
    OpenXLSX::ignore(newExtLst);
    return false;
}

/**
 * @details assemble a string summary about the cell format
 */
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

/**
 * @details Constructor. Initializes an empty XLCellFormats object
 */
XLCellFormats::XLCellFormats() : m_cellFormatsNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLCellFormats object.
 */
XLCellFormats::XLCellFormats(const XMLNode& cellStyleFormats, bool permitXfId)
    : m_cellFormatsNode(std::make_unique<XMLNode>(cellStyleFormats)),
      m_permitXfId(permitXfId)
{
    // initialize XLCellFormats entries and m_cellFormats here
    XMLNode node = cellStyleFormats.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        if (nodeName == "xf")
            m_cellFormats.push_back(XLCellFormat(node, m_permitXfId));
        else
            std::cerr << "WARNING: XLCellFormats constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLCellFormats::~XLCellFormats()
{
    m_cellFormats.clear();    // delete vector with all children
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

/**
 * @details Copy assignment operator
 */
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

/**
 * @details Returns the amount of cell format descriptions held by the class
 */
size_t XLCellFormats::count() const { return m_cellFormats.size(); }

/**
 * @details fetch XLCellFormat from m_cellFormats by index
 */
XLCellFormat XLCellFormats::cellFormatByIndex(XLStyleIndex index) const
{
    Expects(index < m_cellFormats.size());
    return m_cellFormats.at(index);
}

/**
 * @details append a new XLCellFormat to m_cellFormats and m_cellFormatsNode, based on copyFrom
 */
XLStyleIndex XLCellFormats::create(XLCellFormat copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the cell format to be created
    XMLNode      newNode{};          // scope declaration

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
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_cellFormatsNode->insert_child_before(pugi::node_pcdata, newNode)
            .set_value(std::string(styleEntriesPrefix).c_str());    // prefix the new node with styleEntriesPrefix

    XLCellFormat newCellFormat(newNode, m_permitXfId);
    if (copyFrom.m_cellFormatNode->empty()) {    // if no template is given
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
        copyXMLNode(newNode, *copyFrom.m_cellFormatNode);    // will use copyFrom as template, does nothing if copyFrom is empty

    m_cellFormats.push_back(newCellFormat);
    appendAndSetAttribute(*m_cellFormatsNode, "count", std::to_string(m_cellFormats.size()));    // update array count in XML
    return index;
}


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

/**
 * @details Constructor. Initializes an empty XLCellStyle object
 */
XLCellStyle::XLCellStyle() : m_cellStyleNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLCellStyle object.
 */
XLCellStyle::XLCellStyle(const XMLNode& node) : m_cellStyleNode(std::make_unique<XMLNode>(node)) {}

XLCellStyle::~XLCellStyle() = default;

XLCellStyle::XLCellStyle(const XLCellStyle& other) : m_cellStyleNode(std::make_unique<XMLNode>(*other.m_cellStyleNode)) {}

XLCellStyle& XLCellStyle::operator=(const XLCellStyle& other)
{
    if (&other != this) *m_cellStyleNode = *other.m_cellStyleNode;
    return *this;
}

/**
 * @details Returns the style empty status
 */
bool XLCellStyle::empty() const { return m_cellStyleNode->empty(); }

/**
 * @details Getter functions
 */
std::string  XLCellStyle::name() const { return m_cellStyleNode->attribute("name").value(); }
XLStyleIndex XLCellStyle::xfId() const { return m_cellStyleNode->attribute("xfId").as_uint(XLInvalidStyleIndex); }
uint32_t     XLCellStyle::builtinId() const { return m_cellStyleNode->attribute("builtinId").as_uint(XLInvalidUInt32); }
uint32_t     XLCellStyle::outlineStyle() const { return m_cellStyleNode->attribute("iLevel").as_uint(XLInvalidUInt32); }
bool         XLCellStyle::hidden() const { return m_cellStyleNode->attribute("hidden").as_bool(); }
bool         XLCellStyle::customBuiltin() const { return m_cellStyleNode->attribute("customBuiltin").as_bool(); }

/**
 * @details Setter functions
 */
bool XLCellStyle::setName(std::string_view newName) { return appendAndSetAttribute(*m_cellStyleNode, "name", std::string(newName)).empty() == false; }
bool XLCellStyle::setXfId(XLStyleIndex newXfId)
{ return appendAndSetAttribute(*m_cellStyleNode, "xfId", std::to_string(newXfId)).empty() == false; }
bool XLCellStyle::setBuiltinId(uint32_t newBuiltinId)
{ return appendAndSetAttribute(*m_cellStyleNode, "builtinId", std::to_string(newBuiltinId)).empty() == false; }
bool XLCellStyle::setOutlineStyle(uint32_t newOutlineStyle)
{ return appendAndSetAttribute(*m_cellStyleNode, "iLevel", std::to_string(newOutlineStyle)).empty() == false; }
bool XLCellStyle::setHidden(bool set)
{ return appendAndSetAttribute(*m_cellStyleNode, "hidden", (set ? "true" : "false")).empty() == false; }
bool XLCellStyle::setCustomBuiltin(bool set)
{ return appendAndSetAttribute(*m_cellStyleNode, "customBuiltin", (set ? "true" : "false")).empty() == false; }

/**
 * @brief Unsupported setter function
 */
bool XLCellStyle::setExtLst(XLUnsupportedElement const& newExtLst)
{
    OpenXLSX::ignore(newExtLst);
    return false;
}

/**
 * @details assemble a string summary about the cell style
 */
std::string XLCellStyle::summary() const
{
    uint32_t iLevel = outlineStyle();
    return fmt::format("name={}, xfId={}, builtinId={}{}{}{}",
                       name(),
                       xfId(),
                       builtinId(),
                       iLevel != XLInvalidUInt32 ? fmt::format(", iLevel={}", outlineStyle()) : "",
                       hidden() ? ", hidden=true" : "",
                       customBuiltin() ? ", customBuiltin=true" : "");
}

// ===== XLCellStyles, parent of XLCellStyle

/**
 * @details Constructor. Initializes an empty XLCellStyles object
 */
XLCellStyles::XLCellStyles() : m_cellStylesNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLCellStyles object.
 */
XLCellStyles::XLCellStyles(const XMLNode& cellStyles) : m_cellStylesNode(std::make_unique<XMLNode>(cellStyles))
{
    // initialize XLCellStyles entries and m_cellStyles here
    XMLNode node = cellStyles.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        if (nodeName == "cellStyle")
            m_cellStyles.push_back(XLCellStyle(node));
        else
            std::cerr << "WARNING: XLCellStyles constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLCellStyles::~XLCellStyles()
{
    m_cellStyles.clear();    // delete vector with all children
}

XLCellStyles::XLCellStyles(const XLCellStyles& other)
    : m_cellStylesNode(std::make_unique<XMLNode>(*other.m_cellStylesNode)),
      m_cellStyles(other.m_cellStyles)
{}

XLCellStyles::XLCellStyles(XLCellStyles&& other)
    : m_cellStylesNode(std::move(other.m_cellStylesNode)),
      m_cellStyles(std::move(other.m_cellStyles))
{}

/**
 * @details Copy assignment operator
 */
XLCellStyles& XLCellStyles::operator=(const XLCellStyles& other)
{
    if (&other != this) {
        *m_cellStylesNode = *other.m_cellStylesNode;
        m_cellStyles.clear();
        m_cellStyles = other.m_cellStyles;
    }
    return *this;
}

/**
 * @details Returns the amount of numberFormats held by the class
 */
size_t XLCellStyles::count() const { return m_cellStyles.size(); }

/**
 * @details fetch XLCellStyle from m_cellStyles by index
 */
XLCellStyle XLCellStyles::cellStyleByIndex(XLStyleIndex index) const
{
    Expects(index < m_cellStyles.size());
    return m_cellStyles.at(index);
}

/**
 * @details append a new XLCellStyle to m_cellStyles and m_cellStyleNode, based on copyFrom
 */
XLStyleIndex XLCellStyles::create(XLCellStyle copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the cell style to be created
    XMLNode      newNode{};          // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_cellStylesNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty())
        newNode = m_cellStylesNode->prepend_child("cellStyle");
    else
        newNode = m_cellStylesNode->insert_child_after("cellStyle", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCellStyles::"s + __func__ + ": failed to append a new cellStyle node"s);
    }
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_cellStylesNode->insert_child_before(pugi::node_pcdata, newNode)
            .set_value(std::string(styleEntriesPrefix).c_str());    // prefix the new node with styleEntriesPrefix

    XLCellStyle newCellStyle(newNode);
    if (copyFrom.m_cellStyleNode->empty()) {    // if no template is given
        // ===== Create a cell style with default values
        newCellStyle.setName("");
        newCellStyle.setXfId(0);
        newCellStyle.setHidden(false);
    }
    else
        copyXMLNode(newNode, *copyFrom.m_cellStyleNode);    // will use copyFrom as template, does nothing if copyFrom is empty

    m_cellStyles.push_back(newCellStyle);
    appendAndSetAttribute(*m_cellStylesNode, "count", std::to_string(m_cellStyles.size()));    // update array count in XML
    return index;
}


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
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLStyles.hpp"
#include "XLUtilities.hpp"
#include "XLStyles_Internal.hpp"

using namespace OpenXLSX;


/**
 * @details Constructor. Initializes an empty XLNumberFormat object
 */
XLNumberFormat::XLNumberFormat() : m_numberFormatNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLNumberFormat object.
 */
XLNumberFormat::XLNumberFormat(const XMLNode& node) : m_numberFormatNode(std::make_unique<XMLNode>(node)) {}

XLNumberFormat::~XLNumberFormat() = default;

XLNumberFormat::XLNumberFormat(const XLNumberFormat& other) : m_numberFormatNode(std::make_unique<XMLNode>(*other.m_numberFormatNode)) {}

XLNumberFormat& XLNumberFormat::operator=(const XLNumberFormat& other)
{
    if (&other != this) *m_numberFormatNode = *other.m_numberFormatNode;
    return *this;
}

/**
 * @details Returns the numFmtId value
 */
uint32_t XLNumberFormat::numberFormatId() const { return m_numberFormatNode->attribute("numFmtId").as_uint(XLInvalidUInt32); }

/**
 * @details Returns the formatCode value
 */
std::string XLNumberFormat::formatCode() const { return m_numberFormatNode->attribute("formatCode").value(); }

/**
 * @details Setter functions
 */
bool XLNumberFormat::setNumberFormatId(uint32_t newNumberFormatId)
{ return appendAndSetAttribute(*m_numberFormatNode, "numFmtId", std::to_string(newNumberFormatId)).empty() == false; }
bool XLNumberFormat::setFormatCode(std::string_view newFormatCode)
{ return appendAndSetAttribute(*m_numberFormatNode, "formatCode", std::string(newFormatCode).c_str()).empty() == false; }

/**
 * @details assemble a string summary about the number format
 */
std::string XLNumberFormat::summary() const
{
    return fmt::format("numFmtId={}, formatCode={}", numberFormatId(), formatCode());
}

// ===== XLNumberFormats, parent of XLNumberFormat

/**
 * @details Constructor. Initializes an empty XLNumberFormats object
 */
XLNumberFormats::XLNumberFormats() : m_numberFormatsNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLNumberFormats object.
 */
XLNumberFormats::XLNumberFormats(const XMLNode& numberFormats) : m_numberFormatsNode(std::make_unique<XMLNode>(numberFormats))
{
    // initialize XLNumberFormat entries and m_numberFormats here
    XMLNode node = numberFormats.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        if (nodeName == "numFmt")
            m_numberFormats.push_back(XLNumberFormat(node));
        else
            std::cerr << "WARNING: XLNumberFormats constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLNumberFormats::~XLNumberFormats()
{
    m_numberFormats.clear();    // delete vector with all children
}

XLNumberFormats::XLNumberFormats(const XLNumberFormats& other)
    : m_numberFormatsNode(std::make_unique<XMLNode>(*other.m_numberFormatsNode)),
      m_numberFormats(other.m_numberFormats)
{}

XLNumberFormats::XLNumberFormats(XLNumberFormats&& other)
    : m_numberFormatsNode(std::move(other.m_numberFormatsNode)),
      m_numberFormats(std::move(other.m_numberFormats))
{}

/**
 * @details Copy assignment operator
 */
XLNumberFormats& XLNumberFormats::operator=(const XLNumberFormats& other)
{
    if (&other != this) {
        *m_numberFormatsNode = *other.m_numberFormatsNode;
        m_numberFormats.clear();
        m_numberFormats = other.m_numberFormats;
    }
    return *this;
}

/**
 * @details Returns the amount of numberFormats held by the class
 */
size_t XLNumberFormats::count() const { return m_numberFormats.size(); }

/**
 * @details fetch XLNumberFormat from m_numberFormats by index
 */
XLNumberFormat XLNumberFormats::numberFormatByIndex(XLStyleIndex index) const
{
    Expects(index < m_numberFormats.size());
    return m_numberFormats.at(index);
}

/**
 * @details fetch XLNumberFormat from m_numberFormats by its numberFormatId
 */
XLNumberFormat XLNumberFormats::numberFormatById(uint32_t numberFormatId) const
{
    for (XLNumberFormat fmt : m_numberFormats)
        if (fmt.numberFormatId() == numberFormatId) return fmt;
    using namespace std::literals::string_literals;
    throw XLException("XLNumberFormats::"s + __func__ + ": numberFormatId "s + std::to_string(numberFormatId) + " not found"s);
}

/**
 * @details fetch a numFmtId from m_numberFormats by index
 */
uint32_t XLNumberFormats::numberFormatIdFromIndex(XLStyleIndex index) const
{
    Expects(index < m_numberFormats.size());
    return m_numberFormats[index].numberFormatId();
}

/**
 * @details Create a new custom number format with a unique ID and format code
 */
uint32_t XLNumberFormats::createNumberFormat(std::string_view formatCode)
{
    uint32_t maxId = 163;
    for (const auto& fmt : m_numberFormats) {
        if (fmt.formatCode() == formatCode) {
            return fmt.numberFormatId();
        }
        if (fmt.numberFormatId() > maxId) {
            maxId = fmt.numberFormatId();
        }
    }
    
    uint32_t newId = maxId + 1;
    
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
    auto countAttr = m_numberFormatsNode->attribute("count");
    if (countAttr.empty()) {
        m_numberFormatsNode->append_attribute("count").set_value(m_numberFormats.size());
    } else {
        countAttr.set_value(m_numberFormats.size());
    }
    
    return newId;
}

/**
 * @details append a new XLNumberFormat to m_numberFormats and m_numberFormatsNode, based on copyFrom
 */
XLStyleIndex XLNumberFormats::create(XLNumberFormat copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the number format to be created
    XMLNode      newNode{};          // scope declaration

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
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_numberFormatsNode->insert_child_before(pugi::node_pcdata, newNode)
            .set_value(std::string(styleEntriesPrefix).c_str());    // prefix the new node with styleEntriesPrefix

    XLNumberFormat newNumberFormat(newNode);
    if (copyFrom.m_numberFormatNode->empty()) {    // if no template is given
        // ===== Create a number format with default values
        newNumberFormat.setNumberFormatId(0);
        newNumberFormat.setFormatCode("General");
    }
    else
        copyXMLNode(newNode, *copyFrom.m_numberFormatNode);    // will use copyFrom as template, does nothing if copyFrom is empty

    m_numberFormats.push_back(newNumberFormat);
    appendAndSetAttribute(*m_numberFormatsNode, "count", std::to_string(m_numberFormats.size()));    // update array count in XML
    return index;
}


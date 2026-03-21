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

// ===== XLDxf, formerly XLDiffCellFormat

/**
 * @details Default constructor. Initializes an empty XLDxf object with a temporary XML document.
 */
XLDxf::XLDxf() : m_xmlDocument(std::make_unique<XMLDocument>())
{
    m_dxfNode = m_xmlDocument->append_child("dxf");
}

/**
 * @details Constructor. Initializes the member variables for the new XLDxf object.
 */
XLDxf::XLDxf(const XMLNode& node) : m_dxfNode(node) {}

XLDxf::~XLDxf() = default;

XLDxf::XLDxf(const XLDxf& other) : m_dxfNode(other.m_dxfNode)
{
    if (other.m_xmlDocument) {
        m_xmlDocument = std::make_unique<XMLDocument>();
        m_xmlDocument->reset(*other.m_xmlDocument);
        m_dxfNode = m_xmlDocument->document_element();
    }
}

XLDxf::XLDxf(XLDxf&& other) noexcept : m_xmlDocument(std::move(other.m_xmlDocument)), m_dxfNode(other.m_dxfNode) {}

XLDxf& XLDxf::operator=(const XLDxf& other)
{
    if (&other != this) {
        m_dxfNode = other.m_dxfNode;
        if (other.m_xmlDocument) {
            m_xmlDocument = std::make_unique<XMLDocument>();
            m_xmlDocument->reset(*other.m_xmlDocument);
            m_dxfNode = m_xmlDocument->document_element();
        }
        else {
            m_xmlDocument.reset();
        }
    }
    return *this;
}

XLDxf& XLDxf::operator=(XLDxf&& other) noexcept
{
    if (&other != this) {
        m_xmlDocument = std::move(other.m_xmlDocument);
        m_dxfNode     = other.m_dxfNode;
    }
    return *this;
}

/**
 * @details Returns the differential cell format empty status
 */
bool XLDxf::empty() const { return m_dxfNode.empty(); }

/**
 * @details Getter functions
 */
XLFont XLDxf::font() const
{
    XMLNode fontNode = appendAndGetNode(m_dxfNode, "font");
    if (fontNode.empty()) return XLFont{};
    return XLFont(fontNode);
}
XLNumberFormat XLDxf::numFmt() const
{
    XMLNode numFmtNode = appendAndGetNode(m_dxfNode, "numFmt");
    if (numFmtNode.empty()) return XLNumberFormat{};
    return XLNumberFormat(numFmtNode);
}
XLFill XLDxf::fill() const
{
    XMLNode fillNode = appendAndGetNode(m_dxfNode, "fill");
    if (fillNode.empty()) return XLFill{};
    return XLFill(fillNode);
}
XLAlignment XLDxf::alignment() const
{
    XMLNode alignmentNode = appendAndGetNode(m_dxfNode, "alignment");
    if (alignmentNode.empty()) return XLAlignment{};
    return XLAlignment(alignmentNode);
}
XLBorder XLDxf::border() const
{
    XMLNode borderNode = appendAndGetNode(m_dxfNode, "border");
    if (borderNode.empty()) return XLBorder{};
    return XLBorder(borderNode);
}

bool XLDxf::hasFont() const { return !m_dxfNode.child("font").empty(); }
bool XLDxf::hasNumFmt() const { return !m_dxfNode.child("numFmt").empty(); }
bool XLDxf::hasFill() const { return !m_dxfNode.child("fill").empty(); }
bool XLDxf::hasAlignment() const { return !m_dxfNode.child("alignment").empty(); }
bool XLDxf::hasBorder() const { return !m_dxfNode.child("border").empty(); }

void XLDxf::clearFont() { m_dxfNode.remove_child("font"); }
void XLDxf::clearNumFmt() { m_dxfNode.remove_child("numFmt"); }
void XLDxf::clearFill() { m_dxfNode.remove_child("fill"); }
void XLDxf::clearAlignment() { m_dxfNode.remove_child("alignment"); }
void XLDxf::clearBorder() { m_dxfNode.remove_child("border"); }

/**
 * @brief Unsupported setter function
 */
bool XLDxf::setExtLst(XLUnsupportedElement const& newExtLst)
{
    OpenXLSX::ignore(newExtLst);
    return false;
}

/**
 * @details assemble a string summary about the differential cell format
 */
std::string XLDxf::summary() const
{
    return fmt::format("DXF containing: font={}, fill={}, border={}, alignment={}, numFmt={}",
                       hasFont() ? "yes" : "no",
                       hasFill() ? "yes" : "no",
                       hasBorder() ? "yes" : "no",
                       hasAlignment() ? "yes" : "no",
                       hasNumFmt() ? "yes" : "no");
}

// ===== XLDxfs, parent of XLDxf

/**
 * @details Constructor. Initializes an empty XLDxfs object
 */
XLDxfs::XLDxfs() : m_dxfsNode(XMLNode()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLDxfs object.
 */
XLDxfs::XLDxfs(const XMLNode& dxfs) : m_dxfsNode(dxfs)
{
    XMLNode node = m_dxfsNode.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        if (nodeName == "dxf")
            m_dxfs.push_back(XLDxf(node));
        else
            std::cerr << "WARNING: XLDxfs constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLDxfs::~XLDxfs()
{
    m_dxfs.clear();    // delete vector with all children
}

XLDxfs::XLDxfs(const XLDxfs& other)
    : m_dxfsNode(other.m_dxfsNode),
      m_dxfs(other.m_dxfs)
{}

XLDxfs::XLDxfs(XLDxfs&& other)
    : m_dxfsNode(std::move(other.m_dxfsNode)),
      m_dxfs(std::move(other.m_dxfs))
{}

/**
 * @details Copy assignment operator
 */
XLDxfs& XLDxfs::operator=(const XLDxfs& other)
{
    if (&other != this) {
        m_dxfsNode = other.m_dxfsNode;
        m_dxfs.clear();
        m_dxfs = other.m_dxfs;
    }
    return *this;
}

XLDxfs& XLDxfs::operator=(XLDxfs&& other) noexcept
{
    if (&other != this) {
        m_dxfsNode = std::move(other.m_dxfsNode);
        m_dxfs     = std::move(other.m_dxfs);
    }
    return *this;
}

/**
 * @details Returns the amount of differential cell formats held by the class
 */
size_t XLDxfs::count() const { return m_dxfs.size(); }

/**
 * @details fetch XLDxf from m_dxfs by index
 */
XLDxf XLDxfs::dxfByIndex(XLStyleIndex index) const
{
    Expects(index < m_dxfs.size());
    return m_dxfs.at(index);
}

/**
 * @details append a new XLDxf to m_dxfs and m_dxfsNode, based on copyFrom
 */
XLStyleIndex XLDxfs::create(XLDxf copyFrom, std::string styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the cell style to be created
    XMLNode      newNode{};          // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastDxf = m_dxfsNode.last_child_of_type(pugi::node_element);
    if (lastDxf.empty())
        newNode = m_dxfsNode.prepend_child("dxf");
    else
        newNode = m_dxfsNode.insert_child_after("dxf", lastDxf);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLDxfs::"s + __func__ + ": failed to append a new dxf node"s);
    }
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_dxfsNode.insert_child_before(pugi::node_pcdata, newNode)
            .set_value(styleEntriesPrefix.c_str());    // prefix the new node with styleEntriesPrefix

    XLDxf newDxf(newNode);
    if (copyFrom.node().empty()) {    // if no template is given
        // no-op: differential cell format entries have NO default values
    }
    else
        copyXMLNode(newNode, copyFrom.node());    // will use copyFrom as template, does nothing if copyFrom is empty

    m_dxfs.push_back(newDxf);
    appendAndSetAttribute(m_dxfsNode, "count", std::to_string(m_dxfs.size()));    // update array count in XML
    return index;
}


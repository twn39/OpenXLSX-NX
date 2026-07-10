// ===== External Includes ===== //
#include <cstdint>
#include <fmt/format.h>
#include <gsl/gsl>
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
#include "XLStyles_Internal.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

// ===== XLDxf, formerly XLDiffCellFormat

XLDxf::XLDxf() : m_xmlDocument(std::make_unique<XMLDocument>()) { m_dxfNode = m_xmlDocument->append_child("dxf"); }

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

bool XLDxf::empty() const { return m_dxfNode.empty(); }

XLFont XLDxf::font() const
{
    XMLNode fontNode = ensureChild(m_dxfNode, "font");
    if (fontNode.empty()) return XLFont{};
    return XLFont(fontNode);
}
XLNumberFormat XLDxf::numFmt() const
{
    XMLNode numFmtNode = ensureChild(m_dxfNode, "numFmt");
    if (numFmtNode.empty()) return XLNumberFormat{};
    return XLNumberFormat(numFmtNode);
}
XLFill XLDxf::fill() const
{
    XMLNode fillNode = ensureChild(m_dxfNode, "fill");
    if (fillNode.empty()) return XLFill{};
    return XLFill(fillNode);
}
XLAlignment XLDxf::alignment() const
{
    XMLNode alignmentNode = ensureChild(m_dxfNode, "alignment");
    if (alignmentNode.empty()) return XLAlignment{};
    return XLAlignment(alignmentNode);
}
XLBorder XLDxf::border() const
{
    XMLNode borderNode = ensureChild(m_dxfNode, "border");
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

XLDxfs::XLDxfs() : m_dxfsNode(XMLNode()) {}

XLDxfs::XLDxfs(const XMLNode& dxfs) : m_dxfsNode(dxfs)
{
    XMLNode node = m_dxfsNode.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string_view nodeName(node.name());
        if (nodeName == "dxf")
            m_dxfs.push_back(XLDxf(node));
        else
            std::cerr << "WARNING: XLDxfs constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLDxfs::~XLDxfs() { m_dxfs.clear(); }

XLDxfs::XLDxfs(const XLDxfs& other) : m_dxfsNode(other.m_dxfsNode), m_dxfs(other.m_dxfs) {}

XLDxfs::XLDxfs(XLDxfs&& other) : m_dxfsNode(std::move(other.m_dxfsNode)), m_dxfs(std::move(other.m_dxfs)) {}

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

size_t XLDxfs::count() const { return m_dxfs.size(); }

XLDxf XLDxfs::dxfByIndex(XLStyleIndex index) const
{
    Expects(index < m_dxfs.size());
    return m_dxfs.at(index);
}

XLStyleIndex XLDxfs::create(XLDxf copyFrom, std::string_view styleEntriesPrefix)
{
    XLStyleIndex index = count();
    XMLNode      newNode{};

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
    if (styleEntriesPrefix.length() > 0)
        m_dxfsNode.insert_child_before(pugi::node_pcdata, newNode).set_value(std::string(styleEntriesPrefix).c_str());

    XLDxf newDxf(newNode);
    if (copyFrom.node().empty()) {
        // no-op: differential cell format entries have NO default values
    }
    else
        copyXMLNode(newNode, copyFrom.node());

    m_dxfs.push_back(newDxf);
    setAttr(m_dxfsNode, "count", std::to_string(m_dxfs.size()));
    return index;
}

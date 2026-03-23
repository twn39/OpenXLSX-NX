#include "XLSparkline.hpp"
#include "XLUtilities.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

XLSparkline::XLSparkline() : m_node(nullptr) {}

XLSparkline::XLSparkline(XMLNode node) : m_node(node) {}

XLSparkline::~XLSparkline() = default;

XLSparklineType XLSparkline::type() const
{
    std::string typeStr = m_node.attribute("type").value();
    if (typeStr == "column") return XLSparklineType::Column;
    if (typeStr == "stacked") return XLSparklineType::Stacked;
    return XLSparklineType::Line;
}

void XLSparkline::setType(XLSparklineType type)
{
    std::string typeStr;
    switch (type) {
        case XLSparklineType::Column: typeStr = "column"; break;
        case XLSparklineType::Stacked: typeStr = "stacked"; break;
        case XLSparklineType::Line: typeStr = "line"; break;
    }
    
    // Setting attribute on the sparklineGroup node
    if (m_node.attribute("type").empty())
        m_node.append_attribute("type") = typeStr.c_str();
    else
        m_node.attribute("type") = typeStr.c_str();
}

std::string XLSparkline::location() const
{
    XMLNode slNode = m_node.child("x14:sparklines").child("x14:sparkline");
    if (slNode.empty()) return "";
    return slNode.child("xm:sqref").text().get();
}

void XLSparkline::setLocation(const std::string& sqref)
{
    XMLNode slsNode = m_node.child("x14:sparklines");
    if (slsNode.empty()) {
        slsNode = m_node.append_child("x14:sparklines");
    }
    
    XMLNode slNode = slsNode.child("x14:sparkline");
    if (slNode.empty()) {
        slNode = slsNode.append_child("x14:sparkline");
    }
    
    XMLNode sqrefNode = slNode.child("xm:sqref");
    if (sqrefNode.empty()) {
        sqrefNode = slNode.append_child("xm:sqref");
    }
    
    sqrefNode.text().set(sqref.c_str());
}

std::string XLSparkline::dataRange() const
{
    XMLNode slNode = m_node.child("x14:sparklines").child("x14:sparkline");
    if (slNode.empty()) return "";
    return slNode.child("xm:f").text().get();
}

void XLSparkline::setDataRange(const std::string& formula)
{
    XMLNode slsNode = m_node.child("x14:sparklines");
    if (slsNode.empty()) {
        slsNode = m_node.append_child("x14:sparklines");
    }
    
    XMLNode slNode = slsNode.child("x14:sparkline");
    if (slNode.empty()) {
        slNode = slsNode.append_child("x14:sparkline");
    }
    
    XMLNode fNode = slNode.child("xm:f");
    if (fNode.empty()) {
        fNode = slNode.prepend_child("xm:f"); // Must be before sqref
    }
    
    fNode.text().set(formula.c_str());
}
/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#include "XLSparkline.hpp"
#include "XLException.hpp"
#include "XLUtilities.hpp"

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
        case XLSparklineType::Column:
            typeStr = "column";
            break;
        case XLSparklineType::Stacked:
            typeStr = "stacked";
            break;
        case XLSparklineType::Line:
            typeStr = "line";
            break;
    }

    setAttr(m_node, "type", typeStr.c_str());
}

std::string XLSparkline::location() const
{
    XMLNode slNode = m_node.child("x14:sparklines").child("x14:sparkline");
    if (slNode.empty()) return "";
    return slNode.child("xm:sqref").text().get();
}

void XLSparkline::setLocation(const std::string& sqref)
{
    XMLNode slsNode   = ensureChild(m_node, "x14:sparklines");
    XMLNode slNode    = ensureChild(slsNode, "x14:sparkline");
    setChildText(slNode, "xm:sqref", sqref);
}

std::string XLSparkline::dataRange() const
{
    XMLNode slNode = m_node.child("x14:sparklines").child("x14:sparkline");
    if (slNode.empty()) return "";
    return slNode.child("xm:f").text().get();
}

void XLSparkline::setDataRange(const std::string& formula)
{
    XMLNode slsNode = ensureChild(m_node, "x14:sparklines");
    XMLNode slNode  = ensureChild(slsNode, "x14:sparkline");

    // xm:f must appear before xm:sqref per OOXML sparkline schema
    XMLNode fNode = slNode.child("xm:f");
    if (fNode.empty()) fNode = slNode.prepend_child("xm:f");
    fNode.text().set(formula.c_str());
}
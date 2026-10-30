/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

// ===== External Includes ===== //
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLColumn.hpp"
#include "XLStyles_Common.hpp"    // XLDefaultCellFormat
#include "XLXmlHelpers.hpp"
#include "XLXmlSchema.hpp"

using namespace OpenXLSX;

/**
 * @details Assumes each node only has data for one column.
 */
XLColumn::XLColumn(XMLNode columnNode) : m_columnNode(columnNode) {}

XMLAttribute XLColumn::getOrCreateAttribute(const char* attrName)
{
    return ensureAttr(m_columnNode, attrName, "");
}

/**
 * @details
 */
float XLColumn::width() const { return m_columnNode.attribute(XmlAttr::Width).as_float(0.0f); }

/**
 * @details
 */
void XLColumn::setWidth(float width)    // NOLINT
{
    setAttr(m_columnNode, XmlAttr::Width, width);
    setAttr(m_columnNode, XmlAttr::CustomWidth, 1);
}

/**
 * @details
 */
bool XLColumn::isHidden() const { return m_columnNode.attribute(XmlAttr::Hidden).as_bool(false); }

/**
 * @details
 */
void XLColumn::setHidden(bool state)    // NOLINT
{
    if (state) {
        setAttr(m_columnNode, XmlAttr::Hidden, 1);
    } else {
        m_columnNode.remove_attribute(XmlAttr::Hidden);
    }
}

/**
 * @details Get the outline level of the column.
 */
uint8_t XLColumn::outlineLevel() const
{
    return static_cast<uint8_t>(std::clamp(m_columnNode.attribute(XmlAttr::OutlineLevel).as_uint(0), 0u, 7u));
}

/**
 * @details Set the outline level of the column.
 */
void XLColumn::setOutlineLevel(uint8_t level)
{
    level = static_cast<uint8_t>(std::clamp(static_cast<unsigned>(level), 0u, 7u));
    if (level == 0) {
        m_columnNode.remove_attribute(XmlAttr::OutlineLevel);
    } else {
        setAttr(m_columnNode, XmlAttr::OutlineLevel, level);
    }
}

/**
 * @details Is the column collapsed?
 */
bool XLColumn::isCollapsed() const { return m_columnNode.attribute(XmlAttr::Collapsed).as_bool(false); }

/**
 * @details Set the column to be collapsed or expanded.
 */
void XLColumn::setCollapsed(bool state)
{
    if (state) {
        setAttr(m_columnNode, XmlAttr::Collapsed, 1);
    } else {
        m_columnNode.remove_attribute(XmlAttr::Collapsed);
    }
}

/**
 * @details
 */
XMLNode XLColumn::columnNode() const { return m_columnNode; }

/**
 * @details Determine the value of the style attribute - if attribute does not exist, return default value
 */
XLStyleIndex XLColumn::format() const { return m_columnNode.attribute(XmlAttr::ColStyle).as_uint(XLDefaultCellFormat); }

/**
 * @brief Set the column style as a reference to the array index of xl/styles.xml:<styleSheet>:<cellXfs>
 *        If the style attribute does not exist, create it
 */
bool XLColumn::setFormat(XLStyleIndex cellFormatIndex)
{
    setAttr(m_columnNode, XmlAttr::ColStyle, cellFormatIndex);
    return true;
}

XLColumnProperties XLColumn::properties() const
{
    return XmlOps::readColProps(m_columnNode);
}

void XLColumn::setProperties(const XLColumnProperties& props)
{
    XmlOps::writeColProps(m_columnNode, props);
}

void XLColumn::autoFit()
{ throw XLInternalError("autoFit() requires XLWorksheet context. Please use XLWorksheet::autoFitColumn(colIndex)."); }

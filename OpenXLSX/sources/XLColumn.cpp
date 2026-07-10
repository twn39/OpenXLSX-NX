// ===== External Includes ===== //
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLColumn.hpp"
#include "XLStyles.hpp"    // XLDefaultCellFormat
#include "XLXmlHelpers.hpp"

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
float XLColumn::width() const { return m_columnNode.attribute("width").as_float(); }

/**
 * @details
 */
void XLColumn::setWidth(float width)    // NOLINT
{
    setAttr(m_columnNode, "width", width);
    setAttr(m_columnNode, "customWidth", 1);
}

/**
 * @details
 */
bool XLColumn::isHidden() const { return m_columnNode.attribute("hidden").as_bool(); }

/**
 * @details
 */
void XLColumn::setHidden(bool state)    // NOLINT
{ getOrCreateAttribute("hidden").set_value(state ? 1 : 0); }

/**
 * @details Get the outline level of the column.
 */
uint8_t XLColumn::outlineLevel() const { return static_cast<uint8_t>(m_columnNode.attribute("outlineLevel").as_uint(0)); }

/**
 * @details Set the outline level of the column.
 */
void XLColumn::setOutlineLevel(uint8_t level)
{
    constexpr uint8_t maxOutlineLevel = 7;
    if (level > maxOutlineLevel) level = maxOutlineLevel;

    if (level == 0) { m_columnNode.remove_attribute("outlineLevel"); }
    else {
        getOrCreateAttribute("outlineLevel").set_value(level);
    }
}

/**
 * @details Is the column collapsed?
 */
bool XLColumn::isCollapsed() const { return m_columnNode.attribute("collapsed").as_bool(false); }

/**
 * @details Set the column to be collapsed or expanded.
 */
void XLColumn::setCollapsed(bool state)
{
    if (!state) { m_columnNode.remove_attribute("collapsed"); }
    else {
        getOrCreateAttribute("collapsed").set_value(1);
    }
}

/**
 * @details
 */
XMLNode XLColumn::columnNode() const { return m_columnNode; }

/**
 * @details Determine the value of the style attribute - if attribute does not exist, return default value
 */
XLStyleIndex XLColumn::format() const { return m_columnNode.attribute("style").as_uint(XLDefaultCellFormat); }

/**
 * @brief Set the column style as a reference to the array index of xl/styles.xml:<styleSheet>:<cellXfs>
 *        If the style attribute does not exist, create it
 */
bool XLColumn::setFormat(XLStyleIndex cellFormatIndex)
{
    XMLAttribute styleAtt = getOrCreateAttribute("style");
    if (styleAtt.empty()) return false;
    styleAtt.set_value(cellFormatIndex);
    return true;
}

void XLColumn::autoFit()
{ throw XLInternalError("autoFit() requires XLWorksheet context. Please use XLWorksheet::autoFitColumn(colIndex)."); }

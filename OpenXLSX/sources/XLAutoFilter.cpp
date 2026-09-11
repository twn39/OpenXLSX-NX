/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#include "XLAutoFilter.hpp"
#include "XLCellRange.hpp"
#include <string>

namespace OpenXLSX
{

    // ========== XLFilterColumn Implementation ========== //

    XLFilterColumn::XLFilterColumn(XMLNode node) : m_node(node) {}

    void XLFilterColumn::addFilter(const std::string& value)
    {
        if (!m_node) return;

        // Ensure <filters> node exists
        XMLNode filtersNode = m_node.child("filters");
        if (!filtersNode) { filtersNode = m_node.append_child("filters"); }

        // Clear other filter types if they exist
        m_node.remove_child("customFilters");
        m_node.remove_child("top10");
        m_node.remove_child("dynamicFilter");
        m_node.remove_child("colorFilter");
        m_node.remove_child("iconFilter");

        XMLNode filterNode                 = filtersNode.append_child("filter");
        filterNode.append_attribute("val") = value.c_str();
    }

    void XLFilterColumn::clearFilters()
    {
        if (!m_node) return;
        m_node.remove_child("filters");
        m_node.remove_child("customFilters");
        m_node.remove_child("top10");
        m_node.remove_child("dynamicFilter");
        m_node.remove_child("colorFilter");
        m_node.remove_child("iconFilter");
    }

    void XLFilterColumn::setCustomFilter(const std::string& op, const std::string& val)
    {
        if (!m_node) return;

        clearFilters();

        XMLNode customFiltersNode                     = m_node.append_child("customFilters");
        XMLNode customFilterNode                      = customFiltersNode.append_child("customFilter");
        customFilterNode.append_attribute("operator") = op.c_str();
        customFilterNode.append_attribute("val")      = val.c_str();
    }

    void XLFilterColumn::setCustomFilter(const std::string& op1,
                                         const std::string& val1,
                                         XLFilterLogic      logic,
                                         const std::string& op2,
                                         const std::string& val2)
    {
        if (!m_node) return;

        clearFilters();

        XMLNode customFiltersNode = m_node.append_child("customFilters");
        if (logic == XLFilterLogic::And) { customFiltersNode.append_attribute("and") = 1; }    // "or" is default, no attribute needed

        XMLNode customFilterNode1                      = customFiltersNode.append_child("customFilter");
        customFilterNode1.append_attribute("operator") = op1.c_str();
        customFilterNode1.append_attribute("val")      = val1.c_str();

        XMLNode customFilterNode2                      = customFiltersNode.append_child("customFilter");
        customFilterNode2.append_attribute("operator") = op2.c_str();
        customFilterNode2.append_attribute("val")      = val2.c_str();
    }

    void XLFilterColumn::setTop10(double value, bool percent, bool top)
    {
        if (!m_node) return;

        clearFilters();

        XMLNode top10Node                 = m_node.append_child("top10");
        top10Node.append_attribute("val") = value;
        if (percent) { top10Node.append_attribute("percent") = 1; }
        if (!top) {
            top10Node.append_attribute("top") = 0;    // Default is true (1), so only write if false
        }
    }

    void XLFilterColumn::setDynamicFilter(const std::string& type)
    {
        if (!m_node) return;

        clearFilters();

        XMLNode dynamicFilterNode                  = m_node.append_child("dynamicFilter");
        dynamicFilterNode.append_attribute("type") = type.c_str();
    }

    uint16_t XLFilterColumn::colId() const
    {
        if (!m_node) return 0;
        return static_cast<uint16_t>(m_node.attribute("colId").as_uint());
    }

    // ========== XLAutoFilter Implementation ========== //

    XLAutoFilter::XLAutoFilter(XMLNode node) : m_node(node) {}

    XLAutoFilter::operator bool() const { return m_node != nullptr; }

    std::string XLAutoFilter::ref() const
    {
        if (!m_node) return "";
        return m_node.attribute("ref").value();
    }

    XLAutoFilter& XLAutoFilter::setRef(const std::string& ref)
    {
        if (!m_node) return *this;
        auto attr = m_node.attribute("ref");
        if (!attr) { m_node.append_attribute("ref") = ref.c_str(); }
        else {
            attr.set_value(ref.c_str());
        }
        return *this;
    }

    XLAutoFilter& XLAutoFilter::setRef(const XLCellRange& range) { return setRef(range.address()); }

    XLFilterColumn XLAutoFilter::filterColumn(uint16_t colId)
    {
        if (!m_node) return XLFilterColumn(XMLNode());

        for (auto child : m_node.children("filterColumn")) {
            if (child.attribute("colId").as_uint() == colId) { return XLFilterColumn(child); }
        }

        // If it doesn't exist, create it.
        XMLNode filterColumnNode                   = m_node.append_child("filterColumn");
        filterColumnNode.append_attribute("colId") = colId;
        return XLFilterColumn(filterColumnNode);
    }

}    // namespace OpenXLSX
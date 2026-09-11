/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#pragma once

#include "XLXmlParser.hpp"
#include <string>

namespace OpenXLSX
{

    /**
     * @brief Enum class defining the logical operator for custom filters.
     */
    enum class XLFilterLogic { And, Or };

    /**
     * @brief The XLFilterColumn class encapsulates the <filterColumn> XML node.
     * It allows setting different kinds of filter criteria for a specific column in an AutoFilter range.
     */
    class XLFilterColumn
    {
    public:
        /**
         * @brief Constructor
         * @param node The <filterColumn> XMLNode
         */
        explicit XLFilterColumn(XMLNode node);

        /**
         * @brief Add a specific string value to filter by.
         * @param value The value to add to the filter list.
         */
        void addFilter(const std::string& value);

        /**
         * @brief Clear all filters for this column.
         */
        void clearFilters();

        /**
         * @brief Set a single custom filter criteria.
         * @param op The operator (e.g., "greaterThan", "lessThan", "equal", "notEqual", "greaterThanOrEqual", "lessThanOrEqual")
         * @param val The value to compare against.
         */
        void setCustomFilter(const std::string& op, const std::string& val);

        /**
         * @brief Set a compound custom filter criteria.
         * @param op1 The first operator.
         * @param val1 The first value.
         * @param logic The logical operator joining the two conditions (And/Or).
         * @param op2 The second operator.
         * @param val2 The second value.
         */
        void setCustomFilter(const std::string& op1,
                             const std::string& val1,
                             XLFilterLogic      logic,
                             const std::string& op2,
                             const std::string& val2);

        /**
         * @brief Set a top-10 filter.
         * @param value The threshold value.
         * @param percent If true, filters by top percentage rather than count.
         * @param top If true, filters top values; if false, filters bottom values.
         */
        void setTop10(double value, bool percent = false, bool top = true);

        /**
         * @brief Set a dynamic filter (e.g. "aboveAverage", "today", "Q1").
         * @param type The type of dynamic filter.
         */
        void setDynamicFilter(const std::string& type);

        /**
         * @brief Get the column ID (0-based relative to the AutoFilter range).
         * @return The column ID.
         */
        uint16_t colId() const;

    private:
        XMLNode m_node;
    };

    /**
     * @brief The XLAutoFilter class encapsulates the <autoFilter> XML node.
     * It manages the target range and the individual column filters.
     */
    class XLCellRange;

    class XLAutoFilter
    {
    public:
        /**
         * @brief Constructor
         * @param node The <autoFilter> XMLNode
         */
        explicit XLAutoFilter(XMLNode node);

        /**
         * @brief Check if the AutoFilter object is valid (has a corresponding XML node).
         * @return true if valid, false otherwise.
         */
        operator bool() const;

        /**
         * @brief Get the reference range of the AutoFilter.
         * @return A string containing the reference (e.g., "A1:C10").
         */
        std::string ref() const;

        /**
         * @brief Set the reference range of the AutoFilter.
         * @param ref The reference range to set.
         */
        XLAutoFilter& setRef(const std::string& ref);

        /**
         * @brief Set the reference range of the AutoFilter using a strongly-typed XLCellRange.
         * @param range The XLCellRange object.
         */
        XLAutoFilter& setRef(const XLCellRange& range);

        /**
         * @brief Get or create a filter column by its ID.
         * @param colId The 0-based column ID relative to the AutoFilter range.
         * @return The XLFilterColumn object.
         */
        XLFilterColumn filterColumn(uint16_t colId);

    private:
        XMLNode m_node;
    };

}    // namespace OpenXLSX

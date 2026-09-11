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
#include <optional>
#include <string>
#include <string_view>

namespace OpenXLSX
{

    /**
     * @brief Enum class defining the possible functions for a table's totals row.
     */
    enum class XLTotalsRowFunction { None, Sum, Min, Max, Average, Count, CountNums, StdDev, Var, Custom };

    /**
     * @brief The XLTableColumn class encapsulates the <tableColumn> XML node.
     * It allows setting different kinds of column properties in a table.
     */
    class XLTableColumn
    {
    public:
        /**
         * @brief Constructor
         * @param node The <tableColumn> XMLNode
         */
        explicit XLTableColumn(XMLNode node);

        /**
         * @brief Check if the TableColumn object is valid (has a corresponding XML node).
         * @return true if valid, false otherwise.
         */
        operator bool() const;

        /**
         * @brief Get the ID of the table column.
         * @return The column ID.
         */
        uint32_t id() const;

        /**
         * @brief Get the name of the table column.
         * @return The column name.
         */
        std::string name() const;

        /**
         * @brief Set the name of the table column.
         * @param name The new name of the column.
         */
        void setName(std::string_view name);

        /**
         * @brief Get the totals row function for this column.
         * @return The XLTotalsRowFunction.
         */
        XLTotalsRowFunction totalsRowFunction() const;

        /**
         * @brief Set the totals row function for this column.
         * @param func The XLTotalsRowFunction to set.
         */
        void setTotalsRowFunction(XLTotalsRowFunction func);

        /**
         * @brief Get the totals row label for this column.
         * @return The custom string label.
         */
        std::string totalsRowLabel() const;

        /**
         * @brief Set the totals row label for this column.
         * @param label The string to set as label.
         */
        void setTotalsRowLabel(std::string_view label);

        /**
         * @brief Get the calculated column formula.
         * @return The formula string, or an empty string if not set.
         */
        std::string calculatedColumnFormula() const;

        /**
         * @brief Set the calculated column formula.
         * @param formula The formula string (e.g., "[@Price]*[@Quantity]").
         */
        void setCalculatedColumnFormula(std::string_view formula);

        /**
         * @brief Get the custom totals row formula.
         * @return The formula string, or an empty string if not set.
         */
        std::string totalsRowFormula() const;

        /**
         * @brief Set a custom totals row formula.
         * @param formula The formula string.
         */
        void setTotalsRowFormula(std::string_view formula);

    private:
        XMLNode m_node;
    };

}    // namespace OpenXLSX

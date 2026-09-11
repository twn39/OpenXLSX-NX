/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_IXLCELLPROVIDER_HPP
#define OPENXLSX_IXLCELLPROVIDER_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"
#include "XLCellRange.hpp"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

namespace OpenXLSX
{
    /**
     * @brief Abstract immutable interface for accessing cell values and formulas across spreadsheet grids.
     * @details Decouples consumers (such as XLFormulaEngine, XLCalculationEngine, XLQuery,
     *          and data exporters) from the concrete, mutable XLWorksheet implementation.
     *          Enables lightweight mock data providers for unit testing and virtual tabular grids.
     */
    class OPENXLSX_EXPORT IXLCellProvider
    {
    public:
        virtual ~IXLCellProvider() = default;

        /**
         * @brief Get cell value at 1-based (row, column).
         */
        [[nodiscard]] virtual XLCellValue getCellValue(uint32_t row, uint16_t col) const = 0;

        /**
         * @brief Get formula string at 1-based (row, column), or empty string if no formula.
         */
        [[nodiscard]] virtual std::string getCellFormula(uint32_t row, uint16_t col) const = 0;

        /**
         * @brief Check if cell exists and is populated.
         */
        [[nodiscard]] virtual bool hasCell(uint32_t row, uint16_t col) const = 0;

        /**
         * @brief Get sheet name.
         */
        [[nodiscard]] virtual std::string sheetName() const = 0;

        /**
         * @brief Batch fetch values for a rectangular range into an output vector.
         * @details Avoids per-cell virtual dispatch overhead in tight calculation/export loops.
         *          Default implementation falls back to row/column iteration.
         */
        virtual void fetchRangeValues(const XLCellRange& range, std::vector<XLCellValue>& out) const
        {
            uint32_t r1 = range.topLeft().row();
            uint32_t r2 = range.bottomRight().row();
            uint16_t c1 = range.topLeft().column();
            uint16_t c2 = range.bottomRight().column();

            size_t count = static_cast<size_t>(r2 - r1 + 1) * static_cast<size_t>(c2 - c1 + 1);
            out.reserve(out.size() + count);

            for (uint32_t r = r1; r <= r2; ++r) {
                for (uint16_t c = c1; c <= c2; ++c) {
                    out.emplace_back(getCellValue(r, c));
                }
            }
        }
    };
} // namespace OpenXLSX

#endif // OPENXLSX_IXLCELLPROVIDER_HPP

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

#include "OpenXLSX-Exports.hpp"
#include <string>

namespace OpenXLSX
{
    /**
     * @brief Enum defining all valid dynamic filter types according to ECMA-376 18.18.25.
     */
    enum class XLDynamicFilterType {
        Null,
        AboveAverage,
        BelowAverage,
        Tomorrow,
        Today,
        Yesterday,
        NextWeek,
        ThisWeek,
        LastWeek,
        NextMonth,
        ThisMonth,
        LastMonth,
        NextQuarter,
        ThisQuarter,
        LastQuarter,
        NextYear,
        ThisYear,
        LastYear,
        YearToDate,
        Q1,
        Q2,
        Q3,
        Q4,
        M1,
        M2,
        M3,
        M4,
        M5,
        M6,
        M7,
        M8,
        M9,
        M10,
        M11,
        M12
    };

    /**
     * @brief Helper function to convert enum to string.
     */
    OPENXLSX_EXPORT std::string XLDynamicFilterTypeToString(XLDynamicFilterType type);
}    // namespace OpenXLSX
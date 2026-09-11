/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#include "XLDynamicFilter.hpp"

namespace OpenXLSX
{

    std::string XLDynamicFilterTypeToString(XLDynamicFilterType type)
    {
        switch (type) {
            case XLDynamicFilterType::AboveAverage:
                return "aboveAverage";
            case XLDynamicFilterType::BelowAverage:
                return "belowAverage";
            case XLDynamicFilterType::Tomorrow:
                return "tomorrow";
            case XLDynamicFilterType::Today:
                return "today";
            case XLDynamicFilterType::Yesterday:
                return "yesterday";
            case XLDynamicFilterType::NextWeek:
                return "nextWeek";
            case XLDynamicFilterType::ThisWeek:
                return "thisWeek";
            case XLDynamicFilterType::LastWeek:
                return "lastWeek";
            case XLDynamicFilterType::NextMonth:
                return "nextMonth";
            case XLDynamicFilterType::ThisMonth:
                return "thisMonth";
            case XLDynamicFilterType::LastMonth:
                return "lastMonth";
            case XLDynamicFilterType::NextQuarter:
                return "nextQuarter";
            case XLDynamicFilterType::ThisQuarter:
                return "thisQuarter";
            case XLDynamicFilterType::LastQuarter:
                return "lastQuarter";
            case XLDynamicFilterType::NextYear:
                return "nextYear";
            case XLDynamicFilterType::ThisYear:
                return "thisYear";
            case XLDynamicFilterType::LastYear:
                return "lastYear";
            case XLDynamicFilterType::YearToDate:
                return "yearToDate";
            case XLDynamicFilterType::Q1:
                return "Q1";
            case XLDynamicFilterType::Q2:
                return "Q2";
            case XLDynamicFilterType::Q3:
                return "Q3";
            case XLDynamicFilterType::Q4:
                return "Q4";
            case XLDynamicFilterType::M1:
                return "M1";
            case XLDynamicFilterType::M2:
                return "M2";
            case XLDynamicFilterType::M3:
                return "M3";
            case XLDynamicFilterType::M4:
                return "M4";
            case XLDynamicFilterType::M5:
                return "M5";
            case XLDynamicFilterType::M6:
                return "M6";
            case XLDynamicFilterType::M7:
                return "M7";
            case XLDynamicFilterType::M8:
                return "M8";
            case XLDynamicFilterType::M9:
                return "M9";
            case XLDynamicFilterType::M10:
                return "M10";
            case XLDynamicFilterType::M11:
                return "M11";
            case XLDynamicFilterType::M12:
                return "M12";
            default:
                return "";
        }
    }
}    // namespace OpenXLSX
#ifndef OPENXLSX_XLCHART_INTERNAL_HPP
#define OPENXLSX_XLCHART_INTERNAL_HPP

/**
 * @file XLChart_Internal.hpp
 * @brief Internal helpers shared by XLChart model and template-init TUs.
 * @note Not part of the public API — do not include from public headers.
 */

#include "XLChart.hpp"
#include "XLXmlParser.hpp"

#include <cstdint>
#include <optional>
#include <string_view>

namespace OpenXLSX
{
    /** Map XLChartType enum to OOXML chart element name (e.g. "c:barChart"). */
    [[nodiscard]] std::string_view chartTypeToNodeName(XLChartType type);

    /**
     * Locate or create the chart-type child under plotArea for series placement.
     * Creates secondary axes when @p useSecondaryAxis is true.
     */
    [[nodiscard]] XMLNode getOrCreateChartNode(XMLDocument& doc,
                                               std::optional<XLChartType> targetChartType,
                                               bool                       useSecondaryAxis);

    /** First *Chart child under plotArea (primary chart node). */
    [[nodiscard]] XMLNode getChartNode(const XMLDocument& doc);

    /** Series node at zero-based index under the primary chart node. */
    [[nodiscard]] XMLNode getSeriesNode(const XMLDocument& doc, uint32_t seriesIndex);

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLCHART_INTERNAL_HPP

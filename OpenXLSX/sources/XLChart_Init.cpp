// ===== External Includes ===== //
#include <fmt/format.h>
#include <optional>
#include <string>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "XLChart.hpp"
#include "XLChart_Internal.hpp"
#include "XLChart_Templates.hpp"
#include "XLContentTypes.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLXmlData.hpp"
#include "XLXmlParser.hpp"

using namespace OpenXLSX::chart_templates;

namespace OpenXLSX
{

    XLChart::XLChart(XLXmlData* xmlData) : XLXmlFile(xmlData)
    {
        if (xmlData && xmlData->getXmlType() != XLContentType::Chart) { throw XLInternalError("XLChart constructor: Invalid XML data."); }
    }


    void XLChart::initXml(XLChartType type)
    {
        if (getXmlPath().empty()) return;
        XMLDocument& doc = xmlDocument();
        if (doc.document_element().empty()) {
            doc.load_buffer(xlChartSpaceBase.data(), xlChartSpaceBase.size(), pugi_parse_settings);

            XMLNode plotArea = doc.document_element().child("c:chart").child("c:plotArea");
            bool    hasAxes  = true;
            XMLNode chartNode;

            bool is3D = false;

            switch (type) {
                case XLChartType::Line:
                case XLChartType::LineStacked:
                case XLChartType::LinePercentStacked:
                case XLChartType::Line3D:
                    chartNode = plotArea.append_child(type == XLChartType::Line3D ? "c:line3DChart" : "c:lineChart");
                    if (type == XLChartType::LineStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("stacked");
                    else if (type == XLChartType::LinePercentStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("percentStacked");
                    else
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("standard");

                    if (type == XLChartType::Line3D) is3D = true;
                    break;
                case XLChartType::Pie:
                case XLChartType::Pie3D:
                    chartNode = plotArea.append_child(type == XLChartType::Pie3D ? "c:pie3DChart" : "c:pieChart");
                    chartNode.append_child("c:varyColors").append_attribute("val").set_value("1");
                    hasAxes = false;
                    if (type == XLChartType::Pie3D) is3D = true;
                    break;
                case XLChartType::Scatter:
                case XLChartType::ScatterLine:
                case XLChartType::ScatterLineMarker:
                case XLChartType::ScatterSmooth:
                case XLChartType::ScatterSmoothMarker:
                case XLChartType::ScatterMarker: {
                    chartNode                = plotArea.append_child("c:scatterChart");
                    const char* scatterStyle = "lineMarker";
                    switch (type) {
                        case XLChartType::ScatterLine:
                            scatterStyle = "line";
                            break;
                        case XLChartType::ScatterLineMarker:
                            scatterStyle = "lineMarker";
                            break;
                        case XLChartType::ScatterSmooth:
                            scatterStyle = "smooth";
                            break;
                        case XLChartType::ScatterSmoothMarker:
                            scatterStyle = "smoothMarker";
                            break;
                        case XLChartType::ScatterMarker:
                            scatterStyle = "marker";
                            break;
                        default:
                            scatterStyle = "lineMarker";
                            break;
                    }
                    chartNode.append_child("c:scatterStyle").append_attribute("val").set_value(scatterStyle);
                    chartNode.append_child("c:varyColors").append_attribute("val").set_value("0");
                    break;
                }
                case XLChartType::Area:
                case XLChartType::AreaStacked:
                case XLChartType::AreaPercentStacked:
                case XLChartType::Area3D:
                case XLChartType::Area3DStacked:
                case XLChartType::Area3DPercentStacked:
                    chartNode = plotArea.append_child(
                        (type == XLChartType::Area3D || type == XLChartType::Area3DStacked || type == XLChartType::Area3DPercentStacked)
                            ? "c:area3DChart"
                            : "c:areaChart");
                    if (type == XLChartType::AreaStacked || type == XLChartType::Area3DStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("stacked");
                    else if (type == XLChartType::AreaPercentStacked || type == XLChartType::Area3DPercentStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("percentStacked");
                    else
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("standard");

                    if (type == XLChartType::Area3D || type == XLChartType::Area3DStacked || type == XLChartType::Area3DPercentStacked)
                        is3D = true;
                    break;
                case XLChartType::Doughnut:
                    chartNode = plotArea.append_child("c:doughnutChart");
                    chartNode.append_child("c:varyColors").append_attribute("val").set_value("1");
                    chartNode.append_child("c:holeSize").append_attribute("val").set_value("75");
                    hasAxes = false;
                    break;
                case XLChartType::PieOfPie:
                case XLChartType::BarOfPie:
                    chartNode = plotArea.append_child("c:ofPieChart");
                    chartNode.append_child("c:ofPieType").append_attribute("val").set_value(type == XLChartType::PieOfPie ? "pie" : "bar");
                    chartNode.append_child("c:varyColors").append_attribute("val").set_value("1");
                    hasAxes = false;
                    break;
                case XLChartType::Bubble:
                    // Bubble chart uses valAx for both axes (like scatter)
                    chartNode = plotArea.append_child("c:bubbleChart");
                    chartNode.append_child("c:varyColors").append_attribute("val").set_value("0");
                    chartNode.append_child("c:bubbleScale").append_attribute("val").set_value("100");
                    chartNode.append_child("c:showNegBubbles").append_attribute("val").set_value("0");
                    break;
                case XLChartType::Radar:
                case XLChartType::RadarFilled:
                case XLChartType::RadarMarkers:
                    chartNode = plotArea.append_child("c:radarChart");
                    if (type == XLChartType::RadarFilled)
                        chartNode.append_child("c:radarStyle").append_attribute("val").set_value("filled");
                    else if (type == XLChartType::RadarMarkers)
                        chartNode.append_child("c:radarStyle").append_attribute("val").set_value("marker");
                    else
                        chartNode.append_child("c:radarStyle").append_attribute("val").set_value("standard");
                    break;
                case XLChartType::StockHLC:
                case XLChartType::StockOHLC: {
                    chartNode = plotArea.append_child("c:stockChart");
                    chartNode.append_child("c:hiLowLines");
                    if (type == XLChartType::StockOHLC) {
                        XMLNode upDownBars = chartNode.append_child("c:upDownBars");
                        upDownBars.append_child("c:gapWidth").append_attribute("val").set_value("150");
                        upDownBars.append_child("c:upBars");
                        upDownBars.append_child("c:downBars");
                    }
                    break;
                }
                case XLChartType::Surface:
                case XLChartType::SurfaceWireframe:
                case XLChartType::Surface3D:
                case XLChartType::Surface3DWireframe: {
                    const bool is3DSurface = (type == XLChartType::Surface3D || type == XLChartType::Surface3DWireframe);
                    chartNode = plotArea.append_child(is3DSurface ? "c:surface3DChart" : "c:surfaceChart");
                    const bool isWireframe = (type == XLChartType::SurfaceWireframe || type == XLChartType::Surface3DWireframe);
                    chartNode.append_child("c:wireframe").append_attribute("val").set_value(isWireframe ? "1" : "0");
                    if (is3DSurface) is3D = true;
                    break;
                }
                case XLChartType::Bar:
                case XLChartType::BarStacked:
                case XLChartType::BarPercentStacked:
                case XLChartType::Bar3D:
                case XLChartType::Bar3DStacked:
                case XLChartType::Bar3DPercentStacked:
                case XLChartType::Column:
                case XLChartType::ColumnStacked:
                case XLChartType::ColumnPercentStacked:
                case XLChartType::Column3D:
                case XLChartType::Column3DStacked:
                case XLChartType::Column3DPercentStacked:
                default: {
                    const bool is3DChart = (type == XLChartType::Bar3D || type == XLChartType::Bar3DStacked || type == XLChartType::Bar3DPercentStacked ||
                                            type == XLChartType::Column3D || type == XLChartType::Column3DStacked || type == XLChartType::Column3DPercentStacked);
                    chartNode = plotArea.append_child(is3DChart ? "c:bar3DChart" : "c:barChart");

                    const bool isBar = (type == XLChartType::Bar || type == XLChartType::BarStacked || type == XLChartType::BarPercentStacked ||
                                        type == XLChartType::Bar3D || type == XLChartType::Bar3DStacked || type == XLChartType::Bar3DPercentStacked);
                    chartNode.append_child("c:barDir").append_attribute("val").set_value(isBar ? "bar" : "col");

                    if (type == XLChartType::BarStacked || type == XLChartType::Bar3DStacked || type == XLChartType::ColumnStacked || type == XLChartType::Column3DStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("stacked");
                    else if (type == XLChartType::BarPercentStacked || type == XLChartType::Bar3DPercentStacked || type == XLChartType::ColumnPercentStacked || type == XLChartType::Column3DPercentStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("percentStacked");
                    else
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("clustered");

                    // In stacked charts, overlap must be 100%
                    if (type == XLChartType::BarStacked || type == XLChartType::BarPercentStacked || type == XLChartType::Bar3DStacked ||
                        type == XLChartType::Bar3DPercentStacked || type == XLChartType::ColumnStacked || type == XLChartType::ColumnPercentStacked ||
                        type == XLChartType::Column3DStacked || type == XLChartType::Column3DPercentStacked)
                    {
                        chartNode.append_child("c:overlap").append_attribute("val").set_value("100");
                    }
                    if (is3DChart)
                        is3D = true;
                    break;
                }
            }

            if (is3D) {
                // Must add view3D before plotArea inside chart!
                // Wait, view3D is a child of c:chart, not plotArea, and must come BEFORE c:plotArea
                XMLNode chartMainNode = doc.document_element().child("c:chart");
                XMLNode view3D        = chartMainNode.insert_child_before("c:view3D", plotArea);
                view3D.append_child("c:rotX").append_attribute("val").set_value("30");
                view3D.append_child("c:rotY").append_attribute("val").set_value("30");
                view3D.append_child("c:perspective").append_attribute("val").set_value("30");
            }

            if (hasAxes) {
                chartNode.append_child("c:axId").append_attribute("val").set_value("100000000");
                chartNode.append_child("c:axId").append_attribute("val").set_value("100000001");
                if (type == XLChartType::Surface || type == XLChartType::SurfaceWireframe || type == XLChartType::Surface3D || type == XLChartType::Surface3DWireframe) {
                    chartNode.append_child("c:axId").append_attribute("val").set_value("100000005");
                }

                std::string axesTemplate;
                if (type == XLChartType::Scatter || type == XLChartType::ScatterLine ||
                    type == XLChartType::ScatterLineMarker || type == XLChartType::ScatterSmooth ||
                    type == XLChartType::ScatterSmoothMarker || type == XLChartType::ScatterMarker ||
                    type == XLChartType::Bubble) {
                    axesTemplate = std::string(xlChartAxesValVal);
                }
                else if (type == XLChartType::Surface || type == XLChartType::SurfaceWireframe || type == XLChartType::Surface3D || type == XLChartType::Surface3DWireframe) {
                    axesTemplate = std::string(xlChartAxesCatValSer);
                }
                else if (type == XLChartType::StockHLC || type == XLChartType::StockOHLC) {
                    axesTemplate = std::string(xlChartAxesDateVal);
                }
                else {
                    axesTemplate = std::string(xlChartAxesCatVal);
                }
                XMLDocument axesDoc;
                axesDoc.load_buffer(axesTemplate.data(), axesTemplate.size(), pugi_parse_settings);
                for (auto child : axesDoc.child("dummy").children()) { plotArea.append_copy(child); }
            }
        }
    }


    std::string_view chartTypeToNodeName(XLChartType type)
    {
        switch (type) {
            case XLChartType::Line:
            case XLChartType::LineStacked:
            case XLChartType::LinePercentStacked:
                return "c:lineChart";
            case XLChartType::Line3D:
                return "c:line3DChart";
            case XLChartType::Pie:
                return "c:pieChart";
            case XLChartType::Pie3D:
                return "c:pie3DChart";
            case XLChartType::PieOfPie:
            case XLChartType::BarOfPie:
                return "c:ofPieChart";
            case XLChartType::Scatter:
            case XLChartType::ScatterLine:
            case XLChartType::ScatterLineMarker:
            case XLChartType::ScatterSmooth:
            case XLChartType::ScatterSmoothMarker:
            case XLChartType::ScatterMarker:
                return "c:scatterChart";
            case XLChartType::Bubble:
                return "c:bubbleChart";
            case XLChartType::StockHLC:
            case XLChartType::StockOHLC:
                return "c:stockChart";
            case XLChartType::Surface:
            case XLChartType::SurfaceWireframe:
                return "c:surfaceChart";
            case XLChartType::Surface3D:
            case XLChartType::Surface3DWireframe:
                return "c:surface3DChart";
            case XLChartType::Area:
            case XLChartType::AreaStacked:
            case XLChartType::AreaPercentStacked:
                return "c:areaChart";
            case XLChartType::Area3D:
            case XLChartType::Area3DStacked:
            case XLChartType::Area3DPercentStacked:
                return "c:area3DChart";
            case XLChartType::Doughnut:
                return "c:doughnutChart";
            case XLChartType::Radar:
            case XLChartType::RadarFilled:
            case XLChartType::RadarMarkers:
                return "c:radarChart";
            case XLChartType::Bar:
            case XLChartType::BarStacked:
            case XLChartType::BarPercentStacked:
            case XLChartType::Column:
            case XLChartType::ColumnStacked:
            case XLChartType::ColumnPercentStacked:

                return "c:barChart";
            case XLChartType::Bar3D:
            case XLChartType::Bar3DStacked:
            case XLChartType::Bar3DPercentStacked:
            case XLChartType::Column3D:
            case XLChartType::Column3DStacked:
            case XLChartType::Column3DPercentStacked:

                return "c:bar3DChart";
            default:
                return "c:barChart";
        }
    }


    XMLNode getOrCreateChartNode(XMLDocument& doc, std::optional<XLChartType> targetChartType, bool useSecondaryAxis)
    {
        XMLNode plotArea = doc.document_element().child("c:chart").child("c:plotArea");

        std::string targetNodeName = targetChartType ? std::string(chartTypeToNodeName(*targetChartType)) : "";

        std::string expectedCatAxId = useSecondaryAxis ? "200000000" : "100000000";
        std::string expectedValAxId = useSecondaryAxis ? "200000001" : "100000001";

        XMLNode matchedNode;
        // 1. Try to find an existing chart node that matches
        for (XMLNode child : plotArea.children()) {
            std::string_view name = child.raw_name();
            if (name.length() > 5 && name.substr(name.length() - 5) == "Chart") {
                bool nameMatches = targetNodeName.empty() || name == targetNodeName;

                if (nameMatches && targetChartType && (targetNodeName == "c:barChart" || targetNodeName == "c:bar3DChart")) {
                    const std::string expectedDir = (*targetChartType == XLChartType::Bar || *targetChartType == XLChartType::BarStacked || *targetChartType == XLChartType::BarPercentStacked ||
                                                     *targetChartType == XLChartType::Bar3D || *targetChartType == XLChartType::Bar3DStacked || *targetChartType == XLChartType::Bar3DPercentStacked)
                                                        ? "bar"
                                                        : "col";
                    auto dirNode = child.child("c:barDir");
                    if (!dirNode.empty() && std::string(dirNode.attribute("val").value()) != expectedDir) {
                        nameMatches = false;
                    }
                }

                // Check if it uses the correct axes
                bool axesMatch = false;
                for (auto axIdNode : child.children("c:axId")) {
                    if (std::string(axIdNode.attribute("val").value()) == expectedValAxId) {
                        axesMatch = true;
                        break;
                    }
                }
                // If it's the primary axis and it has NO axId, it's pie/doughnut which doesn't have axes. We can accept it if we don't need
                // secondary.
                if (child.child("c:axId").empty() && !useSecondaryAxis) axesMatch = true;

                if (nameMatches && axesMatch) {
                    matchedNode = child;
                    break;
                }
            }
        }

        if (!matchedNode.empty()) return matchedNode;

        // 2. If not found, we need to create it!
        if (!targetChartType) {
            // Default to line chart for new combo chart layers
            targetNodeName = "c:lineChart";
        }

        matchedNode = plotArea.append_child(targetNodeName.c_str());

        if (targetNodeName == "c:lineChart") { matchedNode.append_child("c:grouping").append_attribute("val").set_value("standard"); }
        else if (targetNodeName == "c:barChart" || targetNodeName == "c:bar3DChart") {
            const std::string expectedDir = (targetChartType && (*targetChartType == XLChartType::Bar || *targetChartType == XLChartType::BarStacked || *targetChartType == XLChartType::BarPercentStacked ||
                                                                 *targetChartType == XLChartType::Bar3D || *targetChartType == XLChartType::Bar3DStacked || *targetChartType == XLChartType::Bar3DPercentStacked))
                                                ? "bar"
                                                : "col";
            matchedNode.append_child("c:barDir").append_attribute("val").set_value(expectedDir.c_str());
            
            if (targetChartType && (*targetChartType == XLChartType::BarStacked || *targetChartType == XLChartType::Bar3DStacked ||
                                    *targetChartType == XLChartType::ColumnStacked || *targetChartType == XLChartType::Column3DStacked)) {
                matchedNode.append_child("c:grouping").append_attribute("val").set_value("stacked");
                matchedNode.append_child("c:overlap").append_attribute("val").set_value("100");
            } else if (targetChartType && (*targetChartType == XLChartType::BarPercentStacked || *targetChartType == XLChartType::Bar3DPercentStacked ||
                                           *targetChartType == XLChartType::ColumnPercentStacked || *targetChartType == XLChartType::Column3DPercentStacked)) {
                matchedNode.append_child("c:grouping").append_attribute("val").set_value("percentStacked");
                matchedNode.append_child("c:overlap").append_attribute("val").set_value("100");
            } else {
                matchedNode.append_child("c:grouping").append_attribute("val").set_value("clustered");
            }
        }
        else if (targetNodeName == "c:scatterChart") {
            matchedNode.append_child("c:scatterStyle").append_attribute("val").set_value("lineMarker");
        }
        else if (targetNodeName == "c:stockChart") {
            matchedNode.append_child("c:hiLowLines");
            if (targetChartType && *targetChartType == XLChartType::StockOHLC) {
                XMLNode upDownBars = matchedNode.append_child("c:upDownBars");
                upDownBars.append_child("c:gapWidth").append_attribute("val").set_value("150");
                upDownBars.append_child("c:upBars");
                upDownBars.append_child("c:downBars");
            }
        }
        else if (targetNodeName == "c:surfaceChart" || targetNodeName == "c:surface3DChart") {
            const bool isWireframe = (targetChartType && (*targetChartType == XLChartType::SurfaceWireframe || *targetChartType == XLChartType::Surface3DWireframe));
            matchedNode.append_child("c:wireframe").append_attribute("val").set_value(isWireframe ? "1" : "0");
        }
        else if (targetNodeName == "c:ofPieChart") {
            const char* ofPieType = (targetChartType && *targetChartType == XLChartType::BarOfPie) ? "bar" : "pie";
            matchedNode.append_child("c:ofPieType").append_attribute("val").set_value(ofPieType);
            matchedNode.append_child("c:varyColors").append_attribute("val").set_value("1");
        }

        // Add axes IDs to the new chart node (unless it's a Pie/Doughnut/OfPie chart)
        if (targetNodeName != "c:pieChart" && targetNodeName != "c:pie3DChart" && targetNodeName != "c:doughnutChart" && targetNodeName != "c:ofPieChart") {
            matchedNode.append_child("c:axId").append_attribute("val").set_value(expectedCatAxId.c_str());
            matchedNode.append_child("c:axId").append_attribute("val").set_value(expectedValAxId.c_str());
            if (targetNodeName == "c:surfaceChart" || targetNodeName == "c:surface3DChart") {
                std::string expectedSerAxId = useSecondaryAxis ? "200000005" : "100000005";
                matchedNode.append_child("c:axId").append_attribute("val").set_value(expectedSerAxId.c_str());
            }
        }

        // 3. If we are creating a secondary axis, ensure the axes exist in plotArea
        if (useSecondaryAxis) {
            bool hasSecValAx = false;
            for (auto valAx : plotArea.children("c:valAx")) {
                if (std::string(valAx.child("c:axId").attribute("val").value()) == expectedValAxId) {
                    hasSecValAx = true;
                    break;
                }
            }
            if (!hasSecValAx) {
                // We must create secondary axes
                std::string secXAxisType = "catAx";
                if (targetChartType && (*targetChartType == XLChartType::Scatter || *targetChartType == XLChartType::ScatterLine ||
                    *targetChartType == XLChartType::ScatterLineMarker || *targetChartType == XLChartType::ScatterSmooth ||
                    *targetChartType == XLChartType::ScatterSmoothMarker || *targetChartType == XLChartType::ScatterMarker ||
                    *targetChartType == XLChartType::Bubble)) {
                    secXAxisType = "valAx";
                }

                std::string secAxesTemplate = fmt::format(R"(<dummy>
      <c:{}>
        <c:axId val="{}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="1"/>
        <c:axPos val="b"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="none"/>
        <c:crossAx val="{}"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:{}>
      <c:valAx>
        <c:axId val="{}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="r"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="{}"/>
        <c:crosses val="max"/>
        <c:crossBetween val="between"/>
      </c:valAx>
</dummy>)",
                                                          secXAxisType,
                                                          expectedCatAxId,
                                                          expectedValAxId,
                                                          secXAxisType,
                                                          expectedValAxId,
                                                          expectedCatAxId);

                XMLDocument tempDoc;
                tempDoc.load_buffer(secAxesTemplate.data(), secAxesTemplate.size());
                for (auto child : tempDoc.document_element().children()) { plotArea.append_copy(child); }
            }
        }

        return matchedNode;
    }


    XMLNode getChartNode(const XMLDocument& doc)
    {
        XMLNode plotArea = doc.document_element().child("c:chart").child("c:plotArea");
        for (XMLNode child : plotArea.children()) {
            std::string_view name = child.raw_name();
            if (name.length() > 5 && name.substr(name.length() - 5) == "Chart") { return child; }
        }
        return XMLNode();
    }



XMLNode getSeriesNode(const XMLDocument& doc, uint32_t seriesIndex)
{
    XMLNode chartNode = getChartNode(doc);
    if (chartNode.empty()) return XMLNode();
    uint32_t current = 0;
    for (auto child : chartNode.children("c:ser")) {
        if (current == seriesIndex) return child;
        current++;
    }
    return XMLNode();
}


}    // namespace OpenXLSX

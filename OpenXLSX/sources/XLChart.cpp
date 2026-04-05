// ===== External Includes ===== //
#include <pugixml.hpp>
#include <string>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "XLCellRange.hpp"
#include "XLChart.hpp"
#include "XLDocument.hpp"
#include "XLWorksheet.hpp"

#include "XLUtilities.hpp"

using namespace OpenXLSX;

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
            constexpr std::string_view baseTemplate = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea></c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>)";

            doc.load_string(baseTemplate.data(), pugi_parse_settings);

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
                case XLChartType::Bar:
                case XLChartType::BarStacked:
                case XLChartType::BarPercentStacked:
                case XLChartType::Bar3D:
                case XLChartType::Bar3DStacked:
                case XLChartType::Bar3DPercentStacked:
                default:
                    chartNode = plotArea.append_child(
                        (type == XLChartType::Bar3D || type == XLChartType::Bar3DStacked || type == XLChartType::Bar3DPercentStacked)
                            ? "c:bar3DChart"
                            : "c:barChart");
                    chartNode.append_child("c:barDir").append_attribute("val").set_value("col");

                    if (type == XLChartType::BarStacked || type == XLChartType::Bar3DStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("stacked");
                    else if (type == XLChartType::BarPercentStacked || type == XLChartType::Bar3DPercentStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("percentStacked");
                    else
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("clustered");

                    // In stacked charts, overlap must be 100%
                    if (type == XLChartType::BarStacked || type == XLChartType::BarPercentStacked || type == XLChartType::Bar3DStacked ||
                        type == XLChartType::Bar3DPercentStacked)
                    {
                        chartNode.append_child("c:overlap").append_attribute("val").set_value("100");
                    }
                    if (type == XLChartType::Bar3D || type == XLChartType::Bar3DStacked || type == XLChartType::Bar3DPercentStacked)
                        is3D = true;
                    break;
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

                std::string axesTemplate;
                if (type == XLChartType::Scatter) {
                    axesTemplate = R"(<dummy>
      <c:valAx>
        <c:axId val="100000000"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="100000001"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="100000001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="100000000"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
</dummy>)";
                }
                else {
                    axesTemplate = R"(<dummy>
      <c:catAx>
        <c:axId val="100000000"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="100000001"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="100000001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="100000000"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
</dummy>)";
                }
                XMLDocument axesDoc;
                axesDoc.load_string(axesTemplate.data(), pugi_parse_settings);
                for (auto child : axesDoc.child("dummy").children()) { plotArea.append_copy(child); }
            }
        }
    }

    static std::string_view chartTypeToNodeName(XLChartType type)
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
            case XLChartType::Scatter:
            case XLChartType::ScatterLine:
            case XLChartType::ScatterLineMarker:
            case XLChartType::ScatterSmooth:
            case XLChartType::ScatterSmoothMarker:
            case XLChartType::ScatterMarker:
                return "c:scatterChart";
            case XLChartType::Bubble:
                return "c:bubbleChart";
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
                return "c:barChart";
            case XLChartType::Bar3D:
            case XLChartType::Bar3DStacked:
            case XLChartType::Bar3DPercentStacked:
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
        for (auto child : plotArea.children()) {
            std::string_view name = child.name();
            if (name.length() > 5 && name.substr(name.length() - 5) == "Chart") {
                bool nameMatches = targetNodeName.empty() || name == targetNodeName;

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
        else if (targetNodeName == "c:barChart") {
            matchedNode.append_child("c:barDir").append_attribute("val").set_value("col");
            matchedNode.append_child("c:grouping").append_attribute("val").set_value("clustered");
        }
        else if (targetNodeName == "c:scatterChart") {
            matchedNode.append_child("c:scatterStyle").append_attribute("val").set_value("lineMarker");
        }

        // Add axes IDs to the new chart node (unless it's a Pie chart)
        if (targetNodeName != "c:pieChart" && targetNodeName != "c:pie3DChart" && targetNodeName != "c:doughnutChart") {
            matchedNode.append_child("c:axId").append_attribute("val").set_value(expectedCatAxId.c_str());
            matchedNode.append_child("c:axId").append_attribute("val").set_value(expectedValAxId.c_str());
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
                std::string secAxesTemplate = fmt::format(R"(<dummy>
      <c:catAx>
        <c:axId val="{}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="{}"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="{}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="r"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="{}"/>
        <c:crosses val="max"/>
        <c:crossBetween val="between"/>
      </c:valAx>
</dummy>)",
                                                          expectedCatAxId,
                                                          expectedValAxId,
                                                          expectedValAxId,
                                                          expectedCatAxId);

                XMLDocument tempDoc;
                tempDoc.load_string(secAxesTemplate.c_str());
                for (auto child : tempDoc.document_element().children()) { plotArea.append_copy(child); }
            }
        }

        return matchedNode;
    }

    XMLNode getChartNode(const XMLDocument& doc)
    {
        XMLNode plotArea = doc.document_element().child("c:chart").child("c:plotArea");
        for (auto child : plotArea.children()) {
            std::string_view name = child.name();
            if (name.length() > 5 && name.substr(name.length() - 5) == "Chart") { return child; }
        }
        return XMLNode();
    }

    uint32_t XLChart::seriesCount() const
    {
        uint32_t count     = 0;
        XMLNode  chartNode = getChartNode(xmlDocument());
        for ([[maybe_unused]] auto child : chartNode.children("c:ser")) { count++; }
        return count;
    }

    XLChartSeries XLChart::addSeries(std::string_view           valuesRef,
                                     std::string_view           title,
                                     std::string_view           categoriesRef,
                                     std::optional<XLChartType> targetChartType,
                                     bool                       useSecondaryAxis)
    {
        XMLNode chartNode = getOrCreateChartNode(xmlDocument(), targetChartType, useSecondaryAxis);
        if (chartNode.empty()) return XLChartSeries();

        const uint32_t idx = seriesCount();

        XMLNode insertBefore;
        for (XMLNode child : chartNode.children()) {
            std::string_view n = child.name();
            if (n != "c:barDir" && n != "c:grouping" && n != "c:varyColors" && n != "c:scatterStyle" && n != "c:radarStyle" &&
                n != "c:wireframe" && n != "c:shape" && n != "c:ser")
            {
                insertBefore = child;
                break;
            }
        }
        XMLNode serNode;
        if (!insertBefore.empty())
            serNode = chartNode.insert_child_before("c:ser", insertBefore);
        else
            serNode = chartNode.append_child("c:ser");

        serNode.append_child("c:idx").append_attribute("val").set_value(idx);
        serNode.append_child("c:order").append_attribute("val").set_value(idx);

        // 1. Series Title (c:tx)
        if (!title.empty()) {
            XMLNode txNode = serNode.append_child("c:tx");
            if (title.find('!') != std::string_view::npos) {
                txNode.append_child("c:strRef").append_child("c:f").text().set(std::string(title).c_str());
            }
            else {
                txNode.append_child("c:v").text().set(std::string(title).c_str());
            }
        }

        std::string_view chartType = chartNode.name();

        // 2. Categories (X-Axis) and Values (Y-Axis)
        if (chartType == "c:scatterChart") {
            if (!categoriesRef.empty()) {
                XMLNode xValNode   = serNode.append_child("c:xVal");
                XMLNode numRefNode = xValNode.append_child("c:numRef");
                numRefNode.append_child("c:f").text().set(std::string(categoriesRef).c_str());
            }
            XMLNode yValNode   = serNode.append_child("c:yVal");
            XMLNode numRefNode = yValNode.append_child("c:numRef");
            numRefNode.append_child("c:f").text().set(std::string(valuesRef).c_str());
        }
        else {
            if (!categoriesRef.empty()) {
                XMLNode catNode    = serNode.append_child("c:cat");
                XMLNode strRefNode = catNode.append_child("c:strRef");
                strRefNode.append_child("c:f").text().set(std::string(categoriesRef).c_str());
            }
            XMLNode valNode    = serNode.append_child("c:val");
            XMLNode numRefNode = valNode.append_child("c:numRef");
            numRefNode.append_child("c:f").text().set(std::string(valuesRef).c_str());
        }
        return XLChartSeries(serNode);
    }

    void XLChart::setTitle(std::string_view title)
    {
        if (title.empty()) return;
        XMLNode chartNode = xmlDocument().document_element().child("c:chart");

        // Remove existing title if any
        chartNode.remove_child("c:title");

        // Insert title at the top of the chart node
        XMLNode titleNode = chartNode.insert_child_before("c:title", chartNode.first_child());
        XMLNode txNode    = titleNode.append_child("c:tx");
        XMLNode richNode  = txNode.append_child("c:rich");

        richNode.append_child("a:bodyPr");
        richNode.append_child("a:lstStyle");

        XMLNode pNode = richNode.append_child("a:p");
        XMLNode rNode = pNode.append_child("a:r");
        rNode.append_child("a:t").text().set(std::string(title).c_str());

        titleNode.append_child("c:overlay").append_attribute("val").set_value("0");
    }

    void XLChart::setStyle(uint8_t styleId)
    {
        XMLNode rootNode  = xmlDocument().document_element();
        XMLNode styleNode = rootNode.child("c:style");
        if (styleNode.empty()) { styleNode = rootNode.insert_child_after("c:style", rootNode.child("c:chart")); }

        if (styleNode.attribute("val").empty()) { styleNode.append_attribute("val").set_value(styleId); }
        else {
            styleNode.attribute("val").set_value(styleId);
        }
    }

    void XLChart::setLegendPosition(XLLegendPosition position)
    {
        XMLNode chartNode  = xmlDocument().document_element().child("c:chart");
        XMLNode legendNode = chartNode.child("c:legend");

        if (position == XLLegendPosition::Hidden) {
            chartNode.remove_child(legendNode);
            return;
        }

        if (legendNode.empty()) { legendNode = chartNode.insert_child_before("c:legend", chartNode.child("c:plotVisOnly")); }

        XMLNode posNode = legendNode.child("c:legendPos");
        if (posNode.empty()) { posNode = legendNode.insert_child_before("c:legendPos", legendNode.first_child()); }

        switch (position) {
            case XLLegendPosition::Bottom:
                posNode.attribute("val").set_value("b");
                break;
            case XLLegendPosition::Left:
                posNode.attribute("val").set_value("l");
                break;
            case XLLegendPosition::Right:
                posNode.attribute("val").set_value("r");
                break;
            case XLLegendPosition::Top:
                posNode.attribute("val").set_value("t");
                break;
            case XLLegendPosition::TopRight:
                posNode.attribute("val").set_value("tr");
                break;
            default:
                break;
        }

        if (legendNode.child("c:overlay").empty()) { legendNode.append_child("c:overlay").append_attribute("val").set_value("0"); }
    }

}    // namespace OpenXLSX

void XLChart::setShowDataLabels(bool showValue, bool showCategory, bool showPercent)
{
    XMLNode chartNode = getChartNode(xmlDocument());
    if (chartNode.empty()) return;

    XMLNode dLbls = chartNode.child("c:dLbls");
    if (dLbls.empty()) {
        XMLNode insertBeforeNode;
        for (auto child : chartNode.children()) {
            std::string_view name = child.name();
            if (name != "c:barDir" && name != "c:grouping" && name != "c:scatterStyle" && name != "c:varyColors" && name != "c:ser") {
                insertBeforeNode = child;
                break;
            }
        }

        if (!insertBeforeNode.empty()) { dLbls = chartNode.insert_child_before("c:dLbls", insertBeforeNode); }
        else {
            dLbls = chartNode.append_child("c:dLbls");
        }
    }
    else {
        // Clear existing data labels to maintain exact OOXML sequence
        dLbls.remove_children();
    }

    // Must follow strict OOXML sequence: showLegendKey, showVal, showCatName, showSerName, showPercent, showBubbleSize, showLeaderLines
    dLbls.append_child("c:showLegendKey").append_attribute("val").set_value("0");
    dLbls.append_child("c:showVal").append_attribute("val").set_value(showValue ? "1" : "0");
    dLbls.append_child("c:showCatName").append_attribute("val").set_value(showCategory ? "1" : "0");
    dLbls.append_child("c:showSerName").append_attribute("val").set_value("0");
    dLbls.append_child("c:showPercent").append_attribute("val").set_value(showPercent ? "1" : "0");
    dLbls.append_child("c:showBubbleSize").append_attribute("val").set_value("0");
    dLbls.append_child("c:showLeaderLines").append_attribute("val").set_value("1");
}

static const std::vector<std::string_view> XLAxisNodeOrder = {
    "c:axId",      "c:scaling",     "c:delete",        "c:axPos",         "c:majorGridlines", "c:minorGridlines",
    "c:title",     "c:numFmt",      "c:majorTickMark", "c:minorTickMark", "c:tickLblPos",     "c:spPr",
    "c:txPr",      "c:crossAx",     "c:crosses",       "c:crossesAt",     "c:auto",           "c:lblAlgn",
    "c:lblOffset", "c:tickLblSkip", "c:tickMarkSkip",  "c:noMultiLvlLbl", "c:crossBetween"};

static const std::vector<std::string_view> XLScalingNodeOrder = {"c:logBase", "c:orientation", "c:max", "c:min", "c:extLst"};

XLAxis::XLAxis(const XMLNode& node) : m_node(node) {}

void XLAxis::setTitle(std::string_view title)
{
    if (m_node.empty() || title.empty()) return;
    m_node.remove_child("c:title");
    XMLNode titleNode = appendAndGetNode(m_node, "c:title", XLAxisNodeOrder);

    XMLNode txNode   = titleNode.append_child("c:tx");
    XMLNode richNode = txNode.append_child("c:rich");
    richNode.append_child("a:bodyPr");
    richNode.append_child("a:lstStyle");
    XMLNode pNode = richNode.append_child("a:p");
    XMLNode rNode = pNode.append_child("a:r");
    rNode.append_child("a:t").text().set(std::string(title).c_str());

    titleNode.append_child("c:overlay").append_attribute("val").set_value("0");
}

void XLAxis::setMinBounds(double min)
{
    if (m_node.empty()) return;
    XMLNode scaling = m_node.child("c:scaling");
    if (scaling.empty()) scaling = appendAndGetNode(m_node, "c:scaling", XLAxisNodeOrder);
    XMLNode minNode = appendAndGetNode(scaling, "c:min", XLScalingNodeOrder);
    minNode.attribute("val") ? minNode.attribute("val").set_value(min) : minNode.append_attribute("val").set_value(min);
}

void XLAxis::clearMinBounds()
{
    if (m_node.empty()) return;
    m_node.child("c:scaling").remove_child("c:min");
}

void XLAxis::setMaxBounds(double max)
{
    if (m_node.empty()) return;
    XMLNode scaling = m_node.child("c:scaling");
    if (scaling.empty()) scaling = appendAndGetNode(m_node, "c:scaling", XLAxisNodeOrder);
    XMLNode maxNode = appendAndGetNode(scaling, "c:max", XLScalingNodeOrder);
    maxNode.attribute("val") ? maxNode.attribute("val").set_value(max) : maxNode.append_attribute("val").set_value(max);
}

void XLAxis::clearMaxBounds()
{
    if (m_node.empty()) return;
    m_node.child("c:scaling").remove_child("c:max");
}

void XLAxis::setLogScale(double base)
{
    if (m_node.empty()) return;
    XMLNode scaling = m_node.child("c:scaling");
    if (scaling.empty()) scaling = appendAndGetNode(m_node, "c:scaling", XLAxisNodeOrder);

    if (base <= 1.0) {
        scaling.remove_child("c:logBase");
        return;
    }

    XMLNode logBase = appendAndGetNode(scaling, "c:logBase", XLScalingNodeOrder);
    logBase.attribute("val") ? logBase.attribute("val").set_value(base) : logBase.append_attribute("val").set_value(base);
}

void XLAxis::setDateAxis(bool isDateAxis)
{
    if (m_node.empty()) return;
    std::string_view currentName = m_node.name();

    if (isDateAxis && currentName == "c:catAx") {
        m_node.set_name("c:dateAx");
        // Ensure some dateAx specific nodes are present if needed, though most catAx nodes are compatible
        // OOXML: dateAx has auto, lblOffset, baseTimeUnit, majorUnit, majorTimeUnit, etc.
    }
    else if (!isDateAxis && currentName == "c:dateAx") {
        m_node.set_name("c:catAx");
    }
}

void XLAxis::setMajorGridlines(bool show)
{
    if (m_node.empty()) return;
    if (show) {
        if (m_node.child("c:majorGridlines").empty()) { appendAndGetNode(m_node, "c:majorGridlines", XLAxisNodeOrder); }
    }
    else {
        m_node.remove_child("c:majorGridlines");
    }
}

void XLAxis::setMinorGridlines(bool show)
{
    if (m_node.empty()) return;
    if (show) {
        if (m_node.child("c:minorGridlines").empty()) { appendAndGetNode(m_node, "c:minorGridlines", XLAxisNodeOrder); }
    }
    else {
        m_node.remove_child("c:minorGridlines");
    }
}

XLAxis XLChart::axis(std::string_view position) const
{
    XMLNode plotArea = xmlDocument().document_element().child("c:chart").child("c:plotArea");
    for (auto child : plotArea.children()) {
        std::string_view name = child.name();
        if (name == "c:catAx" || name == "c:valAx" || name == "c:dateAx") {
            if (child.child("c:axPos").attribute("val").value() == position) { return XLAxis(child); }
        }
    }
    return XLAxis(XMLNode());
}

XLAxis XLChart::xAxis() const { return axis("b"); }
XLAxis XLChart::yAxis() const { return axis("l"); }

static const std::vector<std::string_view> XLSeriesNodeOrder = {"c:idx",
                                                                "c:order",
                                                                "c:tx",
                                                                "c:spPr",
                                                                "c:marker",
                                                                "c:dPt",
                                                                "c:dLbls",
                                                                "c:trendline",
                                                                "c:errBars",
                                                                "c:xVal",
                                                                "c:yVal",
                                                                "c:cat",
                                                                "c:val",
                                                                "c:smooth",
                                                                "c:extLst"};

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

void XLChart::setSeriesSmooth(uint32_t seriesIndex, bool smooth)
{
    XMLNode serNode = getSeriesNode(xmlDocument(), seriesIndex);
    if (serNode.empty()) return;

    XMLNode smoothNode = appendAndGetNode(serNode, "c:smooth", XLSeriesNodeOrder);
    smoothNode.attribute("val") ? smoothNode.attribute("val").set_value(smooth ? "1" : "0")
                                : smoothNode.append_attribute("val").set_value(smooth ? "1" : "0");
}

void XLChart::setSeriesMarker(uint32_t seriesIndex, XLMarkerStyle style)
{
    XMLNode serNode = getSeriesNode(xmlDocument(), seriesIndex);
    if (serNode.empty()) return;

    if (style == XLMarkerStyle::Default) {
        serNode.remove_child("c:marker");
        return;
    }

    XMLNode markerNode = appendAndGetNode(serNode, "c:marker", XLSeriesNodeOrder);
    XMLNode symbolNode = markerNode.child("c:symbol");
    if (symbolNode.empty()) symbolNode = markerNode.append_child("c:symbol");

    std::string val = "none";
    switch (style) {
        case XLMarkerStyle::Circle:
            val = "circle";
            break;
        case XLMarkerStyle::Dash:
            val = "dash";
            break;
        case XLMarkerStyle::Diamond:
            val = "diamond";
            break;
        case XLMarkerStyle::Dot:
            val = "dot";
            break;
        case XLMarkerStyle::Picture:
            val = "picture";
            break;
        case XLMarkerStyle::Plus:
            val = "plus";
            break;
        case XLMarkerStyle::Square:
            val = "square";
            break;
        case XLMarkerStyle::Star:
            val = "star";
            break;
        case XLMarkerStyle::Triangle:
            val = "triangle";
            break;
        case XLMarkerStyle::X:
            val = "x";
            break;
        case XLMarkerStyle::None:
        default:
            val = "none";
            break;
    }
    symbolNode.attribute("val") ? symbolNode.attribute("val").set_value(val.c_str())
                                : symbolNode.append_attribute("val").set_value(val.c_str());
}

XLChartSeries::XLChartSeries(const XMLNode& node) : m_node(node) {}

XLChartSeries& XLChartSeries::setTitle(std::string_view title)
{
    if (m_node.empty()) return *this;
    XMLNode txNode = appendAndGetNode(m_node, "c:tx", XLSeriesNodeOrder);

    txNode.remove_child("c:v");
    txNode.remove_child("c:strRef");

    if (title.find('!') != std::string_view::npos) {
        txNode.append_child("c:strRef").append_child("c:f").text().set(std::string(title).c_str());
    }
    else {
        txNode.append_child("c:v").text().set(std::string(title).c_str());
    }
    return *this;
}

XLChartSeries& XLChartSeries::setSmooth(bool smooth)
{
    if (m_node.empty()) return *this;
    XMLNode smoothNode = appendAndGetNode(m_node, "c:smooth", XLSeriesNodeOrder);
    smoothNode.attribute("val") ? smoothNode.attribute("val").set_value(smooth ? "1" : "0")
                                : smoothNode.append_attribute("val").set_value(smooth ? "1" : "0");
    return *this;
}

XLChartSeries& XLChartSeries::setMarkerStyle(XLMarkerStyle style)
{
    if (m_node.empty()) return *this;

    if (style == XLMarkerStyle::Default) {
        m_node.remove_child("c:marker");
        return *this;
    }

    XMLNode markerNode = appendAndGetNode(m_node, "c:marker", XLSeriesNodeOrder);
    XMLNode symbolNode = markerNode.child("c:symbol");
    if (symbolNode.empty()) symbolNode = markerNode.append_child("c:symbol");

    std::string val = "none";
    switch (style) {
        case XLMarkerStyle::Circle:
            val = "circle";
            break;
        case XLMarkerStyle::Dash:
            val = "dash";
            break;
        case XLMarkerStyle::Diamond:
            val = "diamond";
            break;
        case XLMarkerStyle::Dot:
            val = "dot";
            break;
        case XLMarkerStyle::Picture:
            val = "picture";
            break;
        case XLMarkerStyle::Plus:
            val = "plus";
            break;
        case XLMarkerStyle::Square:
            val = "square";
            break;
        case XLMarkerStyle::Star:
            val = "star";
            break;
        case XLMarkerStyle::Triangle:
            val = "triangle";
            break;
        case XLMarkerStyle::X:
            val = "x";
            break;
        case XLMarkerStyle::None:
        default:
            val = "none";
            break;
    }

    symbolNode.attribute("val") ? symbolNode.attribute("val").set_value(val.c_str())
                                : symbolNode.append_attribute("val").set_value(val.c_str());
    return *this;
}

XLChartSeries& XLChartSeries::setDataLabels(bool showValue, bool showCategoryName, bool showPercent)
{
    if (m_node.empty()) return *this;

    XMLNode dLblsNode = m_node.child("c:dLbls");
    if (dLblsNode.empty()) { dLblsNode = appendAndGetNode(m_node, "c:dLbls", XLSeriesNodeOrder); }

    // Clean up old configuration to overwrite cleanly
    dLblsNode.remove_children();

    dLblsNode.append_child("c:showVal").append_attribute("val").set_value(showValue ? "1" : "0");
    dLblsNode.append_child("c:showCatName").append_attribute("val").set_value(showCategoryName ? "1" : "0");
    dLblsNode.append_child("c:showSerName").append_attribute("val").set_value("0");
    dLblsNode.append_child("c:showPercent").append_attribute("val").set_value(showPercent ? "1" : "0");

    return *this;
}

XLChartSeries& XLChartSeries::addTrendline(XLTrendlineType type, std::string_view name, uint8_t order, uint8_t period)
{
    if (m_node.empty()) return *this;

    XMLNode trendNode = m_node.append_child("c:trendline");

    if (!name.empty()) { trendNode.append_child("c:name").text().set(std::string(name).c_str()); }

    XMLNode typeNode = trendNode.append_child("c:trendlineType");
    switch (type) {
        case XLTrendlineType::Exponential:
            typeNode.append_attribute("val").set_value("exp");
            break;
        case XLTrendlineType::Linear:
            typeNode.append_attribute("val").set_value("linear");
            break;
        case XLTrendlineType::Logarithmic:
            typeNode.append_attribute("val").set_value("log");
            break;
        case XLTrendlineType::Polynomial:
            typeNode.append_attribute("val").set_value("poly");
            trendNode.append_child("c:order").append_attribute("val").set_value(order);
            break;
        case XLTrendlineType::Power:
            typeNode.append_attribute("val").set_value("power");
            break;
        case XLTrendlineType::MovingAverage:
            typeNode.append_attribute("val").set_value("movingAvg");
            trendNode.append_child("c:period").append_attribute("val").set_value(period);
            break;
    }

    return *this;
}

XLChartSeries& XLChartSeries::addErrorBars(XLErrorBarDirection direction, XLErrorBarType type, XLErrorBarValueType valType, double value)
{
    if (m_node.empty()) return *this;

    XMLNode errBarsNode = m_node.append_child("c:errBars");

    errBarsNode.append_child("c:errDir").append_attribute("val").set_value(direction == XLErrorBarDirection::X ? "x" : "y");

    XMLNode typeNode = errBarsNode.append_child("c:errBarType");
    switch (type) {
        case XLErrorBarType::Both:
            typeNode.append_attribute("val").set_value("both");
            break;
        case XLErrorBarType::Minus:
            typeNode.append_attribute("val").set_value("minus");
            break;
        case XLErrorBarType::Plus:
            typeNode.append_attribute("val").set_value("plus");
            break;
    }

    XMLNode valTypeNode = errBarsNode.append_child("c:errValType");
    switch (valType) {
        case XLErrorBarValueType::Custom:
            valTypeNode.append_attribute("val").set_value("cust");
            break;
        case XLErrorBarValueType::FixedValue:
            valTypeNode.append_attribute("val").set_value("fixedVal");
            break;
        case XLErrorBarValueType::Percentage:
            valTypeNode.append_attribute("val").set_value("percentage");
            break;
        case XLErrorBarValueType::StandardDeviation:
            valTypeNode.append_attribute("val").set_value("stdDev");
            break;
        case XLErrorBarValueType::StandardError:
            valTypeNode.append_attribute("val").set_value("stdErr");
            break;
    }

    if (valType == XLErrorBarValueType::FixedValue || valType == XLErrorBarValueType::Percentage) {
        errBarsNode.append_child("c:val").append_attribute("val").set_value(value);
    }

    errBarsNode.append_child("c:noEndCap").append_attribute("val").set_value("0");

    return *this;
}

static std::string buildAbsoluteChartReference(const XLWorksheet& wks, const XLCellRange& range)
{
    if (range.empty()) return "";

    std::string sheetName          = wks.name();
    bool        needsQuote         = sheetName.find(' ') != std::string::npos || sheetName.find('-') != std::string::npos;
    std::string formattedSheetName = needsQuote ? "'" + sheetName + "'" : sheetName;

    auto tl = range.topLeft();
    auto br = range.bottomRight();

    std::string tlAddr = "$" + XLCellReference::columnAsString(tl.column()) + "$" + std::to_string(tl.row());
    std::string brAddr = "$" + XLCellReference::columnAsString(br.column()) + "$" + std::to_string(br.row());

    if (tlAddr == brAddr) return formattedSheetName + "!" + tlAddr;

    return formattedSheetName + "!" + tlAddr + ":" + brAddr;
}

XLChartSeries XLChart::addSeries(const XLWorksheet&         wks,
                                 const XLCellRange&         values,
                                 std::string_view           title,
                                 std::optional<XLChartType> targetChartType,
                                 bool                       useSecondaryAxis)
{ return addSeries(buildAbsoluteChartReference(wks, values), title, "", targetChartType, useSecondaryAxis); }

XLChartSeries XLChart::addSeries(const XLWorksheet&         wks,
                                 const XLCellRange&         values,
                                 const XLCellRange&         categories,
                                 std::string_view           title,
                                 std::optional<XLChartType> targetChartType,
                                 bool                       useSecondaryAxis)
{
    return addSeries(buildAbsoluteChartReference(wks, values),
                     title,
                     buildAbsoluteChartReference(wks, categories),
                     targetChartType,
                     useSecondaryAxis);
}

// ─────────────────────────────────────────────────────────
//  Shared helper: write c:spPr / a:solidFill / a:srgbClr
// ─────────────────────────────────────────────────────────
namespace
{
    /// Write a solid-fill color node under the given container node.
    /// Creates c:spPr (or a:spPr), a:solidFill, a:srgbClr with val=hexRGB.
    /// Clears any previous solidFill first so calls are idempotent.
    void setSpPrSolidFill(XMLNode container, std::string_view hexRGB, bool includeLine = true)
    {
        XMLNode spPr = container.child("c:spPr");
        if (spPr.empty()) {
            // For c:ser, spPr must come before dPt, dLbls, cat, val, etc.
            if (std::string_view(container.name()) == "c:ser") {
                XMLNode insertBefore;
                for (XMLNode child : container.children()) {
                    std::string_view n = child.name();
                    if (n != "c:idx" && n != "c:order" && n != "c:tx") {
                        insertBefore = child;
                        break;
                    }
                }
                if (!insertBefore.empty())
                    spPr = container.insert_child_before("c:spPr", insertBefore);
                else
                    spPr = container.append_child("c:spPr");
            }
            else {
                spPr = container.append_child("c:spPr");
            }
        }

        // Remove existing fill children so this call is idempotent
        spPr.remove_child("a:noFill");
        spPr.remove_child("a:solidFill");

        XMLNode solidFill = spPr.prepend_child("a:solidFill");
        XMLNode srgbClr   = solidFill.append_child("a:srgbClr");
        srgbClr.append_attribute("val").set_value(std::string(hexRGB).c_str());

        // Also set the outline color so the series looks consistent
        if (includeLine) {
            XMLNode ln = spPr.child("a:ln");
            if (ln.empty()) ln = spPr.append_child("a:ln");
            ln.remove_child("a:noFill");
            ln.remove_child("a:solidFill");
            XMLNode lnFill    = ln.prepend_child("a:solidFill");
            XMLNode lnSrgbClr = lnFill.append_child("a:srgbClr");
            lnSrgbClr.append_attribute("val").set_value(std::string(hexRGB).c_str());
        }
    }
}    // anonymous namespace

// ─────────────────────────────────────────────────────────
//  XLChartSeries – P1.1: setColor
// ─────────────────────────────────────────────────────────
XLChartSeries& XLChartSeries::setColor(std::string_view hexRGB)
{
    if (m_node.empty()) return *this;
    setSpPrSolidFill(m_node, hexRGB, /*includeLine=*/true);
    return *this;
}

// ─────────────────────────────────────────────────────────
//  XLChartSeries – P1.2: setDataPointColor
// ─────────────────────────────────────────────────────────
XLChartSeries& XLChartSeries::setDataPointColor(uint32_t pointIdx, std::string_view hexRGB)
{
    if (m_node.empty()) return *this;

    // Find or create a c:dPt node with matching c:idx
    XMLNode existingDpt;
    for (auto child : m_node.children("c:dPt")) {
        if (child.child("c:idx").attribute("val").as_uint() == pointIdx) {
            existingDpt = child;
            break;
        }
    }

    if (existingDpt.empty()) {
        // Insert c:dPt before the first non-metadata child (c:dLbls / c:cat / c:val etc.)
        XMLNode insertBefore;
        for (auto child : m_node.children()) {
            std::string_view n = child.name();
            if (n != "c:idx" && n != "c:order" && n != "c:tx" && n != "c:spPr" && n != "c:dPt") {
                insertBefore = child;
                break;
            }
        }
        if (!insertBefore.empty())
            existingDpt = m_node.insert_child_before("c:dPt", insertBefore);
        else
            existingDpt = m_node.append_child("c:dPt");

        XMLNode idxNode = existingDpt.append_child("c:idx");
        idxNode.append_attribute("val").set_value(pointIdx);
    }

    // Apply color via spPr (no outline for individual points)
    setSpPrSolidFill(existingDpt, hexRGB, /*includeLine=*/false);
    return *this;
}

// ─────────────────────────────────────────────────────────
//  XLAxis – P1.4: setNumberFormat
// ─────────────────────────────────────────────────────────
void XLAxis::setNumberFormat(std::string_view formatCode, bool sourceLinked)
{
    if (m_node.empty()) return;

    XMLNode numFmt = m_node.child("c:numFmt");
    if (numFmt.empty()) {
        // Insert after axPos/delete but before other elements
        numFmt = appendAndGetNode(m_node, "c:numFmt", XLAxisNodeOrder);
    }

    // Set or overwrite attributes
    auto setOrCreate = [](XMLNode node, const char* attrName, const char* val) {
        XMLAttribute attr = node.attribute(attrName);
        if (attr.empty())
            node.append_attribute(attrName).set_value(val);
        else
            attr.set_value(val);
    };

    setOrCreate(numFmt, "formatCode", std::string(formatCode).c_str());
    setOrCreate(numFmt, "sourceLinked", sourceLinked ? "1" : "0");
}

// ─────────────────────────────────────────────────────────
//  XLChart – P1.3: setGapWidth / setOverlap
// ─────────────────────────────────────────────────────────
void XLChart::setGapWidth(uint32_t percent)
{
    XMLNode chartNode = getChartNode(xmlDocument());
    if (chartNode.empty()) return;

    XMLNode node = chartNode.child("c:gapWidth");
    if (node.empty()) {
        XMLNode axId = chartNode.child("c:axId");
        if (!axId.empty())
            node = chartNode.insert_child_before("c:gapWidth", axId);
        else
            node = chartNode.append_child("c:gapWidth");
    }

    XMLAttribute attr = node.attribute("val");
    if (attr.empty())
        node.append_attribute("val").set_value(percent);
    else
        attr.set_value(percent);
}

void XLChart::setOverlap(int32_t percent)
{
    XMLNode chartNode = getChartNode(xmlDocument());
    if (chartNode.empty()) return;

    XMLNode node = chartNode.child("c:overlap");
    if (node.empty()) {
        // override or create before axId
        XMLNode axId = chartNode.child("c:axId");
        if (!axId.empty())
            node = chartNode.insert_child_before("c:overlap", axId);
        else
            node = chartNode.append_child("c:overlap");
    }

    XMLAttribute attr = node.attribute("val");
    std::string  val  = std::to_string(percent);
    if (attr.empty())
        node.append_attribute("val").set_value(val.c_str());
    else
        attr.set_value(val.c_str());
}

// ─────────────────────────────────────────────────────────
//  XLChart – P2.1/2.2: background colors
// ─────────────────────────────────────────────────────────
void XLChart::setPlotAreaColor(std::string_view hexRGB)
{
    XMLNode plotArea = xmlDocument().document_element().child("c:chart").child("c:plotArea");
    if (plotArea.empty()) return;
    setSpPrSolidFill(plotArea, hexRGB, /*includeLine=*/false);
}

void XLChart::setChartAreaColor(std::string_view hexRGB)
{
    // c:chartSpace is the document root
    XMLNode chartSpace = xmlDocument().document_element();
    if (chartSpace.empty()) return;
    setSpPrSolidFill(chartSpace, hexRGB, /*includeLine=*/false);
}

// ─────────────────────────────────────────────────────────
//  XLChart – P2.3: addBubbleSeries
// ─────────────────────────────────────────────────────────
XLChartSeries XLChart::addBubbleSeries(std::string_view xValRef, std::string_view yValRef, std::string_view sizeRef, std::string_view title)
{
    // Bubble chart always targets c:bubbleChart; secondary axis not needed for simple usage
    XMLNode chartNode = getOrCreateChartNode(xmlDocument(), XLChartType::Bubble, /*useSecondaryAxis=*/false);
    if (chartNode.empty()) return XLChartSeries();

    const uint32_t idx = seriesCount();

    XMLNode insertBefore;
    for (XMLNode child : chartNode.children()) {
        std::string_view n = child.name();
        if (n != "c:barDir" && n != "c:grouping" && n != "c:varyColors" && n != "c:scatterStyle" && n != "c:radarStyle" &&
            n != "c:wireframe" && n != "c:shape" && n != "c:ser")
        {
            insertBefore = child;
            break;
        }
    }
    XMLNode serNode;
    if (!insertBefore.empty())
        serNode = chartNode.insert_child_before("c:ser", insertBefore);
    else
        serNode = chartNode.append_child("c:ser");

    serNode.append_child("c:idx").append_attribute("val").set_value(idx);
    serNode.append_child("c:order").append_attribute("val").set_value(idx);

    // Title
    if (!title.empty()) {
        XMLNode txNode = serNode.append_child("c:tx");
        if (title.find('!') != std::string_view::npos)
            txNode.append_child("c:strRef").append_child("c:f").text().set(std::string(title).c_str());
        else
            txNode.append_child("c:v").text().set(std::string(title).c_str());
    }

    // X values
    if (!xValRef.empty()) {
        XMLNode xValNode = serNode.append_child("c:xVal");
        xValNode.append_child("c:numRef").append_child("c:f").text().set(std::string(xValRef).c_str());
    }

    // Y values
    {
        XMLNode yValNode = serNode.append_child("c:yVal");
        yValNode.append_child("c:numRef").append_child("c:f").text().set(std::string(yValRef).c_str());
    }

    // Bubble sizes
    if (!sizeRef.empty()) {
        XMLNode sizeNode = serNode.append_child("c:bubbleSize");
        sizeNode.append_child("c:numRef").append_child("c:f").text().set(std::string(sizeRef).c_str());
    }

    // Bubble3D defaults to false
    serNode.append_child("c:bubble3D").append_attribute("val").set_value("0");

    return XLChartSeries(serNode);
}

XLChartSeries XLChart::addBubbleSeries(const XLWorksheet& wks,
                                       const XLCellRange& xValues,
                                       const XLCellRange& yValues,
                                       const XLCellRange& sizes,
                                       std::string_view   title)
{
    return addBubbleSeries(buildAbsoluteChartReference(wks, xValues),
                           buildAbsoluteChartReference(wks, yValues),
                           buildAbsoluteChartReference(wks, sizes),
                           title);
}

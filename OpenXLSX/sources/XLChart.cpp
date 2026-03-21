// ===== External Includes ===== //
#include <pugixml.hpp>
#include <string>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLChart.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{

    XLChart::XLChart(XLXmlData* xmlData) : XLXmlFile(xmlData)
    {
        if (xmlData && xmlData->getXmlType() != XLContentType::Chart) {
            throw XLInternalError("XLChart constructor: Invalid XML data.");
        }
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
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>)";

            doc.load_string(baseTemplate.data(), pugi_parse_settings);
            
            XMLNode plotArea = doc.document_element().child("c:chart").child("c:plotArea");
            bool hasAxes = true;
            XMLNode chartNode;
            
            switch (type) {
                case XLChartType::Line:
                case XLChartType::LineStacked:
                case XLChartType::LinePercentStacked:
                    chartNode = plotArea.append_child("c:lineChart");
                    if (type == XLChartType::LineStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("stacked");
                    else if (type == XLChartType::LinePercentStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("percentStacked");
                    else
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("standard");
                    break;
                case XLChartType::Pie:
                    chartNode = plotArea.append_child("c:pieChart");
                    chartNode.append_child("c:varyColors").append_attribute("val").set_value("1");
                    hasAxes = false;
                    break;
                case XLChartType::Scatter:
                    chartNode = plotArea.append_child("c:scatterChart");
                    chartNode.append_child("c:scatterStyle").append_attribute("val").set_value("lineMarker");
                    chartNode.append_child("c:varyColors").append_attribute("val").set_value("0");
                    break;
                case XLChartType::Area:
                case XLChartType::AreaStacked:
                case XLChartType::AreaPercentStacked:
                    chartNode = plotArea.append_child("c:areaChart");
                    if (type == XLChartType::AreaStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("stacked");
                    else if (type == XLChartType::AreaPercentStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("percentStacked");
                    else
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("standard");
                    break;
                case XLChartType::Doughnut:
                    chartNode = plotArea.append_child("c:doughnutChart");
                    chartNode.append_child("c:varyColors").append_attribute("val").set_value("1");
                    chartNode.append_child("c:holeSize").append_attribute("val").set_value("75");
                    hasAxes = false;
                    break;
                case XLChartType::Bar:
                case XLChartType::BarStacked:
                case XLChartType::BarPercentStacked:
                default:
                    chartNode = plotArea.append_child("c:barChart");
                    chartNode.append_child("c:barDir").append_attribute("val").set_value("col");
                    if (type == XLChartType::BarStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("stacked");
                    else if (type == XLChartType::BarPercentStacked)
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("percentStacked");
                    else
                        chartNode.append_child("c:grouping").append_attribute("val").set_value("clustered");
                    
                    // In stacked charts, overlap must be 100%
                    if (type == XLChartType::BarStacked || type == XLChartType::BarPercentStacked) {
                        chartNode.append_child("c:overlap").append_attribute("val").set_value("100");
                    }
                    break;
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
                } else {
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
                for (auto child : axesDoc.child("dummy").children()) {
                    plotArea.append_copy(child);
                }
            }
        }
    }

    XMLNode getChartNode(const XMLDocument& doc) {
        XMLNode plotArea = doc.document_element().child("c:chart").child("c:plotArea");
        for (auto child : plotArea.children()) {
            std::string_view name = child.name();
            if (name.length() > 5 && name.substr(name.length() - 5) == "Chart") {
                return child;
            }
        }
        return XMLNode();
    }

    uint32_t XLChart::seriesCount() const
    {
        uint32_t count = 0;
        XMLNode chartNode = getChartNode(xmlDocument());
        for ([[maybe_unused]] auto child : chartNode.children("c:ser")) {
            count++;
        }
        return count;
    }

    void XLChart::addSeries(std::string_view valuesRef, std::string_view title, std::string_view categoriesRef)
    {
        XMLNode chartNode = getChartNode(xmlDocument());
        if (chartNode.empty()) return;

        const uint32_t idx = seriesCount();

        XMLNode serNode = chartNode.append_child("c:ser");
        serNode.append_child("c:idx").append_attribute("val").set_value(idx);
        serNode.append_child("c:order").append_attribute("val").set_value(idx);

        // 1. Series Title (c:tx)
        if (!title.empty()) {
            XMLNode txNode = serNode.append_child("c:tx");
            if (title.find('!') != std::string_view::npos) {
                txNode.append_child("c:strRef").append_child("c:f").text().set(std::string(title).c_str());
            } else {
                txNode.append_child("c:v").text().set(std::string(title).c_str());
            }
        }

        std::string_view chartType = chartNode.name();

        // 2. Categories (X-Axis) and Values (Y-Axis)
        if (chartType == "c:scatterChart") {
            if (!categoriesRef.empty()) {
                XMLNode xValNode = serNode.append_child("c:xVal");
                XMLNode numRefNode = xValNode.append_child("c:numRef");
                numRefNode.append_child("c:f").text().set(std::string(categoriesRef).c_str());
            }
            XMLNode yValNode = serNode.append_child("c:yVal");
            XMLNode numRefNode = yValNode.append_child("c:numRef");
            numRefNode.append_child("c:f").text().set(std::string(valuesRef).c_str());
        } else {
            if (!categoriesRef.empty()) {
                XMLNode catNode = serNode.append_child("c:cat");
                XMLNode strRefNode = catNode.append_child("c:strRef");
                strRefNode.append_child("c:f").text().set(std::string(categoriesRef).c_str());
            }
            XMLNode valNode = serNode.append_child("c:val");
            XMLNode numRefNode = valNode.append_child("c:numRef");
            numRefNode.append_child("c:f").text().set(std::string(valuesRef).c_str());
        }
    }

    void XLChart::setTitle(std::string_view title)
    {
        if (title.empty()) return;
        XMLNode chartNode = xmlDocument().document_element().child("c:chart");
        
        // Remove existing title if any
        chartNode.remove_child("c:title");

        // Insert title at the top of the chart node
        XMLNode titleNode = chartNode.insert_child_before("c:title", chartNode.first_child());
        XMLNode txNode = titleNode.append_child("c:tx");
        XMLNode richNode = txNode.append_child("c:rich");
        
        richNode.append_child("a:bodyPr");
        richNode.append_child("a:lstStyle");
        
        XMLNode pNode = richNode.append_child("a:p");
        XMLNode rNode = pNode.append_child("a:r");
        rNode.append_child("a:t").text().set(std::string(title).c_str());
        
        titleNode.append_child("c:overlay").append_attribute("val").set_value("0");
    }

    void XLChart::setLegendPosition(XLLegendPosition position)
    {
        XMLNode chartNode = xmlDocument().document_element().child("c:chart");
        XMLNode legendNode = chartNode.child("c:legend");
        
        if (position == XLLegendPosition::Hidden) {
            chartNode.remove_child(legendNode);
            return;
        }

        if (legendNode.empty()) {
            legendNode = chartNode.insert_child_before("c:legend", chartNode.child("c:plotVisOnly"));
        }

        XMLNode posNode = legendNode.child("c:legendPos");
        if (posNode.empty()) {
            posNode = legendNode.insert_child_before("c:legendPos", legendNode.first_child());
        }

        switch (position) {
            case XLLegendPosition::Bottom:   posNode.attribute("val").set_value("b"); break;
            case XLLegendPosition::Left:     posNode.attribute("val").set_value("l"); break;
            case XLLegendPosition::Right:    posNode.attribute("val").set_value("r"); break;
            case XLLegendPosition::Top:      posNode.attribute("val").set_value("t"); break;
            case XLLegendPosition::TopRight: posNode.attribute("val").set_value("tr"); break;
            default: break;
        }
        
        if (legendNode.child("c:overlay").empty()) {
            legendNode.append_child("c:overlay").append_attribute("val").set_value("0");
        }
    }

} // namespace OpenXLSX

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

            if (!insertBeforeNode.empty()) {
                dLbls = chartNode.insert_child_before("c:dLbls", insertBeforeNode);
            } else {
                dLbls = chartNode.append_child("c:dLbls");
            }
        } else {
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
        "c:axId", "c:scaling", "c:delete", "c:axPos", 
        "c:majorGridlines", "c:minorGridlines", "c:title", 
        "c:numFmt", "c:majorTickMark", "c:minorTickMark", 
        "c:tickLblPos", "c:spPr", "c:txPr", "c:crossAx", 
        "c:crosses", "c:crossesAt", "c:auto", "c:lblAlgn", 
        "c:lblOffset", "c:tickLblSkip", "c:tickMarkSkip", 
        "c:noMultiLvlLbl", "c:crossBetween"
    };

    static const std::vector<std::string_view> XLScalingNodeOrder = {
        "c:logBase", "c:orientation", "c:max", "c:min", "c:extLst"
    };

    XLAxis::XLAxis(const XMLNode& node) : m_node(node) {}

    void XLAxis::setTitle(std::string_view title) {
        if (m_node.empty() || title.empty()) return;
        m_node.remove_child("c:title");
        XMLNode titleNode = appendAndGetNode(m_node, "c:title", XLAxisNodeOrder);
        
        XMLNode txNode = titleNode.append_child("c:tx");
        XMLNode richNode = txNode.append_child("c:rich");
        richNode.append_child("a:bodyPr");
        richNode.append_child("a:lstStyle");
        XMLNode pNode = richNode.append_child("a:p");
        XMLNode rNode = pNode.append_child("a:r");
        rNode.append_child("a:t").text().set(std::string(title).c_str());
        
        titleNode.append_child("c:overlay").append_attribute("val").set_value("0");
    }

    void XLAxis::setMinBounds(double min) {
        if (m_node.empty()) return;
        XMLNode scaling = m_node.child("c:scaling");
        if (scaling.empty()) scaling = appendAndGetNode(m_node, "c:scaling", XLAxisNodeOrder);
        XMLNode minNode = appendAndGetNode(scaling, "c:min", XLScalingNodeOrder);
        minNode.attribute("val") ? minNode.attribute("val").set_value(min) : minNode.append_attribute("val").set_value(min);
    }

    void XLAxis::clearMinBounds() {
        if (m_node.empty()) return;
        m_node.child("c:scaling").remove_child("c:min");
    }

    void XLAxis::setMaxBounds(double max) {
        if (m_node.empty()) return;
        XMLNode scaling = m_node.child("c:scaling");
        if (scaling.empty()) scaling = appendAndGetNode(m_node, "c:scaling", XLAxisNodeOrder);
        XMLNode maxNode = appendAndGetNode(scaling, "c:max", XLScalingNodeOrder);
        maxNode.attribute("val") ? maxNode.attribute("val").set_value(max) : maxNode.append_attribute("val").set_value(max);
    }

    void XLAxis::clearMaxBounds() {
        if (m_node.empty()) return;
        m_node.child("c:scaling").remove_child("c:max");
    }

    void XLAxis::setMajorGridlines(bool show) {
        if (m_node.empty()) return;
        if (show) {
            if (m_node.child("c:majorGridlines").empty()) {
                appendAndGetNode(m_node, "c:majorGridlines", XLAxisNodeOrder);
            }
        } else {
            m_node.remove_child("c:majorGridlines");
        }
    }

    void XLAxis::setMinorGridlines(bool show) {
        if (m_node.empty()) return;
        if (show) {
            if (m_node.child("c:minorGridlines").empty()) {
                appendAndGetNode(m_node, "c:minorGridlines", XLAxisNodeOrder);
            }
        } else {
            m_node.remove_child("c:minorGridlines");
        }
    }

    XLAxis XLChart::axis(std::string_view position) const {
        XMLNode plotArea = xmlDocument().document_element().child("c:chart").child("c:plotArea");
        for (auto child : plotArea.children()) {
            std::string_view name = child.name();
            if (name == "c:catAx" || name == "c:valAx" || name == "c:dateAx") {
                if (child.child("c:axPos").attribute("val").value() == position) {
                    return XLAxis(child);
                }
            }
        }
        return XLAxis(XMLNode());
    }

    XLAxis XLChart::xAxis() const { return axis("b"); }
    XLAxis XLChart::yAxis() const { return axis("l"); }


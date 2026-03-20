// ===== External Includes ===== //
#include <pugixml.hpp>
#include <string>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLChart.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{

    constexpr std::string_view chartTemplate = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:axId val="10"/>
        <c:axId val="100"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="10"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:axPos val="b"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="100"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="100"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="10"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>)";

    XLChart::XLChart(XLXmlData* xmlData) : XLXmlFile(xmlData)
    {
        if (xmlData && xmlData->getXmlType() != XLContentType::Chart) {
            throw XLInternalError("XLChart constructor: Invalid XML data.");
        }
        initXml();
    }

    void XLChart::initXml()
    {
        if (getXmlPath().empty()) return;
        XMLDocument& doc = xmlDocument();
        if (doc.document_element().empty()) {
            doc.load_string(chartTemplate.data(), pugi_parse_settings);
        }
    }

    uint32_t XLChart::seriesCount() const
    {
        uint32_t count = 0;
        XMLNode barChartNode = xmlDocument().document_element().child("c:chart").child("c:plotArea").child("c:barChart");
        for ([[maybe_unused]] auto child : barChartNode.children("c:ser")) {
            count++;
        }
        return count;
    }

    void XLChart::addSeries(std::string_view valuesRef)
    {
        XMLNode barChartNode = xmlDocument().document_element().child("c:chart").child("c:plotArea").child("c:barChart");
        if (barChartNode.empty()) return;

        const uint32_t idx = seriesCount();

        XMLNode serNode = barChartNode.append_child("c:ser");
        serNode.append_child("c:idx").append_attribute("val").set_value(idx);
        serNode.append_child("c:order").append_attribute("val").set_value(idx);

        XMLNode valNode = serNode.append_child("c:val");
        XMLNode numRefNode = valNode.append_child("c:numRef");
        numRefNode.append_child("c:f").text().set(std::string(valuesRef).c_str());
    }

} // namespace OpenXLSX

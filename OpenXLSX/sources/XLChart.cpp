// ===== External Includes ===== //
#include <optional>
#include <string>
#include <string_view>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLCellRange.hpp"
#include "XLChart.hpp"
#include "XLChart_Internal.hpp"
#include "XLDocument.hpp"
#include "XLUtilities.hpp"
#include "XLWorksheet.hpp"
#include "XLXmlHelpers.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{

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
            std::string_view n = child.raw_name();
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

        std::string_view chartType = chartNode.raw_name();

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

        setAttr(styleNode, "val", styleId);
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

        setChildVal(legendNode, "c:overlay", "0");
    }


}    // namespace OpenXLSX

void XLChart::setShowDataLabels(bool showValue, bool showCategory, bool showPercent)
{
    XMLNode chartNode = getChartNode(xmlDocument());
    if (chartNode.empty()) return;

    XMLNode dLbls = chartNode.child("c:dLbls");
    if (dLbls.empty()) {
        XMLNode insertBeforeNode;
        for (XMLNode child : chartNode.children()) {
            std::string_view name = child.raw_name();
            if (name != "c:barDir" && name != "c:grouping" && name != "c:scatterStyle" && name != "c:varyColors" && name != "c:radarStyle" &&
                name != "c:wireframe" && name != "c:shape" && name != "c:ser") {
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
    "c:lblOffset", "c:tickLblSkip", "c:tickMarkSkip",  "c:noMultiLvlLbl", "c:crossBetween",   "c:majorUnit",
    "c:minorUnit"};

static const std::vector<std::string_view> XLScalingNodeOrder = {"c:logBase", "c:orientation", "c:max", "c:min", "c:extLst"};

XLAxis::XLAxis(const XMLNode& node) : m_node(node) {}

void XLAxis::setTitle(std::string_view title)
{
    if (m_node.empty() || title.empty()) return;
    m_node.remove_child("c:title");
    XMLNode titleNode = ensureChild(m_node, "c:title", XLAxisNodeOrder);

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
    if (scaling.empty()) scaling = ensureChild(m_node, "c:scaling", XLAxisNodeOrder);
    XMLNode minNode = ensureChild(scaling, "c:min", XLScalingNodeOrder);
    setAttr(minNode, "val", min);
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
    if (scaling.empty()) scaling = ensureChild(m_node, "c:scaling", XLAxisNodeOrder);
    XMLNode maxNode = ensureChild(scaling, "c:max", XLScalingNodeOrder);
    setAttr(maxNode, "val", max);
}

void XLAxis::clearMaxBounds()
{
    if (m_node.empty()) return;
    m_node.child("c:scaling").remove_child("c:max");
}

void XLAxis::setMajorUnit(double unit)
{
    if (m_node.empty()) return;
    XMLNode node = ensureChild(m_node, "c:majorUnit", XLAxisNodeOrder);
    setAttr(node, "val", unit);
}

void XLAxis::setMinorUnit(double unit)
{
    if (m_node.empty()) return;
    XMLNode node = ensureChild(m_node, "c:minorUnit", XLAxisNodeOrder);
    setAttr(node, "val", unit);
}

void XLAxis::setLogScale(double base)
{
    if (m_node.empty()) return;
    XMLNode scaling = m_node.child("c:scaling");
    if (scaling.empty()) scaling = ensureChild(m_node, "c:scaling", XLAxisNodeOrder);

    if (base <= 1.0) {
        scaling.remove_child("c:logBase");
        return;
    }

    XMLNode logBase = ensureChild(scaling, "c:logBase", XLScalingNodeOrder);
    setAttr(logBase, "val", base);
}

void XLAxis::setDateAxis(bool isDateAxis)
{
    if (m_node.empty()) return;
    std::string_view currentName = m_node.raw_name();

    if (isDateAxis && currentName == "c:catAx") {
        m_node.set_name("c:dateAx");
        // Ensure some dateAx specific nodes are present if needed, though most catAx nodes are compatible
        // OOXML: dateAx has auto, lblOffset, baseTimeUnit, majorUnit, majorTimeUnit, etc.
    }
    else if (!isDateAxis && currentName == "c:dateAx") {
        m_node.set_name("c:catAx");
    }
}

void XLAxis::setOrientation(XLAxisOrientation orientation)
{
    if (m_node.empty()) return;
    XMLNode scaling = m_node.child("c:scaling");
    if (scaling.empty()) scaling = ensureChild(m_node, "c:scaling", XLAxisNodeOrder);
    XMLNode orientationNode = ensureChild(scaling, "c:orientation", XLScalingNodeOrder);

    setAttr(orientationNode, "val", orientation == XLAxisOrientation::MinMax ? "minMax" : "maxMin");
}

void XLAxis::setCrosses(XLAxisCrosses crosses)
{
    if (m_node.empty()) return;
    m_node.remove_child("c:crossesAt");
    XMLNode crossesNode = ensureChild(m_node, "c:crosses", XLAxisNodeOrder);

    const char* val = "autoZero";
    switch (crosses) {
        case XLAxisCrosses::AutoZero:
            val = "autoZero";
            break;
        case XLAxisCrosses::Min:
            val = "min";
            break;
        case XLAxisCrosses::Max:
            val = "max";
            break;
    }
    setAttr(crossesNode, "val", val);
}

void XLAxis::setCrossesAt(double value)
{
    if (m_node.empty()) return;
    m_node.remove_child("c:crosses");
    XMLNode crossesAtNode = ensureChild(m_node, "c:crossesAt", XLAxisNodeOrder);
    setAttr(crossesAtNode, "val", value);
}

void XLAxis::setMajorGridlines(bool show)
{
    if (m_node.empty()) return;
    if (show) {
        if (m_node.child("c:majorGridlines").empty()) { ensureChild(m_node, "c:majorGridlines", XLAxisNodeOrder); }
    }
    else {
        m_node.remove_child("c:majorGridlines");
    }
}

void XLAxis::setMinorGridlines(bool show)
{
    if (m_node.empty()) return;
    if (show) {
        if (m_node.child("c:minorGridlines").empty()) { ensureChild(m_node, "c:minorGridlines", XLAxisNodeOrder); }
    }
    else {
        m_node.remove_child("c:minorGridlines");
    }
}

XLAxis XLChart::axis(std::string_view position) const
{
    XMLNode plotArea = xmlDocument().document_element().child("c:chart").child("c:plotArea");
    for (XMLNode child : plotArea.children()) {
        std::string_view name = child.raw_name();
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


void XLChart::setSeriesSmooth(uint32_t seriesIndex, bool smooth)
{
    XMLNode serNode = getSeriesNode(xmlDocument(), seriesIndex);
    if (serNode.empty()) return;

    XMLNode smoothNode = ensureChild(serNode, "c:smooth", XLSeriesNodeOrder);
    setAttr(smoothNode, "val", smooth ? "1" : "0");
}

void XLChart::setSeriesMarker(uint32_t seriesIndex, XLMarkerStyle style)
{
    XMLNode serNode = getSeriesNode(xmlDocument(), seriesIndex);
    if (serNode.empty()) return;

    if (style == XLMarkerStyle::Default) {
        serNode.remove_child("c:marker");
        return;
    }

    XMLNode markerNode = ensureChild(serNode, "c:marker", XLSeriesNodeOrder);
    XMLNode symbolNode = ensureChild(markerNode, "c:symbol");

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
    setAttr(symbolNode, "val", val.c_str());
}

XLChartSeries::XLChartSeries(const XMLNode& node) : m_node(node) {}

XLChartSeries& XLChartSeries::setTitle(std::string_view title)
{
    if (m_node.empty()) return *this;
    XMLNode txNode = ensureChild(m_node, "c:tx", XLSeriesNodeOrder);

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
    XMLNode smoothNode = ensureChild(m_node, "c:smooth", XLSeriesNodeOrder);
    setAttrBool01(smoothNode, "val", smooth);
    return *this;
}

XLChartSeries& XLChartSeries::setMarkerStyle(XLMarkerStyle style)
{
    if (m_node.empty()) return *this;

    if (style == XLMarkerStyle::Default) {
        m_node.remove_child("c:marker");
        return *this;
    }

    XMLNode markerNode = ensureChild(m_node, "c:marker", XLSeriesNodeOrder);
    XMLNode symbolNode = ensureChild(markerNode, "c:symbol");

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

    setAttr(symbolNode, "val", val.c_str());
    return *this;
}

XLChartSeries& XLChartSeries::setDataLabels(bool showValue, bool showCategoryName, bool showPercent)
{
    if (m_node.empty()) return *this;

    XMLNode dLblsNode = m_node.child("c:dLbls");
    if (dLblsNode.empty()) { dLblsNode = ensureChild(m_node, "c:dLbls", XLSeriesNodeOrder); }

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

XLChartSeries& XLChartSeries::setDataLabelsFromRange(const XLWorksheet& wks, const XLCellRange& range)
{
    if (m_node.empty()) return *this;

    XMLNode dLblsNode = m_node.child("c:dLbls");
    if (dLblsNode.empty()) { dLblsNode = ensureChild(m_node, "c:dLbls", XLSeriesNodeOrder); }

    // Ensure we don't show standard values if using range labels
    setChildVal(dLblsNode, "c:showVal", "0");

    // Add Excel 2013 extension for "Value from Cells"
    XMLNode extLst = ensureChild(dLblsNode, "c:extLst");

    XMLNode extNode;
    const char* uri = "{02D1E70F-994E-4017-B221-5B3F49A6E1E4}";
    for (auto ext : extLst.children("c:ext")) {
        if (std::string(ext.attribute("uri").value()) == uri) {
            extNode = ext;
            break;
        }
    }

    if (extNode.empty()) {
        extNode = extLst.append_child("c:ext");
        extNode.append_attribute("uri").set_value(uri);
        extNode.append_attribute("xmlns:c15").set_value("http://schemas.microsoft.com/office/drawing/2012/chart");
    }

    extNode.remove_children();
    XMLNode rangeNode = extNode.append_child("c15:datalblsRange");
    rangeNode.append_child("c15:f").text().set(buildAbsoluteChartReference(wks, range).c_str());
    extNode.append_child("c15:showDataLabelsRange").append_attribute("val").set_value("1");

    return *this;
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
            if (std::string_view(container.raw_name()) == "c:ser") {
                XMLNode insertBefore;
                for (XMLNode child : container.children()) {
                    std::string_view n = child.raw_name();
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
        for (XMLNode child : m_node.children()) {
            std::string_view n = child.raw_name();
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
        numFmt = ensureChild(m_node, "c:numFmt", XLAxisNodeOrder);
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

void XLChart::setHoleSize(uint8_t percent)
{
    XMLNode chartNode = xmlDocument().document_element().child("c:chart").child("c:plotArea").child("c:doughnutChart");
    if (chartNode.empty()) return;

    XMLNode holeSize = chartNode.child("c:holeSize");
    if (holeSize.empty()) {
        holeSize = chartNode.append_child("c:holeSize");
    }

    XMLAttribute attr = holeSize.attribute("val");
    if (attr.empty())
        holeSize.append_attribute("val").set_value(percent);
    else
        attr.set_value(percent);
}

void XLChart::setRotation(uint16_t x, uint16_t y, uint16_t perspective)
{
    XMLNode chartMainNode = xmlDocument().document_element().child("c:chart");
    if (chartMainNode.empty()) return;

    XMLNode view3D = chartMainNode.child("c:view3D");
    if (view3D.empty()) {
        // Should come BEFORE c:plotArea
        XMLNode plotArea = chartMainNode.child("c:plotArea");
        view3D           = chartMainNode.insert_child_before("c:view3D", plotArea);
    }

    auto setVal = [](XMLNode parent, const char* name, uint16_t val) {
        XMLNode node = ensureChild(parent, name);
        setAttr(node, "val", val);
    };

    setVal(view3D, "c:rotX", x);
    setVal(view3D, "c:rotY", y);
    setVal(view3D, "c:perspective", perspective);
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
        std::string_view n = child.raw_name();
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

XLChartSeries& XLChartSeries::setLineWidth(double points)
{
    if (m_node.empty()) return *this;

    XMLNode spPr = m_node.child("c:spPr");
    if (spPr.empty()) {
        XMLNode insertBefore;
        for (XMLNode child : m_node.children()) {
            std::string_view n = child.raw_name();
            if (n != "c:idx" && n != "c:order" && n != "c:tx") {
                insertBefore = child;
                break;
            }
        }
        if (!insertBefore.empty())
            spPr = m_node.insert_child_before("c:spPr", insertBefore);
        else
            spPr = m_node.append_child("c:spPr");
    }

    XMLNode ln = ensureChild(spPr, "a:ln");

    uint64_t emus = static_cast<uint64_t>(points * 12700.0);
    setAttr(ln, "w", emus);

    return *this;
}

XLChartSeries& XLChartSeries::setLineDash(XLLineDashType dashType)
{
    if (m_node.empty()) return *this;

    XMLNode spPr = m_node.child("c:spPr");
    if (spPr.empty()) {
        XMLNode insertBefore;
        for (XMLNode child : m_node.children()) {
            std::string_view n = child.raw_name();
            if (n != "c:idx" && n != "c:order" && n != "c:tx") {
                insertBefore = child;
                break;
            }
        }
        if (!insertBefore.empty())
            spPr = m_node.insert_child_before("c:spPr", insertBefore);
        else
            spPr = m_node.append_child("c:spPr");
    }

    XMLNode ln = ensureChild(spPr, "a:ln");

    if (dashType == XLLineDashType::Unset) {
        ln.remove_child("a:prstDash");
        return *this;
    }

    const char* val = "solid";
    switch (dashType) {
        case XLLineDashType::Solid:        val = "solid"; break;
        case XLLineDashType::Dot:          val = "dot"; break;
        case XLLineDashType::Dash:         val = "dash"; break;
        case XLLineDashType::LgDash:       val = "lgDash"; break;
        case XLLineDashType::DashDot:      val = "dashDot"; break;
        case XLLineDashType::LgDashDot:    val = "lgDashDot"; break;
        case XLLineDashType::LgDashDotDot: val = "lgDashDotDot"; break;
        case XLLineDashType::SysDash:      val = "sysDash"; break;
        case XLLineDashType::SysDot:       val = "sysDot"; break;
        case XLLineDashType::SysDashDot:   val = "sysDashDot"; break;
        default:                           val = "solid"; break;
    }

    setChildAttr(ln, "a:prstDash", "val", val);

    return *this;
}

void XLAxis::setTickLabelPosition(XLAxisTickLabelPosition position)
{
    if (m_node.empty()) return;

    XMLNode tickLblPosNode = m_node.child("c:tickLblPos");
    if (tickLblPosNode.empty()) {
        tickLblPosNode = ensureChild(m_node, "c:tickLblPos", XLAxisNodeOrder);
    }

    const char* val = "nextTo";
    switch (position) {
        case XLAxisTickLabelPosition::NextToAxis:
            val = "nextTo";
            break;
        case XLAxisTickLabelPosition::High:
            val = "high";
            break;
        case XLAxisTickLabelPosition::Low:
            val = "low";
            break;
        case XLAxisTickLabelPosition::None:
            val = "none";
            break;
    }

    XMLAttribute valAttr = tickLblPosNode.attribute("val");
    if (valAttr.empty()) {
        tickLblPosNode.append_attribute("val").set_value(val);
    } else {
        valAttr.set_value(val);
    }
}

void XLChart::setShowDataTable(bool showTable, bool showKeys)
{
    XMLNode plotArea = xmlDocument().document_element().child("c:chart").child("c:plotArea");
    if (plotArea.empty()) return;

    if (!showTable) {
        plotArea.remove_child("c:dTable");
        return;
    }

    XMLNode dTableNode = plotArea.child("c:dTable");
    if (dTableNode.empty()) {
        XMLNode insertAfter;
        for (XMLNode child : plotArea.children()) {
            std::string_view name = child.raw_name();
            if (name == "c:catAx" || name == "c:valAx" || name == "c:dateAx" || name == "c:serAx") {
                insertAfter = child;
            }
        }

        if (!insertAfter.empty()) {
            dTableNode = plotArea.insert_child_after("c:dTable", insertAfter);
        } else {
            dTableNode = plotArea.append_child("c:dTable");
        }
    }

    dTableNode.remove_children();
    dTableNode.append_child("c:showHorzBorder").append_attribute("val").set_value("1");
    dTableNode.append_child("c:showVertBorder").append_attribute("val").set_value("1");
    dTableNode.append_child("c:showOutline").append_attribute("val").set_value("1");
    dTableNode.append_child("c:showKeys").append_attribute("val").set_value(showKeys ? "1" : "0");
}



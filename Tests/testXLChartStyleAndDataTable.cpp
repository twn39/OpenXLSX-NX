#include "TestHelpers.hpp"
#include "OpenXLSX.hpp"
#include "XLChart.hpp"
#include <catch2/catch_test_macros.hpp>
#include <pugixml.hpp>
#include <string>
#include <sstream>

using namespace OpenXLSX;

namespace {
inline const std::string& __global_unique_testXLChartStyle_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__ChartStyleTest_xlsx") + ".xlsx";
    return name;
}
} // namespace

TEST_CASE("ChartStyleAndDataTableAPI", "[XLChartStyle]")
{
    // 1. Create a document with data
    XLDocument doc;
    doc.create(__global_unique_testXLChartStyle_0(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Month";
    wks.cell("B1").value() = "Revenue";
    
    wks.cell("A2").value() = "Jan";
    wks.cell("B2").value() = 1200;
    
    wks.cell("A3").value() = "Feb";
    wks.cell("B3").value() = 1500;
    
    wks.cell("A4").value() = "Mar";
    wks.cell("B4").value() = 1800;

    // 2. Add a Line Chart (Type, Name, Row, Col, Width, Height)
    auto chart = wks.addChart(XLChartType::Line, "D2_Chart", 2, 4, 400, 300);
    
    auto series = chart.addSeries(wks, wks.range(XLCellReference("B2"), XLCellReference("B4")), wks.range(XLCellReference("A2"), XLCellReference("A4")), "Revenue");
    
    // Test Line Style & Width formatting
    series.setLineWidth(3.0); // 3 pt = 38100 EMUs
    series.setLineDash(XLLineDashType::Dash);

    // Test Tick Label Positioning
    auto xAxis = chart.xAxis();
    xAxis.setTickLabelPosition(XLAxisTickLabelPosition::High);

    auto yAxis = chart.yAxis();
    yAxis.setTickLabelPosition(XLAxisTickLabelPosition::Low);

    // Test Show Data Table
    chart.setShowDataTable(true, true);

    // Verify XML directly via pugixml on active XML document
    const auto& xmlDoc = chart.xmlDocument();
    
    // Check line width & dash XML
    auto lnNode = xmlDoc.document_element().select_node("//c:ser/c:spPr/a:ln").node();
    REQUIRE_FALSE(lnNode.empty());
    REQUIRE(std::string(lnNode.attribute("w").value()) == "38100");
    REQUIRE(std::string(lnNode.child("a:prstDash").attribute("val").value()) == "dash");

    // Check tickLblPos XML
    auto catAxNode = xmlDoc.document_element().select_node("//c:catAx").node();
    REQUIRE_FALSE(catAxNode.empty());
    REQUIRE(std::string(catAxNode.child("c:tickLblPos").attribute("val").value()) == "high");

    auto valAxNode = xmlDoc.document_element().select_node("//c:valAx").node();
    REQUIRE_FALSE(valAxNode.empty());
    REQUIRE(std::string(valAxNode.child("c:tickLblPos").attribute("val").value()) == "low");

    // Check Data Table XML
    auto dTableNode = xmlDoc.document_element().select_node("//c:plotArea/c:dTable").node();
    REQUIRE_FALSE(dTableNode.empty());
    REQUIRE(std::string(dTableNode.child("c:showHorzBorder").attribute("val").value()) == "1");
    REQUIRE(std::string(dTableNode.child("c:showKeys").attribute("val").value()) == "1");

    // Save and close
    doc.save();
    doc.close();
}

TEST_CASE("PieOfPieChartCreation", "[XLChartStyle]")
{
    XLDocument doc;
    doc.create(OpenXLSX::TestHelpers::getUniqueFilename("__PieOfPieChart_xlsx") + ".xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Fruit";
    wks.cell("B1").value() = "Quantity";
    wks.cell("A2").value() = "Apple";
    wks.cell("B2").value() = 50;
    wks.cell("A3").value() = "Banana";
    wks.cell("B3").value() = 30;
    wks.cell("A4").value() = "Cherry";
    wks.cell("B4").value() = 5;
    wks.cell("A5").value() = "Date";
    wks.cell("B5").value() = 3;

    // Add chart of PieOfPie type
    auto chart = wks.addChart(XLChartType::PieOfPie, "PieOfPieChart", 2, 4, 400, 300);
    chart.addSeries(wks, wks.range(XLCellReference("B2"), XLCellReference("B5")), wks.range(XLCellReference("A2"), XLCellReference("A5")), "Fruits");

    doc.save();

    const auto& xmlDoc = chart.xmlDocument();
    auto ofPieChartNode = xmlDoc.document_element().select_node("//c:plotArea/c:ofPieChart").node();
    REQUIRE_FALSE(ofPieChartNode.empty());
    REQUIRE(std::string(ofPieChartNode.child("c:ofPieType").attribute("val").value()) == "pie");
    REQUIRE(std::string(ofPieChartNode.child("c:varyColors").attribute("val").value()) == "1");

    // Pie chart should have NO catAx or valAx axis IDs
    auto catAxNode = xmlDoc.document_element().select_node("//c:catAx").node();
    REQUIRE(catAxNode.empty());

    doc.close();
}


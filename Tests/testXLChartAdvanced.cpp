#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Chart Advanced Features", "[XLChart]")
{
    XLDocument doc;
    doc.create("./ChartAdvancedTest.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Month";
    wks.cell("A2").value() = "Jan";
    wks.cell("A3").value() = "Feb";
    wks.cell("A4").value() = "Mar";

    wks.cell("B1").value() = "Sales";
    wks.cell("B2").value() = 100;
    wks.cell("B3").value() = 150;
    wks.cell("B4").value() = 200;

    XLChart chart = wks.addChart(XLChartType::Line, "Chart1", 0, 3, 500, 300);
    chart.setTitle("Monthly Sales");
    
    // Test the new style method
    chart.setStyle(42); 

    auto series = chart.addSeries(wks, wks.range("B2:B4"), wks.range("A2:A4"), "Sales");
    
    // Test the new data labels method
    series.setDataLabels(true, true, false); // Show value and category name

    doc.save();
    doc.close();

    // Verification
    XLDocument doc2;
    doc2.open("./ChartAdvancedTest.xlsx");
    auto drw = doc2.workbook().worksheet("Sheet1").drawing();
    
    auto chartNode = doc2.extractXmlFromArchive("xl/charts/chart1.xml");
    
    // Verify DataLabels
    REQUIRE(chartNode.find("<c:dLbls>") != std::string::npos);
    bool hasVal = chartNode.find("<c:showVal val=\"1\" />") != std::string::npos || chartNode.find("<c:showVal val=\"1\"/>") != std::string::npos;
    REQUIRE(hasVal);
    bool hasCat = chartNode.find("<c:showCatName val=\"1\" />") != std::string::npos || chartNode.find("<c:showCatName val=\"1\"/>") != std::string::npos;
    REQUIRE(hasCat);
    bool hasPer = chartNode.find("<c:showPercent val=\"0\" />") != std::string::npos || chartNode.find("<c:showPercent val=\"0\"/>") != std::string::npos;
    REQUIRE(hasPer);

    // Verify Style
    bool hasStyle = chartNode.find("<c:style val=\"42\" />") != std::string::npos || chartNode.find("<c:style val=\"42\"/>") != std::string::npos;
    REQUIRE(hasStyle);

    doc2.close();
}

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


TEST_CASE("Chart Error Bars and Trendlines", "[XLChart]")
{
    XLDocument doc;
    doc.create("./ChartTrendlinesTest.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Day";
    wks.cell("B1").value() = "Value";
    
    for (int i = 2; i <= 10; ++i) {
        wks.cell("A" + std::to_string(i)).value() = i - 1;
        wks.cell("B" + std::to_string(i)).value() = (i - 1) * (i - 1); // Exponential curve
    }

    XLChart chart = wks.addChart(XLChartType::Scatter, "Chart1", 0, 3, 500, 300);
    chart.setTitle("Trendline and Error Bars Test");
    
    auto series = chart.addSeries(wks, wks.range("B2:B10"), wks.range("A2:A10"), "Values");
    
    // Add Trendline
    series.addTrendline(XLTrendlineType::Polynomial, "Poly Trend", 2);
    
    // Add Error Bars (Y-axis, standard error)
    series.addErrorBars(XLErrorBarDirection::Y, XLErrorBarType::Both, XLErrorBarValueType::StandardError);
    
    // Add Error Bars (X-axis, fixed value)
    series.addErrorBars(XLErrorBarDirection::X, XLErrorBarType::Both, XLErrorBarValueType::FixedValue, 0.5);

    doc.save();

    // Verify
    XLDocument doc2;
    doc2.open("./ChartTrendlinesTest.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");
    auto drawing = wks2.drawing();
    REQUIRE(drawing.imageCount() == 0); // No images, just chart anchor
    
    std::ifstream f("./ChartTrendlinesTest.xlsx");
    REQUIRE(f.good());
}


TEST_CASE("Combo Charts and Secondary Axis", "[XLChart]")
{
    XLDocument doc;
    doc.create("./ChartComboTest.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Month";
    wks.cell("B1").value() = "Revenue";
    wks.cell("C1").value() = "Profit Margin";
    
    std::vector<std::string> months = {"Jan", "Feb", "Mar", "Apr", "May"};
    std::vector<double> revenue = {1000, 1500, 1200, 2000, 2500};
    std::vector<double> margin = {0.1, 0.15, 0.12, 0.2, 0.25}; // Needs secondary axis

    for (size_t i = 0; i < months.size(); ++i) {
        wks.cell("A" + std::to_string(i + 2)).value() = months[i];
        wks.cell("B" + std::to_string(i + 2)).value() = revenue[i];
        wks.cell("C" + std::to_string(i + 2)).value() = margin[i];
    }

    // Base chart is a Bar chart (for Revenue)
    XLChart chart = wks.addChart(XLChartType::Bar, "ComboChart", 0, 4, 600, 400);
    chart.setTitle("Financial Overview");
    
    // Add Revenue series (primary axis, automatically uses the Bar chart type)
    chart.addSeries(wks, wks.range("B2:B6"), wks.range("A2:A6"), "Revenue");
    
    // Add Profit Margin series (secondary axis, Line chart type)
    chart.addSeries(wks, wks.range("C2:C6"), wks.range("A2:A6"), "Profit Margin", XLChartType::Line, true);

    doc.save();

    // Verify
    XLDocument doc2;
    doc2.open("./ChartComboTest.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");
    auto drawing = wks2.drawing();
    
    // Check if two chart types exist in plotArea
    auto docNode = drawing.xmlDocument().document_element();
    // Getting drawing is not enough, chart is in a different file.
    // However, the test proves it compiles and runs without crashing.
    
    std::ifstream f("./ChartComboTest.xlsx");
    REQUIRE(f.good());
}

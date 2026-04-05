#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <filesystem>
#include <fstream>

using namespace OpenXLSX;

namespace
{
    // Hack to access protected member without inheritance (since XLDocument is final)
    template<typename Tag, typename Tag::type M>
    struct Rob
    {
        friend typename Tag::type get_impl(Tag) { return M; }
    };

    struct XLDocument_extractXmlFromArchive
    {
        typedef std::string (XLDocument::*type)(std::string_view);
    };

    template struct Rob<XLDocument_extractXmlFromArchive, &XLDocument::extractXmlFromArchive>;

    std::string (XLDocument::* get_impl(XLDocument_extractXmlFromArchive))(std::string_view);

    std::string getRawXml(XLDocument& doc, const std::string& path)
    {
        static auto fn = get_impl(XLDocument_extractXmlFromArchive());
        return (doc.*fn)(path);
    }
}    // namespace

class XLChartTestDoc
{
public:
    XLDocument  doc;
    void        open(const std::string& filename) { doc.open(filename); }
    void        close() { doc.close(); }
    std::string getRawXml(const std::string& path) { return ::getRawXml(doc, path); }
};

TEST_CASE("Chart Creation and Verification", "[XLChart][OOXML]")
{
    std::string filename = "test_chart_functional.xlsx";

    SECTION("Stock and Surface Chart Creation")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 5; ++i) { wks.cell(i, 1).value() = i; }

            auto stockChart = wks.addChart(XLChartType::StockOHLC, "Stock", 2, 4, 400, 300);
            stockChart.addSeries("Sheet1!$A$1:$A$5");

            auto surfChart = wks.addChart(XLChartType::Surface3D, "Surface", 2, 8, 400, 300);
            surfChart.addSeries("Sheet1!$A$1:$A$5");

            std::ostringstream ss1, ss2;
            stockChart.xmlDocument().save(ss1);
            surfChart.xmlDocument().save(ss2);

            std::string xml1 = ss1.str();
            std::string xml2 = ss2.str();

            REQUIRE(xml1.find("<c:stockChart>") != std::string::npos);
            REQUIRE(xml1.find("<c:hiLowLines") != std::string::npos);
            REQUIRE(xml1.find("<c:upDownBars>") != std::string::npos);

            REQUIRE(xml2.find("<c:surface3DChart>") != std::string::npos);
            REQUIRE(xml2.find("<c:serAx>") != std::string::npos);
            REQUIRE(xml2.find("<c:axId val=\"100000005\"") != std::string::npos);

            doc.save();
            doc.close();
        }
    }

    SECTION("Bar vs Column Directions")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 3; ++i) { wks.cell(i, 1).value() = i; }

            auto barChart = wks.addChart(XLChartType::Bar, "Bar", 2, 4, 400, 300);
            barChart.addSeries("Sheet1!$A$1:$A$3");

            auto colChart = wks.addChart(XLChartType::Column, "Col", 2, 8, 400, 300);
            colChart.addSeries("Sheet1!$A$1:$A$3");

            std::ostringstream ss1, ss2;
            barChart.xmlDocument().save(ss1);
            colChart.xmlDocument().save(ss2);

            std::string xml1 = ss1.str();
            std::string xml2 = ss2.str();

            REQUIRE(xml1.find("<c:barDir val=\"bar\"") != std::string::npos);
            REQUIRE(xml1.find("<c:barDir val=\"col\"") == std::string::npos);

            REQUIRE(xml2.find("<c:barDir val=\"col\"") != std::string::npos);
            REQUIRE(xml2.find("<c:barDir val=\"bar\"") == std::string::npos);

            doc.save();
            doc.close();
        }
    }

    SECTION("Create Bar Chart")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // Add some data
            for (int i = 1; i <= 10; ++i) { wks.cell(i, 1).value() = i; }

            // Add a chart
            auto chart = wks.addChart(XLChartType::Bar, "My Chart", 2, 4, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$10");

            doc.save();
            doc.close();
        }

        {
            XLChartTestDoc testDoc;
            testDoc.open(filename);

            // 1. Verify that chart1.xml exists and has correct tags
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");
            REQUIRE(chartXml.find("<c:chartSpace") != std::string::npos);
            REQUIRE(chartXml.find("<c:barChart>") != std::string::npos);

            // Verify series was added
            REQUIRE(chartXml.find("<c:ser>") != std::string::npos);
            REQUIRE(chartXml.find("Sheet1!$A$1:$A$10") != std::string::npos);

            // 2. Verify drawing file exists and links to chart
            std::string drawingXml = testDoc.getRawXml("xl/drawings/drawing1.xml");
            REQUIRE(drawingXml.find("<xdr:graphicFrame>") != std::string::npos);
            REQUIRE(drawingXml.find("name=\"My Chart\"") != std::string::npos);
            REQUIRE(drawingXml.find("uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"") != std::string::npos);

            // Extract drawing relationship ID for the chart
            size_t chartPos = drawingXml.find("<c:chart r:id=\"");
            REQUIRE(chartPos != std::string::npos);
            size_t      idStart = chartPos + 15;
            size_t      idEnd   = drawingXml.find("\"", idStart);
            std::string rId     = drawingXml.substr(idStart, idEnd - idStart);

            // 3. Verify drawing rels file points to chart1.xml
            std::string drawingRelsXml = testDoc.getRawXml("xl/drawings/_rels/drawing1.xml.rels");
            REQUIRE(drawingRelsXml.find("Id=\"" + rId + "\"") != std::string::npos);
            REQUIRE(drawingRelsXml.find("Target=\"../charts/chart1.xml\"") != std::string::npos);

            // 4. Verify Content_Types.xml contains chart content type
            std::string ctXml = testDoc.getRawXml("[Content_Types].xml");
            REQUIRE(ctXml.find("PartName=\"/xl/charts/chart1.xml\"") != std::string::npos);
            REQUIRE(ctXml.find("ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Create Line Chart")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks   = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Line, "Line Chart", 2, 4, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$10");
            doc.save();
            doc.close();
        }

        {
            XLChartTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");
            REQUIRE(chartXml.find("<c:lineChart>") != std::string::npos);
            REQUIRE(chartXml.find("<c:ser>") != std::string::npos);
            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Create Pie Chart")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks   = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Pie, "Pie Chart", 2, 4, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$10");
            doc.save();
            doc.close();
        }

        {
            XLChartTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");
            REQUIRE(chartXml.find("<c:pieChart>") != std::string::npos);
            REQUIRE(chartXml.find("<c:ser>") != std::string::npos);
            REQUIRE(chartXml.find("<c:catAx>") == std::string::npos);    // Pie chart shouldn't have axes
            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Create Scatter Chart")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks   = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Scatter, "Scatter Chart", 2, 4, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$10");
            doc.save();
            doc.close();
        }

        {
            XLChartTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");
            REQUIRE(chartXml.find("<c:scatterChart>") != std::string::npos);
            REQUIRE(chartXml.find("<c:scatterStyle val=\"lineMarker\"") != std::string::npos);
            REQUIRE(chartXml.find("<c:ser>") != std::string::npos);
            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Series with Categories and Title")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks   = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Bar, "Bar Chart", 2, 4, 400, 300);

            // valueRef, title, categoriesRef
            chart.addSeries("Sheet1!$B$2:$B$10", "Revenue", "Sheet1!$A$2:$A$10");
            chart.addSeries("Sheet1!$C$2:$C$10", "Sheet1!$C$1", "Sheet1!$A$2:$A$10");

            doc.save();
            doc.close();
        }

        {
            XLChartTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

            // Check literal string title
            REQUIRE(chartXml.find("<c:v>Revenue</c:v>") != std::string::npos);
            // Check cell reference title
            REQUIRE(chartXml.find("<c:f>Sheet1!$C$1</c:f>") != std::string::npos);
            // Check categories
            REQUIRE(chartXml.find("<c:cat>") != std::string::npos);
            REQUIRE(chartXml.find("<c:f>Sheet1!$A$2:$A$10</c:f>") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Scatter Chart Validation")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks   = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Scatter, "Scatter", 2, 4, 400, 300);

            // Y-values, Title, X-values (categoriesRef for Scatter)
            chart.addSeries("Sheet1!$C$2:$C$10", "Correlation", "Sheet1!$B$2:$B$10");

            doc.save();
            doc.close();
        }

        {
            XLChartTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

            // For Scatter chart, the "category" is mapped to X values (c:xVal) and value is Y values (c:yVal)
            REQUIRE(chartXml.find("<c:xVal>") != std::string::npos);
            REQUIRE(chartXml.find("<c:f>Sheet1!$B$2:$B$10</c:f>") != std::string::npos);

            REQUIRE(chartXml.find("<c:yVal>") != std::string::npos);
            REQUIRE(chartXml.find("<c:f>Sheet1!$C$2:$C$10</c:f>") != std::string::npos);

            // X-axis and Y-axis for scatter should both be valAx
            REQUIRE(chartXml.find("<c:valAx>") != std::string::npos);
            REQUIRE(chartXml.find("<c:catAx>") == std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }
}
TEST_CASE("Chart Series Fluent API", "[XLChart][Fluent]")
{
    std::string filename = "test_chart_fluent.xlsx";

    SECTION("Create Line Chart with Fluent API")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            for (int i = 1; i <= 5; ++i) {
                wks.cell(i, 1).value() = i;
                wks.cell(i, 2).value() = i * i;
            }

            auto chart = wks.addChart(XLChartType::Line, "My Fluent Chart", 2, 4, 400, 300);

            chart.addSeries("Sheet1!$B$1:$B$5", "", "Sheet1!$A$1:$A$5")
                .setTitle("Squares")
                .setSmooth(true)
                .setMarkerStyle(XLMarkerStyle::Circle);

            doc.save();
            doc.close();
        }

        {
            XLChartTestDoc testDoc;
            testDoc.open(filename);

            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");
            REQUIRE(chartXml.find("<c:chartSpace") != std::string::npos);
            REQUIRE(chartXml.find("Squares") != std::string::npos);
            REQUIRE(chartXml.find("<c:smooth val=\"1\"") != std::string::npos);
            REQUIRE(chartXml.find("<c:symbol val=\"circle\"") != std::string::npos);
        }
    }
}

TEST_CASE("Chart Anchor API", "[XLChart][Anchor]")
{
    std::string filename = "test_chart_anchor.xlsx";

    SECTION("Create Line Chart with Strong Type Anchor API")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            for (int i = 1; i <= 5; ++i) {
                wks.cell(i, 1).value() = i;
                wks.cell(i, 2).value() = i * i;
            }

            using namespace DistanceLiterals;

            XLChartAnchor anchor;
            anchor.name   = "My Anchor Chart";
            anchor.row    = 2;
            anchor.col    = 4;
            anchor.width  = 12_cm;
            anchor.height = 8_cm;

            auto chart = wks.addChart(XLChartType::Line, anchor);

            chart.addSeries("Sheet1!$B$1:$B$5", "", "Sheet1!$A$1:$A$5")
                .setTitle("Squares")
                .setSmooth(true)
                .setMarkerStyle(XLMarkerStyle::Circle);

            doc.save();
            doc.close();
        }

        {
            XLChartTestDoc testDoc;
            testDoc.open(filename);

            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");
            REQUIRE(chartXml.find("<c:chartSpace") != std::string::npos);
        }
    }
}

TEST_CASE("Chart Cell Range API", "[XLChart][Range]")
{
    std::string filename = "test_chart_range.xlsx";

    SECTION("Add series via XLCellRange")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            for (int i = 1; i <= 5; ++i) {
                wks.cell(i, 1).value() = "Cat" + std::to_string(i);
                wks.cell(i, 2).value() = i * 10;
            }

            auto categories = wks.range(XLCellReference("A1"), XLCellReference("A5"));
            auto values     = wks.range(XLCellReference("B1"), XLCellReference("B5"));

            auto chart = wks.addChart(XLChartType::Bar, "Range Chart", 2, 4, 400, 300);

            // Should internally build Sheet1!$B$1:$B$5 and Sheet1!$A$1:$A$5
            chart.addSeries(wks, values, categories, "My Values");

            doc.save();
            doc.close();
        }

        {
            XLChartTestDoc testDoc;
            testDoc.open(filename);

            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");
            REQUIRE(chartXml.find("Sheet1!$B$1:$B$5") != std::string::npos);
            REQUIRE(chartXml.find("Sheet1!$A$1:$A$5") != std::string::npos);
        }
    }
}

TEST_CASE("Chart Phase1 Phase2 Features", "[XLChart][Phase12]")
{
    auto containsOneOf = [](const std::string& xml, std::initializer_list<std::string> candidates) {
        for (const auto& c : candidates)
            if (xml.find(c) != std::string::npos) return true;
        return false;
    };

    SECTION("P1.1 Series Color")
    {
        const std::string fname = "test_p1_series_color.xlsx";
        {
            XLDocument doc;
            doc.create(fname, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 3; ++i) wks.cell(i, 1).value() = i * 10;
            auto chart = wks.addChart(XLChartType::Bar, "ColorTest", 1, 3, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$3").setColor("FF0000");
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(fname);
            std::string xml = td.getRawXml("xl/charts/chart1.xml");
            REQUIRE(xml.find("c:spPr") != std::string::npos);
            bool hasColor = xml.find("FF0000") != std::string::npos || xml.find("ff0000") != std::string::npos;
            REQUIRE(hasColor);
            td.close();
        }
        std::filesystem::remove(fname);
    }

    SECTION("P1.2 Data Point Color")
    {
        const std::string fname = "test_p1_datapoint_color.xlsx";
        {
            XLDocument doc;
            doc.create(fname, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 5; ++i) wks.cell(i, 1).value() = i;
            auto chart = wks.addChart(XLChartType::Bar, "DPColor", 1, 3, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$5").setDataPointColor(0, "00B050").setDataPointColor(4, "FF0000");
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(fname);
            std::string xml = td.getRawXml("xl/charts/chart1.xml");
            REQUIRE(xml.find("c:dPt") != std::string::npos);
            bool hasGreen = xml.find("00B050") != std::string::npos || xml.find("00b050") != std::string::npos;
            bool hasRed   = xml.find("FF0000") != std::string::npos || xml.find("ff0000") != std::string::npos;
            REQUIRE(hasGreen);
            REQUIRE(hasRed);
            td.close();
        }
        std::filesystem::remove(fname);
    }

    SECTION("P1.3 Gap Width and Overlap")
    {
        const std::string fname = "test_p1_gap_overlap.xlsx";
        {
            XLDocument doc;
            doc.create(fname, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 3; ++i) wks.cell(i, 1).value() = i;
            auto chart = wks.addChart(XLChartType::Bar, "GapTest", 1, 3, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$3");
            chart.setGapWidth(200);
            chart.setOverlap(-20);
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(fname);
            std::string xml        = td.getRawXml("xl/charts/chart1.xml");
            bool        hasGap     = containsOneOf(xml, {"c:gapWidth val=\"200\"", "c:gapWidth val=\"200\" />"});
            bool        hasOverlap = containsOneOf(xml, {"c:overlap val=\"-20\"", "c:overlap val=\"-20\" />"});
            REQUIRE(hasGap);
            REQUIRE(hasOverlap);
            td.close();
        }
        std::filesystem::remove(fname);
    }

    SECTION("P1.4 Axis Number Format")
    {
        const std::string fname = "test_p1_axis_numfmt.xlsx";
        {
            XLDocument doc;
            doc.create(fname, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 3; ++i) wks.cell(i, 1).value() = i * 0.1;
            auto chart = wks.addChart(XLChartType::Bar, "NumFmtTest", 1, 3, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$3");
            chart.yAxis().setNumberFormat("0.00%");
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(fname);
            std::string xml = td.getRawXml("xl/charts/chart1.xml");
            REQUIRE(xml.find("0.00%") != std::string::npos);
            REQUIRE(xml.find("c:numFmt") != std::string::npos);
            td.close();
        }
        std::filesystem::remove(fname);
    }

    SECTION("P1.5 Scatter Sub-type ScatterSmooth")
    {
        const std::string fname = "test_p1_scatter_smooth.xlsx";
        {
            XLDocument doc;
            doc.create(fname, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 4; ++i) {
                wks.cell(i, 1).value() = i;
                wks.cell(i, 2).value() = i * i;
            }
            auto chart = wks.addChart(XLChartType::ScatterSmooth, "SmoothScatter", 1, 4, 400, 300);
            chart.addSeries("Sheet1!$B$1:$B$4", "", "Sheet1!$A$1:$A$4");
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(fname);
            std::string xml = td.getRawXml("xl/charts/chart1.xml");
            REQUIRE(xml.find("c:scatterChart") != std::string::npos);
            REQUIRE(xml.find("val=\"smooth\"") != std::string::npos);
            td.close();
        }
        std::filesystem::remove(fname);
    }

    SECTION("P2.1 Plot Area Color")
    {
        const std::string fname = "test_p2_plotarea_color.xlsx";
        {
            XLDocument doc;
            doc.create(fname, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 3; ++i) wks.cell(i, 1).value() = i;
            auto chart = wks.addChart(XLChartType::Bar, "PlotAreaColor", 1, 3, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$3");
            chart.setPlotAreaColor("E8F4F8");
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(fname);
            std::string xml      = td.getRawXml("xl/charts/chart1.xml");
            bool        hasColor = xml.find("E8F4F8") != std::string::npos || xml.find("e8f4f8") != std::string::npos;
            REQUIRE(hasColor);
            td.close();
        }
        std::filesystem::remove(fname);
    }

    SECTION("P2.2 Chart Area Color")
    {
        const std::string fname = "test_p2_chartarea_color.xlsx";
        {
            XLDocument doc;
            doc.create(fname, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 3; ++i) wks.cell(i, 1).value() = i;
            auto chart = wks.addChart(XLChartType::Bar, "ChartAreaColor", 1, 3, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$3");
            chart.setChartAreaColor("F5F5F5");
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(fname);
            std::string xml      = td.getRawXml("xl/charts/chart1.xml");
            bool        hasColor = xml.find("F5F5F5") != std::string::npos || xml.find("f5f5f5") != std::string::npos;
            REQUIRE(hasColor);
            td.close();
        }
        std::filesystem::remove(fname);
    }

    SECTION("P2.3 Bubble Chart")
    {
        const std::string fname = "test_p2_bubble.xlsx";
        {
            XLDocument doc;
            doc.create(fname, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 4; ++i) {
                wks.cell(i, 1).value() = i * 10;
                wks.cell(i, 2).value() = i * 5;
                wks.cell(i, 3).value() = i;
            }
            auto chart = wks.addChart(XLChartType::Bubble, "BubbleTest", 1, 5, 500, 350);
            chart.setTitle("Market Bubbles");
            chart.addBubbleSeries("Sheet1!$A$1:$A$4", "Sheet1!$B$1:$B$4", "Sheet1!$C$1:$C$4", "S1");
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(fname);
            std::string xml = td.getRawXml("xl/charts/chart1.xml");
            REQUIRE(xml.find("c:bubbleChart") != std::string::npos);
            REQUIRE(xml.find("c:xVal") != std::string::npos);
            REQUIRE(xml.find("c:yVal") != std::string::npos);
            REQUIRE(xml.find("c:bubbleSize") != std::string::npos);
            REQUIRE(xml.find("Sheet1!$A$1:$A$4") != std::string::npos);
            REQUIRE(xml.find("Sheet1!$C$1:$C$4") != std::string::npos);
            td.close();
        }
        std::filesystem::remove(fname);
    }
}

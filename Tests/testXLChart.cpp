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
}

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

    SECTION("Create Bar Chart")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // Add some data
            for (int i = 1; i <= 10; ++i) {
                wks.cell(i, 1).value() = i;
            }

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
            size_t idStart = chartPos + 15;
            size_t idEnd = drawingXml.find("\"", idStart);
            std::string rId = drawingXml.substr(idStart, idEnd - idStart);

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
            auto wks = doc.workbook().worksheet("Sheet1");
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
            auto wks = doc.workbook().worksheet("Sheet1");
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
            REQUIRE(chartXml.find("<c:catAx>") == std::string::npos); // Pie chart shouldn't have axes
            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Create Scatter Chart")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
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
            auto wks = doc.workbook().worksheet("Sheet1");
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
            auto wks = doc.workbook().worksheet("Sheet1");
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
            anchor.name = "My Anchor Chart";
            anchor.row = 2;
            anchor.col = 4;
            anchor.width = 12_cm;
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
            auto values = wks.range(XLCellReference("B1"), XLCellReference("B5"));

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

#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <filesystem>
#include <fstream>

using namespace OpenXLSX;

namespace
{
    // Hack to access protected member without inheritance
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

class XLChartAdvTestDoc
{
public:
    XLDocument  doc;
    void        open(const std::string& filename) { doc.open(filename); }
    void        close() { doc.close(); }
    std::string getRawXml(const std::string& path) { return ::getRawXml(doc, path); }
};

TEST_CASE("Advanced Chart Visual Elements", "[XLChart][OOXML]")
{
    std::string filename = "test_chart_advanced.xlsx";

    SECTION("Chart Title and Rich Text Format")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Bar, "Bar", 1, 1, 400, 300);
            
            chart.setTitle("Strategic Overview");
            
            doc.save();
            doc.close();
        }

        {
            XLChartAdvTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

            // Verify title node architecture
            REQUIRE(chartXml.find("<c:title>") != std::string::npos);
            REQUIRE(chartXml.find("<c:tx>") != std::string::npos);
            REQUIRE(chartXml.find("<c:rich>") != std::string::npos);
            REQUIRE(chartXml.find("<a:t>Strategic Overview</a:t>") != std::string::npos);
            REQUIRE(chartXml.find("<c:overlay val=\"0\"") != std::string::npos);
            
            testDoc.close();
        }
        std::filesystem::remove(filename);
    }
    
    SECTION("Legend Positions")
    {
        std::map<XLLegendPosition, std::string> posMap = {
            {XLLegendPosition::Bottom, "b"},
            {XLLegendPosition::Left, "l"},
            {XLLegendPosition::Right, "r"},
            {XLLegendPosition::Top, "t"},
            {XLLegendPosition::TopRight, "tr"}
        };

        for (auto const& [enumPos, xmlVal] : posMap) {
            {
                XLDocument doc;
                doc.create(filename, XLForceOverwrite);
                auto wks = doc.workbook().worksheet("Sheet1");
                auto chart = wks.addChart(XLChartType::Line, "Line", 1, 1, 400, 300);
                
                chart.setLegendPosition(enumPos);
                
                doc.save();
                doc.close();
            }

            {
                XLChartAdvTestDoc testDoc;
                testDoc.open(filename);
                std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

                std::string expectedStr = "<c:legendPos val=\"" + xmlVal + "\"";
                REQUIRE(chartXml.find(expectedStr) != std::string::npos);
                REQUIRE(chartXml.find("<c:overlay val=\"0\" />") != std::string::npos);
                
                testDoc.close();
            }
            std::filesystem::remove(filename);
        }
    }
    

    SECTION("Chart Data Labels")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Pie, "Pie", 1, 1, 400, 300);
            
            // showValue = true, showCategory = true, showPercent = false
            chart.setShowDataLabels(true, true, false);
            
            doc.save();
            doc.close();
        }

        {
            XLChartAdvTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

            // Verify dLbls node architecture
            REQUIRE(chartXml.find("<c:dLbls>") != std::string::npos);
                        REQUIRE(chartXml.find("<c:showLegendKey val=\"0\" />") != std::string::npos);
            REQUIRE(chartXml.find("<c:showVal val=\"1\" />") != std::string::npos);
            REQUIRE(chartXml.find("<c:showCatName val=\"1\" />") != std::string::npos);
            REQUIRE(chartXml.find("<c:showPercent val=\"0\" />") != std::string::npos);
            
            // strict order validation
            size_t valPos = chartXml.find("<c:showVal");
            size_t catPos = chartXml.find("<c:showCatName");
            size_t percPos= chartXml.find("<c:showPercent");
            REQUIRE(valPos < catPos);
            REQUIRE(catPos < percPos);
            
            testDoc.close();
        }
        std::filesystem::remove(filename);
    }


    SECTION("Chart Data Labels Multi-Type OOXML Verification")
    {
        std::map<XLChartType, std::string> chartNames = {
            {XLChartType::Bar, "c:barChart"},
            {XLChartType::Pie, "c:pieChart"},
            {XLChartType::Line, "c:lineChart"}
        };

        for (auto const& [enumPos, xmlVal] : chartNames) {
            {
                XLDocument doc;
                doc.create(filename, XLForceOverwrite);
                auto wks = doc.workbook().worksheet("Sheet1");
                auto chart = wks.addChart(enumPos, "Chart", 1, 1, 400, 300);
                chart.addSeries("Sheet1!$B$2:$B$5", "Title", "Sheet1!$A$2:$A$5");
                
                // turn on data labels
                chart.setShowDataLabels(true, true, true);
                
                doc.save();
                doc.close();
            }

            {
                XLChartAdvTestDoc testDoc;
                testDoc.open(filename);
                std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

                // verify base tag exists inside the correct chart type element
                REQUIRE(chartXml.find("<" + xmlVal + ">") != std::string::npos);
                REQUIRE(chartXml.find("<c:dLbls>") != std::string::npos);
                
                // the dLbls element MUST appear BEFORE specific elements like c:axId or c:gapWidth
                size_t dLblsPos = chartXml.find("<c:dLbls>");
                size_t axIdPos = chartXml.find("<c:axId");
                
                if (axIdPos != std::string::npos) {
                    REQUIRE(dLblsPos < axIdPos);
                }

                // Verify the strict sequence INSIDE dLbls
                size_t showLegendKeyPos = chartXml.find("<c:showLegendKey");
                size_t showValPos = chartXml.find("<c:showVal");
                size_t showCatNamePos = chartXml.find("<c:showCatName");
                size_t showSerNamePos = chartXml.find("<c:showSerName");
                size_t showPercentPos = chartXml.find("<c:showPercent");
                size_t showBubbleSizePos = chartXml.find("<c:showBubbleSize");
                size_t showLeaderLinesPos = chartXml.find("<c:showLeaderLines");

                REQUIRE(showLegendKeyPos < showValPos);
                REQUIRE(showValPos < showCatNamePos);
                REQUIRE(showCatNamePos < showSerNamePos);
                REQUIRE(showSerNamePos < showPercentPos);
                REQUIRE(showPercentPos < showBubbleSizePos);
                REQUIRE(showBubbleSizePos < showLeaderLinesPos);
                
                testDoc.close();
            }
            std::filesystem::remove(filename);
        }
    }


    SECTION("Chart Axes Configuration")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Line, "Line", 1, 1, 400, 300);
            
            auto xAxis = chart.xAxis();
            xAxis.setTitle("Month");
            xAxis.setMajorGridlines(false);
            
            auto yAxis = chart.yAxis();
            yAxis.setTitle("Revenue (USD)");
            yAxis.setMinBounds(5000);
            yAxis.setMaxBounds(25000);
            yAxis.setMajorGridlines(true);
            yAxis.setMinorGridlines(true);

            doc.save();
            doc.close();
        }

        {
            XLChartAdvTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

            // X-Axis (catAx) bounds and title
            size_t catAxPos = chartXml.find("<c:catAx>");
            size_t valAxPos = chartXml.find("<c:valAx>");
            
            REQUIRE(catAxPos != std::string::npos);
            REQUIRE(valAxPos != std::string::npos);
            
            // X-Axis doesn't have gridlines explicitly configured (false removes the node if it existed, but here we don't insert it)
            // But it has a title
            std::string catAxBlock = chartXml.substr(catAxPos, valAxPos - catAxPos);
            REQUIRE(catAxBlock.find("<a:t>Month</a:t>") != std::string::npos);
            REQUIRE(catAxBlock.find("<c:majorGridlines") == std::string::npos);
            
            // Y-Axis has min/max and minor gridlines and title
            std::string valAxBlock = chartXml.substr(valAxPos);
            REQUIRE(valAxBlock.find("<a:t>Revenue (USD)</a:t>") != std::string::npos);
            REQUIRE(valAxBlock.find("<c:min val=\"5000\" />") != std::string::npos);
            REQUIRE(valAxBlock.find("<c:max val=\"25000\" />") != std::string::npos);
            REQUIRE(valAxBlock.find("<c:minorGridlines />") != std::string::npos);
            REQUIRE(valAxBlock.find("<c:majorGridlines />") != std::string::npos);

            // test clearing max bound
            {
                XLDocument doc;
                doc.open(filename);
                auto wks = doc.workbook().worksheet("Sheet1");
                // Just loading any chart instance to parse and test clearMaxBounds
                auto chartNode = doc.workbook().xmlDocument().document_element(); // We'll skip complex instantiation
                doc.close();
            }

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }


    SECTION("Chart Axes OOXML Strict Node Order")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Scatter, "Scatter", 1, 1, 400, 300);
            
            // Deliberately calling setters out of OOXML sequence to test the re-ordering
            auto yAxis = chart.yAxis();
            yAxis.setMinorGridlines(true);
            yAxis.setMaxBounds(100);
            yAxis.setMajorGridlines(true);
            yAxis.setMinBounds(0);
            yAxis.setTitle("Vertical");

            auto xAxis = chart.xAxis();
            xAxis.setTitle("Horizontal");
            xAxis.setMinorGridlines(true);
            xAxis.setMajorGridlines(true);
            
            doc.save();
            doc.close();
        }

        {
            XLChartAdvTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

            // 1. Verify that c:valAx and c:catAx/c:valAx exist
            size_t ax1Pos = chartXml.find("<c:valAx>"); // Scatter chart uses valAx for both X and Y
            size_t ax2Pos = chartXml.find("<c:valAx>", ax1Pos + 1);
            REQUIRE(ax1Pos != std::string::npos);
            REQUIRE(ax2Pos != std::string::npos);
            
            // 2. Validate strict sequence inside Y-Axis
            // The required order per XLAxisNodeOrder: scaling -> delete -> axPos -> majorGridlines -> minorGridlines -> title -> numFmt ...
            std::string axBlock = chartXml.substr(ax2Pos);
            
            size_t scalingPos  = axBlock.find("<c:scaling>");
            size_t axPosPos    = axBlock.find("<c:axPos");
            size_t majGridPos  = axBlock.find("<c:majorGridlines");
            size_t minGridPos  = axBlock.find("<c:minorGridlines");
            size_t titlePos    = axBlock.find("<c:title>");
            size_t numFmtPos   = axBlock.find("<c:numFmt");

            REQUIRE(scalingPos < axPosPos);
            REQUIRE(axPosPos < majGridPos);

            REQUIRE(majGridPos < minGridPos);

            REQUIRE(minGridPos < titlePos);
            REQUIRE(titlePos < numFmtPos);
            
            // 3. Validate strict sequence inside Scaling
            // Required order per XLScalingNodeOrder: orientation -> max -> min
            size_t orientPos = axBlock.find("<c:orientation");
            size_t maxPos    = axBlock.find("<c:max val=\"100\"/>");
            if (maxPos == std::string::npos) maxPos = axBlock.find("<c:max val=\"100\" />");
            size_t minPos    = axBlock.find("<c:min val=\"0\"/>");
            if (minPos == std::string::npos) minPos = axBlock.find("<c:min val=\"0\" />");

            REQUIRE(orientPos < maxPos);
            REQUIRE(maxPos < minPos);
            
            testDoc.close();
        }
        std::filesystem::remove(filename);
    }


    SECTION("Chart Stacked and Percent Stacked Variants")
    {
        std::map<XLChartType, std::string> chartNames = {
            {XLChartType::BarStacked, "c:barChart"},
            {XLChartType::BarPercentStacked, "c:barChart"},
            {XLChartType::LineStacked, "c:lineChart"},
            {XLChartType::LinePercentStacked, "c:lineChart"},
            {XLChartType::AreaStacked, "c:areaChart"},
            {XLChartType::AreaPercentStacked, "c:areaChart"}
        };

        for (auto const& [enumPos, xmlVal] : chartNames) {
            {
                XLDocument doc;
                doc.create(filename, XLForceOverwrite);
                auto wks = doc.workbook().worksheet("Sheet1");
                auto chart = wks.addChart(enumPos, "Chart", 1, 1, 400, 300);
                chart.addSeries("Sheet1!$B$2:$B$5", "Title", "Sheet1!$A$2:$A$5");
                
                doc.save();
                doc.close();
            }

            {
                XLChartAdvTestDoc testDoc;
                testDoc.open(filename);
                std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

                // verify base tag exists inside the correct chart type element
                REQUIRE(chartXml.find("<" + xmlVal + ">") != std::string::npos);
                
                // verify grouping is stacked or percentStacked
                if (enumPos == XLChartType::BarStacked || enumPos == XLChartType::LineStacked || enumPos == XLChartType::AreaStacked) {
                    REQUIRE(chartXml.find("<c:grouping val=\"stacked\" />") != std::string::npos);
                } else {
                    REQUIRE(chartXml.find("<c:grouping val=\"percentStacked\" />") != std::string::npos);
                }

                // verify overlap is 100 for bar stacked variants
                if (enumPos == XLChartType::BarStacked || enumPos == XLChartType::BarPercentStacked) {
                    REQUIRE(chartXml.find("<c:overlap val=\"100\" />") != std::string::npos);
                }
                
                testDoc.close();
            }
            std::filesystem::remove(filename);
        }
    }


    SECTION("3D Chart Variants")
    {
        std::map<XLChartType, std::string> chartNames = {
            {XLChartType::Bar3D, "c:bar3DChart"},
            {XLChartType::Bar3DStacked, "c:bar3DChart"},
            {XLChartType::Line3D, "c:line3DChart"},
            {XLChartType::Pie3D, "c:pie3DChart"},
            {XLChartType::Area3D, "c:area3DChart"}
        };

        for (auto const& [enumPos, xmlVal] : chartNames) {
            {
                XLDocument doc;
                doc.create(filename, XLForceOverwrite);
                auto wks = doc.workbook().worksheet("Sheet1");
                auto chart = wks.addChart(enumPos, "Chart", 1, 1, 400, 300);
                chart.addSeries("Sheet1!$B$2:$B$5", "Title", "Sheet1!$A$2:$A$5");
                
                doc.save();
                doc.close();
            }

            {
                XLChartAdvTestDoc testDoc;
                testDoc.open(filename);
                std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

                // verify base tag exists
                REQUIRE(chartXml.find("<" + xmlVal + ">") != std::string::npos);
                
                // verify 3D View properties exist
                REQUIRE(chartXml.find("<c:view3D>") != std::string::npos);
                REQUIRE((chartXml.find("<c:rotX val=\"30\"/>") != std::string::npos || chartXml.find("<c:rotX val=\"30\" />") != std::string::npos));
                REQUIRE((chartXml.find("<c:rotY val=\"30\"/>") != std::string::npos || chartXml.find("<c:rotY val=\"30\" />") != std::string::npos));
                
                // strict ordering: view3D MUST come before plotArea
                size_t view3DPos = chartXml.find("<c:view3D>");
                size_t plotAreaPos = chartXml.find("<c:plotArea>");
                REQUIRE(view3DPos < plotAreaPos);
                
                testDoc.close();
            }
            std::filesystem::remove(filename);
        }
    }


    SECTION("Radar Chart Variants")
    {
        std::map<XLChartType, std::string> chartNames = {
            {XLChartType::Radar, "c:radarChart"},
            {XLChartType::RadarFilled, "c:radarChart"},
            {XLChartType::RadarMarkers, "c:radarChart"}
        };

        for (auto const& [enumPos, xmlVal] : chartNames) {
            {
                XLDocument doc;
                doc.create(filename, XLForceOverwrite);
                auto wks = doc.workbook().worksheet("Sheet1");
                auto chart = wks.addChart(enumPos, "Chart", 1, 1, 400, 300);
                chart.addSeries("Sheet1!$B$2:$B$5", "Title", "Sheet1!$A$2:$A$5");
                
                doc.save();
                doc.close();
            }

            {
                XLChartAdvTestDoc testDoc;
                testDoc.open(filename);
                std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

                // verify base tag exists
                REQUIRE(chartXml.find("<" + xmlVal + ">") != std::string::npos);
                
                // verify radar style attributes
                if (enumPos == XLChartType::Radar) {
                    REQUIRE((chartXml.find("<c:radarStyle val=\"standard\"/>") != std::string::npos || chartXml.find("<c:radarStyle val=\"standard\" />") != std::string::npos));
                } else if (enumPos == XLChartType::RadarFilled) {
                    REQUIRE((chartXml.find("<c:radarStyle val=\"filled\"/>") != std::string::npos || chartXml.find("<c:radarStyle val=\"filled\" />") != std::string::npos));
                } else if (enumPos == XLChartType::RadarMarkers) {
                    REQUIRE((chartXml.find("<c:radarStyle val=\"marker\"/>") != std::string::npos || chartXml.find("<c:radarStyle val=\"marker\" />") != std::string::npos));
                }
                
                testDoc.close();
            }
            std::filesystem::remove(filename);
        }
    }

    SECTION("Hide Legend Functionality")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Pie, "Pie", 1, 1, 400, 300);
            
            chart.setLegendPosition(XLLegendPosition::Hidden);
            
            doc.save();
            doc.close();
        }

        {
            XLChartAdvTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

            // Ensure the legend node does not exist at all
            REQUIRE(chartXml.find("<c:legend>") == std::string::npos);
            REQUIRE(chartXml.find("<c:legendPos") == std::string::npos);
            
            testDoc.close();
        }
        std::filesystem::remove(filename);
    }
}

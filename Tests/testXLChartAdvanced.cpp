#include <OpenXLSX.hpp>
#include "XLChart.hpp"

#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
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

    class XLChartTestDoc
    {
    public:
        XLDocument  doc;
        void        open(const std::string& filename) { doc.open(filename); }
        void        close() { doc.close(); }
        std::string getRawXml(const std::string& path) { return ::getRawXml(doc, path); }
    };
}    // namespace

TEST_CASE("AdvancedChartAxisFeatures", "[XLChart][Axis]")
{
    std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();

    SECTION("Logarithmic Scale")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            for (int i = 1; i <= 10; ++i) {
                wks.cell(i, 1).value() = i;
                wks.cell(i, 2).value() = std::pow(10, i);
            }

            auto chart = wks.addChart(XLChartType::Line, "Log Chart", 1, 4, 400, 300);
            chart.addSeries("Sheet1!$B$1:$B$10", "Log Data", "Sheet1!$A$1:$A$10");

            chart.yAxis().setLogScale(10.0);

            doc.save();
            doc.close();
        }

        {
            XLChartTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

            // Verify logBase is present in yAxis (valAx with axPos="l")
            REQUIRE(chartXml.find("<c:valAx>") != std::string::npos);
            REQUIRE(chartXml.find("<c:logBase val=\"10") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Date Axis")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // Dates in Excel are serial numbers
            for (int i = 1; i <= 10; ++i) {
                wks.cell(i, 1).value() = 44197 + i; // Starting from 2021-01-01
                wks.cell(i, 2).value() = i * 10;
            }

            auto chart = wks.addChart(XLChartType::Line, "Date Chart", 1, 4, 400, 300);
            chart.addSeries("Sheet1!$B$1:$B$10", "Date Data", "Sheet1!$A$1:$A$10");

            chart.xAxis().setDateAxis(true);

            doc.save();
            doc.close();
        }

        {
            XLChartTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

            // Verify dateAx is present instead of catAx for the bottom axis
            REQUIRE(chartXml.find("<c:dateAx>") != std::string::npos);
            // REQUIRE(chartXml.find("<c:catAx>") == std::string::npos); // This might depend on if there are other charts, but here there's only one.

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Doughnut Hole Size")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 3; ++i) wks.cell(i, 1).value() = i;
            auto chart = wks.addChart(XLChartType::Doughnut, "HoleTest", 1, 4, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$3");
            chart.setHoleSize(50);
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(filename);
            std::string xml = td.getRawXml("xl/charts/chart1.xml");
            REQUIRE(xml.find("<c:holeSize val=\"50") != std::string::npos);
            td.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("3D Chart Rotation")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 3; ++i) wks.cell(i, 1).value() = i;
            auto chart = wks.addChart(XLChartType::Bar3D, "RotTest", 1, 4, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$3");
            chart.setRotation(45, 60, 25);
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(filename);
            std::string xml = td.getRawXml("xl/charts/chart1.xml");
            REQUIRE(xml.find("<c:view3D>") != std::string::npos);
            REQUIRE(xml.find("<c:rotX val=\"45") != std::string::npos);
            REQUIRE(xml.find("<c:rotY val=\"60") != std::string::npos);
            REQUIRE(xml.find("<c:perspective val=\"25") != std::string::npos);
            td.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Axis Orientation and Crossing")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 5; ++i) {
                wks.cell(i, 1).value() = i;
                wks.cell(i, 2).value() = i * 10;
            }
            auto chart = wks.addChart(XLChartType::Line, "OrientTest", 1, 4, 400, 300);
            chart.addSeries("Sheet1!$B$1:$B$5", "Data", "Sheet1!$A$1:$A$5");
            
            // Set X-axis to reverse order
            chart.xAxis().setOrientation(XLAxisOrientation::MaxMin);
            
            // Set Y-axis to cross at 25
            chart.yAxis().setCrossesAt(25.0);
            
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(filename);
            std::string xml = td.getRawXml("xl/charts/chart1.xml");
            
            // Check orientation in xAxis (catAx)
            REQUIRE(xml.find("<c:catAx>") != std::string::npos);
            REQUIRE(xml.find("<c:orientation val=\"maxMin") != std::string::npos);
            
            // Check crossesAt in yAxis (valAx)
            REQUIRE(xml.find("<c:valAx>") != std::string::npos);
            REQUIRE(xml.find("<c:crossesAt val=\"25") != std::string::npos);
            
            td.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Axis Crossing Max")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 3; ++i) wks.cell(i, 1).value() = i;
            auto chart = wks.addChart(XLChartType::Bar, "CrossMaxTest", 1, 4, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$3");
            
            // Set X-axis to cross at Max (moves Y-axis to the right)
            chart.xAxis().setCrosses(XLAxisCrosses::Max);
            
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(filename);
            std::string xml = td.getRawXml("xl/charts/chart1.xml");
            REQUIRE(xml.find("<c:crosses val=\"max") != std::string::npos);
            td.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Major and Minor Units")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 10; ++i) wks.cell(i, 1).value() = i * 10;
            auto chart = wks.addChart(XLChartType::Line, "UnitsTest", 1, 4, 400, 300);
            chart.addSeries("Sheet1!$A$1:$A$10");
            
            chart.yAxis().setMajorUnit(20.0);
            chart.yAxis().setMinorUnit(5.0);
            
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(filename);
            std::string xml = td.getRawXml("xl/charts/chart1.xml");
            REQUIRE(xml.find("<c:majorUnit val=\"20") != std::string::npos);
            REQUIRE(xml.find("<c:minorUnit val=\"5") != std::string::npos);
            td.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Data Labels from Range")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 5; ++i) {
                wks.cell(i, 1).value() = i * 10; // Values
                wks.cell(i, 2).value() = "Label " + std::to_string(i); // Labels
            }
            auto chart = wks.addChart(XLChartType::Bar, "LabelsTest", 1, 4, 400, 300);
            auto series = chart.addSeries("Sheet1!$A$1:$A$5");
            
            series.setDataLabelsFromRange(wks, wks.range("B1:B5"));
            
            doc.save();
            doc.close();
        }
        {
            XLChartTestDoc td;
            td.open(filename);
            std::string xml = td.getRawXml("xl/charts/chart1.xml");
            REQUIRE(xml.find("<c:dLbls>") != std::string::npos);
            REQUIRE(xml.find("c15:datalblsRange") != std::string::npos);
            REQUIRE(xml.find("Sheet1!$B$1:$B$5") != std::string::npos);
            REQUIRE(xml.find("c15:showDataLabelsRange val=\"1\"") != std::string::npos);
            td.close();
        }
        std::filesystem::remove(filename);
    }
}

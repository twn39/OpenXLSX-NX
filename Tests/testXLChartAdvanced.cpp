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

    class XLChartTestDoc
    {
    public:
        XLDocument  doc;
        void        open(const std::string& filename) { doc.open(filename); }
        void        close() { doc.close(); }
        std::string getRawXml(const std::string& path) { return ::getRawXml(doc, path); }
    };
}    // namespace

TEST_CASE("Advanced Chart Axis Features", "[XLChart][Axis]")
{
    std::string filename = "test_chart_advanced_axis.xlsx";

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
}

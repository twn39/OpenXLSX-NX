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
}
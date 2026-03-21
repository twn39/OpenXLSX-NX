#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <filesystem>
#include <fstream>

using namespace OpenXLSX;

namespace
{
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

class XLChartsheetTestDoc
{
public:
    XLDocument  doc;
    void        open(const std::string& filename) { doc.open(filename); }
    void        close() { doc.close(); }
    std::string getRawXml(const std::string& path) { return ::getRawXml(doc, path); }
};

TEST_CASE("Chartsheet Creation and Verification", "[XLChartsheet][OOXML]")
{
    std::string filename = "test_chartsheet.xlsx";

    SECTION("Create Chartsheet")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            
            // Add a normal sheet and put data there
            auto wks = doc.workbook().worksheet("Sheet1");
            for (int i = 1; i <= 5; ++i) {
                wks.cell(i, 1).value() = i;
            }

            // Create a chartsheet
            doc.workbook().addChartsheet("My ChartSheet");
            auto cs = doc.workbook().chartsheet("My ChartSheet");
            
            // Add chart to chartsheet
            auto chart = cs.addChart(XLChartType::Bar, "Chart1");
            chart.addSeries("Sheet1!$A$1:$A$5");
            chart.setTitle("Chartsheet Title");

            doc.save();
            doc.close();
        }

        {
            XLChartsheetTestDoc testDoc;
            testDoc.open(filename);

            // 1. Check workbook.xml for chartsheet entry
            std::string wbkXml = testDoc.getRawXml("xl/workbook.xml");
            REQUIRE(wbkXml.find("name=\"My ChartSheet\"") != std::string::npos);

            // 2. Extract rId to find sheet path
            size_t sheetPos = wbkXml.find("name=\"My ChartSheet\"");
            size_t rIdPos = wbkXml.find("r:id=\"", sheetPos);
            REQUIRE(rIdPos != std::string::npos);
            size_t idStart = rIdPos + 6;
            size_t idEnd = wbkXml.find("\"", idStart);
            std::string rId = wbkXml.substr(idStart, idEnd - idStart);

            // 3. Find target in workbook.xml.rels
            std::string wbkRels = testDoc.getRawXml("xl/_rels/workbook.xml.rels");
            REQUIRE(wbkRels.find("Id=\"" + rId + "\"") != std::string::npos);
            REQUIRE(wbkRels.find("chartsheets/sheet") != std::string::npos);

            // Assuming it's chartsheets/sheet2.xml based on normal incremental ID
            std::string csXml = testDoc.getRawXml("xl/chartsheets/sheet2.xml");
            REQUIRE(csXml.find("<chartsheet") != std::string::npos);
            REQUIRE(csXml.find("<drawing r:id=") != std::string::npos);

            // 4. Verify Content_Types.xml contains chartsheet content type
            std::string ctXml = testDoc.getRawXml("[Content_Types].xml");
            REQUIRE(ctXml.find("application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml") != std::string::npos);

            // 5. Verify the relationship file exists in the CORRECT path (xl/chartsheets/_rels)
            std::string csRelsXml = testDoc.getRawXml("xl/chartsheets/_rels/sheet2.xml.rels");
            REQUIRE(csRelsXml.find("<Relationships") != std::string::npos);
            REQUIRE(csRelsXml.find("Target=\"../drawings/drawing1.xml\"") != std::string::npos);

            // 6. Verify the drawing has an absoluteAnchor
            std::string drawingXml = testDoc.getRawXml("xl/drawings/drawing1.xml");
            REQUIRE(drawingXml.find("<xdr:absoluteAnchor>") != std::string::npos);
            
            // 7. Verify chart exists
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");
            REQUIRE(chartXml.find("<c:barChart>") != std::string::npos);
            REQUIRE(chartXml.find("Chartsheet Title") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }
}

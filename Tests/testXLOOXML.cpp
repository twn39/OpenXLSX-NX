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
        typedef std::string (XLDocument::*type)(const std::string&);
    };

    template struct Rob<XLDocument_extractXmlFromArchive, &XLDocument::extractXmlFromArchive>;

    // Prototype declaration for the friend function
    std::string (XLDocument::* get_impl(XLDocument_extractXmlFromArchive))(const std::string&);

    // Function to call the protected member
    std::string getRawXml(XLDocument& doc, const std::string& path)
    {
        static auto fn = get_impl(XLDocument_extractXmlFromArchive());
        return (doc.*fn)(path);
    }
}    // namespace

// Wrapper class for easier testing
class XLDocumentTest
{
public:
    XLDocument  doc;
    void        open(const std::string& filename) { doc.open(filename); }
    void        close() { doc.close(); }
    std::string getRawXml(const std::string& path) { return ::getRawXml(doc, path); }
    std::string property(XLProperty prop) { return doc.property(prop); }
};

TEST_CASE("OOXML Verification Tests", "[OOXML]")
{
    SECTION("Verify Document Properties in core.xml")
    {
        std::string filename = "ooxml_test.xlsx";
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            doc.setProperty(XLProperty::Title, "Test Title");
            doc.setProperty(XLProperty::Subject, "Test Subject");
            doc.setProperty(XLProperty::Creator, "OpenXLSX Tester");
            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
testTestDoc:    // Wait, I'll just use the doc directly.
            testDoc.open(filename);

            // Verify properties via API
            REQUIRE(testDoc.property(XLProperty::Title) == "Test Title");
            REQUIRE(testDoc.property(XLProperty::Subject) == "Test Subject");
            REQUIRE(testDoc.property(XLProperty::Creator) == "OpenXLSX Tester");

            // Verify raw XML content in core.xml
            std::string coreXml = testDoc.getRawXml("docProps/core.xml");

            // 1. Check for XML declaration (preserved by format_default)
            REQUIRE(coreXml.find("<?xml") != std::string::npos);
            REQUIRE(coreXml.find("version=\"1.0\"") != std::string::npos);
            REQUIRE(coreXml.find("encoding=\"UTF-8\"") != std::string::npos);

            // 2. Check for Title with correct namespace
            REQUIRE(coreXml.find("<dc:title>Test Title</dc:title>") != std::string::npos);

            // 3. Check for Subject with correct namespace
            REQUIRE(coreXml.find("<dc:subject>Test Subject</dc:subject>") != std::string::npos);

            // 4. Check for Creator with correct namespace
            REQUIRE(coreXml.find("<dc:creator>OpenXLSX Tester</dc:creator>") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify Data Validation in sheet1.xml")
    {
        std::string filename = "ooxml_dv_test.xlsx";
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // 1. Add a list validation
            auto dv1 = wks.dataValidations().append();
            dv1.setSqref("A1");
            dv1.setList({"Alpha", "Beta", "Gamma"});
            dv1.setPrompt("Title1", "Message1");

            // 2. Add a numeric range validation
            auto dv2 = wks.dataValidations().append();
            dv2.setSqref("B1:B10");
            dv2.setType(XLDataValidationType::Whole);
            dv2.setOperator(XLDataValidationOperator::Between);
            dv2.setFormula1("1");
            dv2.setFormula2("100");
            dv2.setError("ErrorTitle", "Value must be between 1 and 100");

            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            // Verify raw XML content in sheet1.xml
            std::string sheetXml = testDoc.getRawXml("xl/worksheets/sheet1.xml");

            // 1. Check for dataValidations container and count
            REQUIRE(sheetXml.find("<dataValidations count=\"2\">") != std::string::npos);

            // 2. Check for the list validation
            REQUIRE(sheetXml.find("type=\"list\"") != std::string::npos);
            REQUIRE(sheetXml.find("sqref=\"A1\"") != std::string::npos);
            REQUIRE(sheetXml.find("<formula1>\"Alpha,Beta,Gamma\"</formula1>") != std::string::npos);
            REQUIRE(sheetXml.find("promptTitle=\"Title1\"") != std::string::npos);
            REQUIRE(sheetXml.find("prompt=\"Message1\"") != std::string::npos);

            // 3. Check for the numeric validation
            REQUIRE(sheetXml.find("type=\"whole\"") != std::string::npos);
            REQUIRE(sheetXml.find("operator=\"between\"") != std::string::npos);
            REQUIRE(sheetXml.find("sqref=\"B1:B10\"") != std::string::npos);
            REQUIRE(sheetXml.find("<formula1>1</formula1>") != std::string::npos);
            REQUIRE(sheetXml.find("<formula2>100</formula2>") != std::string::npos);
            REQUIRE(sheetXml.find("errorTitle=\"ErrorTitle\"") != std::string::npos);
            REQUIRE(sheetXml.find("error=\"Value must be between 1 and 100\"") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify Hyperlinks in worksheet and relationships")
    {
        std::string filename = "ooxml_hyperlink_test.xlsx";
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // 1. Add an external hyperlink
            wks.cell("A1").value() = "External Link";
            wks.addHyperlink("A1", "https://www.github.com", "GitHub");

            // 2. Add an internal hyperlink
            wks.cell("B2").value() = "Internal Link";
            wks.addInternalHyperlink("B2", "Sheet1!A10", "Scroll Down");

            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            // 1. Verify sheet1.xml
            std::string sheetXml = testDoc.getRawXml("xl/worksheets/sheet1.xml");

            // Check hyperlinks container
            REQUIRE(sheetXml.find("<hyperlinks>") != std::string::npos);

            // Check external link entry (should have r:id and tooltip, but no location)
            REQUIRE(sheetXml.find("ref=\"A1\"") != std::string::npos);
            REQUIRE(sheetXml.find("tooltip=\"GitHub\"") != std::string::npos);
            REQUIRE(sheetXml.find("r:id=\"") != std::string::npos);

            // Check internal link entry (should have location and tooltip)
            REQUIRE(sheetXml.find("ref=\"B2\"") != std::string::npos);
            REQUIRE(sheetXml.find("location=\"Sheet1!A10\"") != std::string::npos);
            REQUIRE(sheetXml.find("tooltip=\"Scroll Down\"") != std::string::npos);

            // 2. Verify sheet1.xml.rels for the external link
            std::string relsXml = testDoc.getRawXml("xl/worksheets/_rels/sheet1.xml.rels");

            // Find the relationship ID used in A1
            size_t      a1Pos   = sheetXml.find("ref=\"A1\"");
            size_t      idStart = sheetXml.find("r:id=\"", a1Pos) + 6;
            size_t      idEnd   = sheetXml.find("\"", idStart);
            std::string rId     = sheetXml.substr(idStart, idEnd - idStart);

            // Verify this ID exists in rels and points to GitHub
            REQUIRE(relsXml.find("Id=\"" + rId + "\"") != std::string::npos);
            REQUIRE(relsXml.find("Target=\"https://www.github.com\"") != std::string::npos);
            REQUIRE(relsXml.find("TargetMode=\"External\"") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify Image insertion OOXML structure")
    {
        std::string filename  = "ooxml_image_test.xlsx";
        std::string imagePath = "./Tests/test.png";

        // Skip if test image doesn't exist
        if (!std::filesystem::exists(imagePath)) return;

        std::string imgData;
        {
            std::ifstream file(imagePath, std::ios::binary);
            imgData = std::string((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
        }

        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            wks.addImage("inserted_test.png", imgData, 5, 5, 100, 100);
            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            // 1. Verify sheet1.xml has a drawing reference
            std::string sheetXml   = testDoc.getRawXml("xl/worksheets/sheet1.xml");
            size_t      drawingPos = sheetXml.find("<drawing r:id=\"");
            REQUIRE(drawingPos != std::string::npos);

            size_t      idStart = drawingPos + 15;
            size_t      idEnd   = sheetXml.find("\"", idStart);
            std::string rId     = sheetXml.substr(idStart, idEnd - idStart);

            // 2. Verify sheet1.xml.rels points to drawing1.xml
            std::string sheetRelsXml = testDoc.getRawXml("xl/worksheets/_rels/sheet1.xml.rels");
            REQUIRE(sheetRelsXml.find("Id=\"" + rId + "\"") != std::string::npos);
            REQUIRE(sheetRelsXml.find("Target=\"../drawings/drawing1.xml\"") != std::string::npos);

            // 3. Verify drawing1.xml exists and contains a picture element
            std::string drawingXml = testDoc.getRawXml("xl/drawings/drawing1.xml");
            REQUIRE(drawingXml.find("<xdr:pic>") != std::string::npos);
            REQUIRE(drawingXml.find("<xdr:cNvPr id=\"") != std::string::npos);
            REQUIRE(drawingXml.find("name=\"inserted_test.png\"") != std::string::npos);

            // Get image relationship ID from drawing1.xml
            size_t picPos  = drawingXml.find("<xdr:blipFill>");
            size_t blipPos = drawingXml.find("r:embed=\"", picPos);
            REQUIRE(blipPos != std::string::npos);
            size_t      imgIdStart = blipPos + 9;
            size_t      imgIdEnd   = drawingXml.find("\"", imgIdStart);
            std::string imgRId     = drawingXml.substr(imgIdStart, imgIdEnd - imgIdStart);

            // 4. Verify drawing1.xml.rels points to the media file
            std::string drawingRelsXml = testDoc.getRawXml("xl/drawings/_rels/drawing1.xml.rels");
            REQUIRE(drawingRelsXml.find("Id=\"" + imgRId + "\"") != std::string::npos);
            REQUIRE(drawingRelsXml.find("Target=\"../media/inserted_test.png\"") != std::string::npos);

            // 5. Verify the media file exists and has correct data
            std::string retrievedImgData = testDoc.getRawXml("xl/media/inserted_test.png");
            REQUIRE(retrievedImgData == imgData);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify XML Declaration persistence in other files")
    {
        std::string filename = "ooxml_decl_test.xlsx";
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            // Check workbook.xml
            std::string workbookXml = testDoc.getRawXml("xl/workbook.xml");
            REQUIRE(workbookXml.find("<?xml") != std::string::npos);

            // Check sheet1.xml
            std::string sheetXml = testDoc.getRawXml("xl/worksheets/sheet1.xml");
            REQUIRE(sheetXml.find("<?xml") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify Default Styles Generation in styles.xml")
    {
        std::string filename = "ooxml_default_styles_test.xlsx";
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto styles = doc.styles();
            
            // Create default styles without copying from another template
            styles.fonts().create();
            styles.fills().create();
            styles.borders().create();
            
            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            // Verify raw XML content in styles.xml
            std::string stylesXml = testDoc.getRawXml("xl/styles.xml");

            // 1. Check for Default Font
            REQUIRE(stylesXml.find("<font>") != std::string::npos);
            REQUIRE(stylesXml.find("name val=\"Arial\"") != std::string::npos);
            REQUIRE(stylesXml.find("sz val=\"12\"") != std::string::npos);
            REQUIRE(stylesXml.find("color rgb=\"ff000000\"") != std::string::npos);
            REQUIRE(stylesXml.find("family val=\"0\"") != std::string::npos);
            REQUIRE(stylesXml.find("charset val=\"1\"") != std::string::npos);

            // 2. Check for Default Fill (Pattern none)
            REQUIRE(stylesXml.find("<fill>") != std::string::npos);
            REQUIRE(stylesXml.find("patternFill patternType=\"none\"") != std::string::npos);

            // 3. Check for Default Borders (All sides none, with black color)
            REQUIRE(stylesXml.find("<border>") != std::string::npos);
            REQUIRE(stylesXml.find("<left style=\"\">") != std::string::npos);
            REQUIRE(stylesXml.find("<right style=\"\">") != std::string::npos);
            REQUIRE(stylesXml.find("<top style=\"\">") != std::string::npos);
            REQUIRE(stylesXml.find("<bottom style=\"\">") != std::string::npos);
            REQUIRE(stylesXml.find("<diagonal style=\"\">") != std::string::npos);
            // Verify black colors on border nodes
            REQUIRE(stylesXml.find("<color rgb=\"ff000000\"") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }
}

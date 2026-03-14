#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <filesystem>

using namespace OpenXLSX;

// Hack to access protected member without inheritance (since XLDocument is final)
template <typename Tag, typename Tag::type M>
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
std::string (XLDocument::*get_impl(XLDocument_extractXmlFromArchive)) (const std::string&);

// Function to call the protected member
std::string getRawXml(XLDocument& doc, const std::string& path)
{
    static auto fn = get_impl(XLDocument_extractXmlFromArchive());
    return (doc.*fn)(path);
}

// Wrapper class for easier testing
class XLDocumentTest
{
public:
    XLDocument doc;
    void open(const std::string& filename) { doc.open(filename); }
    void close() { doc.close(); }
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
            testTestDoc: // Wait, I'll just use the doc directly.
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
            size_t a1Pos = sheetXml.find("ref=\"A1\"");
            size_t idStart = sheetXml.find("r:id=\"", a1Pos) + 6;
            size_t idEnd = sheetXml.find("\"", idStart);
            std::string rId = sheetXml.substr(idStart, idEnd - idStart);
            
            // Verify this ID exists in rels and points to GitHub
            REQUIRE(relsXml.find("Id=\"" + rId + "\"") != std::string::npos);
            REQUIRE(relsXml.find("Target=\"https://www.github.com\"") != std::string::npos);
            REQUIRE(relsXml.find("TargetMode=\"External\"") != std::string::npos);

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
}

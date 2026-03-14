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
            XLDocument doc;
            doc.open(filename);

            // Verify properties via API
            REQUIRE(doc.property(XLProperty::Title) == "Test Title");
            REQUIRE(doc.property(XLProperty::Subject) == "Test Subject");
            REQUIRE(doc.property(XLProperty::Creator) == "OpenXLSX Tester");

            // Verify raw XML content in core.xml
            std::string coreXml = getRawXml(doc, "docProps/core.xml");

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

            doc.close();
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
            XLDocument doc;
            doc.open(filename);

            // Check workbook.xml
            std::string workbookXml = getRawXml(doc, "xl/workbook.xml");
            REQUIRE(workbookXml.find("<?xml") != std::string::npos);

            // Check sheet1.xml
            std::string sheetXml = getRawXml(doc, "xl/worksheets/sheet1.xml");
            REQUIRE(sheetXml.find("<?xml") != std::string::npos);

            doc.close();
        }
        std::filesystem::remove(filename);
    }
}

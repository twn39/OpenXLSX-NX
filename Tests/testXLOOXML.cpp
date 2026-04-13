#include <OpenXLSX.hpp>
#include "XLTables.hpp"

#include "XLDataValidation.hpp"

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

    // Prototype declaration for the friend function
    std::string (XLDocument::* get_impl(XLDocument_extractXmlFromArchive))(std::string_view);

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

TEST_CASE("OOXMLVerificationTests", "[OOXML]")
{
    SECTION("Verify Document Properties in core.xml")
    {
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
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
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
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
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
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
        std::string filename  = OpenXLSX::TestHelpers::getUniqueFilename();
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
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
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
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
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
            REQUIRE(stylesXml.find("color rgb=\"FF000000\"") != std::string::npos);
            REQUIRE(stylesXml.find("family val=\"0\"") != std::string::npos);
            REQUIRE(stylesXml.find("charset val=\"1\"") != std::string::npos);

            // 2. Check for Default Fill (Pattern none)
            REQUIRE(stylesXml.find("<fill>") != std::string::npos);
            REQUIRE(stylesXml.find("patternFill patternType=\"none\"") != std::string::npos);

            // 3. Check for Default Borders (All sides none, with black color)
            REQUIRE(stylesXml.find("<border>") != std::string::npos);
            REQUIRE(stylesXml.find("<left>") != std::string::npos);
            REQUIRE(stylesXml.find("<right>") != std::string::npos);
            REQUIRE(stylesXml.find("<top>") != std::string::npos);
            REQUIRE(stylesXml.find("<bottom>") != std::string::npos);
            REQUIRE(stylesXml.find("<diagonal>") != std::string::npos);
            // Verify black colors on border nodes
            REQUIRE(stylesXml.find("<color rgb=\"FF000000\"") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify Table Structure and Global Uniqueness in OOXML")
    {
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);

            // Table 1 on Sheet 1
            auto wks1               = doc.workbook().worksheet("Sheet1");
            wks1.cell("A1").value() = "H1";
            auto table1             = wks1.tables().add("TableAlpha", "A1:A2");
            table1.appendColumn("H1");

            // Table 2 on Sheet 2
            doc.workbook().addWorksheet("Sheet2");
            auto wks2               = doc.workbook().worksheet("Sheet2");
            wks2.cell("A1").value() = "H2";
            auto table2             = wks2.tables().add("TableBeta", "A1:A2");
            table2.appendColumn("H2");
            table2.setShowFirstColumn(true);
            table2.setShowRowStripes(false);

            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            // 1. Verify [Content_Types].xml doesn't contain .rels overrides
            std::string ctXml = testDoc.getRawXml("[Content_Types].xml");
            REQUIRE(ctXml.find("PartName=\"/xl/worksheets/_rels/sheet1.xml.rels\"") == std::string::npos);
            REQUIRE(ctXml.find("ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"") !=
                    std::string::npos);    // Should be in Default

            // 2. Verify Table 1 (ID 1)
            std::string table1Xml = testDoc.getRawXml("xl/tables/table1.xml");
            REQUIRE(table1Xml.find("id=\"1\"") != std::string::npos);
            REQUIRE(table1Xml.find("name=\"TableAlpha\"") != std::string::npos);
            REQUIRE(table1Xml.find("displayName=\"TableAlpha\"") != std::string::npos);

            // 3. Verify Table 2 (ID 2 - Global Uniqueness)
            std::string table2Xml = testDoc.getRawXml("xl/tables/table2.xml");
            REQUIRE(table2Xml.find("id=\"2\"") != std::string::npos);
            REQUIRE(table2Xml.find("name=\"TableBeta\"") != std::string::npos);
            REQUIRE(table2Xml.find("displayName=\"TableBeta\"") != std::string::npos);

            // 4. Verify tableStyleInfo attribute order and boolean format
            // Order: showFirstColumn -> showLastColumn -> showRowStripes -> showColumnStripes
            size_t posFirst = table2Xml.find("showFirstColumn=\"1\"");
            size_t posLast  = table2Xml.find("showLastColumn=\"0\"");
            size_t posRow   = table2Xml.find("showRowStripes=\"0\"");
            size_t posCol   = table2Xml.find("showColumnStripes=\"0\"");

            REQUIRE(posFirst != std::string::npos);
            REQUIRE(posLast != std::string::npos);
            REQUIRE(posRow != std::string::npos);
            REQUIRE(posCol != std::string::npos);

            REQUIRE(posFirst < posLast);
            REQUIRE(posLast < posRow);
            REQUIRE(posRow < posCol);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify Table TotalsRow and AutoFilter Structure in OOXML")
    {
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);

            auto wks   = doc.workbook().worksheet("Sheet1");
            auto table = wks.tables().add("FeatureTable", "A1:B3");

            // AutoFilter
            auto filter = table.autoFilter();
            filter.filterColumn(1).setCustomFilter("greaterThan", "100");

            // Totals Row
            table.setShowTotalsRow(true);

            auto col1 = table.column(1);
            col1.setName("H1");
            col1.setTotalsRowLabel("Total");

            auto col2 = table.column(2);
            col2.setName("H2");
            col2.setTotalsRowFunction(XLTotalsRowFunction::Sum);

            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            std::string tableXml = testDoc.getRawXml("xl/tables/table1.xml");

            // 1. Verify totals row attributes on root node
            REQUIRE(tableXml.find("totalsRowShown=\"1\"") != std::string::npos);
            REQUIRE(tableXml.find("totalsRowCount=\"1\"") != std::string::npos);

            // 2. Verify autoFilter and custom filter
            REQUIRE(tableXml.find("<autoFilter ref=\"A1:B3\">") != std::string::npos);
            REQUIRE(tableXml.find("<filterColumn colId=\"1\">") != std::string::npos);
            REQUIRE(tableXml.find("<customFilter operator=\"greaterThan\" val=\"100\"") != std::string::npos);

            // 3. Verify table column attributes
            REQUIRE(tableXml.find("<tableColumn id=\"1\" name=\"H1\" totalsRowLabel=\"Total\"") != std::string::npos);
            REQUIRE(tableXml.find("<tableColumn id=\"2\" name=\"H2\" totalsRowFunction=\"sum\"") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify Properties Structure in OOXML")
    {
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);

            auto& core = doc.coreProperties();
            core.setTitle("OOXML Title");
            core.setCreator("OOXML Creator");
            core.setCreated(XLDateTime::fromString("2020-01-01 12:00:00"));

            auto& app = doc.appProperties();
            app.setProperty("Company", "OOXML Company");
            app.setProperty("Application", "OpenXLSX");
            app.setProperty("AppVersion", "1.0");

            auto& custom = doc.customProperties();
            custom.setProperty("StringProp", "Value");
            custom.setProperty("IntProp", 123);
            // 2024-03-12 12:00:00
            custom.setProperty("DateProp", XLDateTime(45363.5));

            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            // 1. Verify docProps/core.xml
            std::string coreXml = testDoc.getRawXml("docProps/core.xml");
            REQUIRE(coreXml.find("xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"") !=
                    std::string::npos);
            REQUIRE(coreXml.find("<dc:title>OOXML Title</dc:title>") != std::string::npos);
            REQUIRE(coreXml.find("<dc:creator>OOXML Creator</dc:creator>") != std::string::npos);
            REQUIRE(coreXml.find("<dcterms:created xsi:type=\"dcterms:W3CDTF\">2020-01-01T12:00:00Z</dcterms:created>") !=
                    std::string::npos);

            // 2. Verify docProps/app.xml
            std::string appXml = testDoc.getRawXml("docProps/app.xml");
            REQUIRE(appXml.find("xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"") !=
                    std::string::npos);
            REQUIRE(appXml.find("<Company>OOXML Company</Company>") != std::string::npos);
            REQUIRE(appXml.find("<Application>") != std::string::npos);
            REQUIRE(appXml.find("OpenXLSX") != std::string::npos);
            REQUIRE(appXml.find("</Application>") != std::string::npos);

            // 3. Verify docProps/custom.xml
            std::string customXml = testDoc.getRawXml("docProps/custom.xml");
            REQUIRE(customXml.find("xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\"") !=
                    std::string::npos);

            // String property
            REQUIRE(customXml.find("name=\"StringProp\"") != std::string::npos);
            REQUIRE(customXml.find("<vt:lpwstr>Value</vt:lpwstr>") != std::string::npos);

            // Int property
            REQUIRE(customXml.find("name=\"IntProp\"") != std::string::npos);
            REQUIRE(customXml.find("<vt:i4>123</vt:i4>") != std::string::npos);

            // Date property (vt:filetime) - 2024-03-12 12:00:00
            REQUIRE(customXml.find("name=\"DateProp\"") != std::string::npos);
            REQUIRE(customXml.find("<vt:filetime>2024-03-12T12:00:00Z</vt:filetime>") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify Workbook Protection and Node Ordering in workbook.xml")
    {
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wb = doc.workbook();

            // 1. Set workbook protection
            wb.protect(true, true, "OpenXLSX");

            // 2. Set some other properties to ensure order is maintained
            wb.setFullCalculationOnLoad();

            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            // Verify raw XML content in workbook.xml
            std::string workbookXml = testDoc.getRawXml("xl/workbook.xml");

            // 1. Check for workbookProtection element and attributes
            REQUIRE(workbookXml.find("<workbookProtection") != std::string::npos);
            REQUIRE(workbookXml.find("lockStructure=\"true\"") != std::string::npos);
            REQUIRE(workbookXml.find("lockWindows=\"true\"") != std::string::npos);
            REQUIRE(workbookXml.find("workbookPassword=\"a355\"") != std::string::npos);

            // 2. Check for calcPr (set by setFullCalculationOnLoad)
            REQUIRE(workbookXml.find("<calcPr") != std::string::npos);
            REQUIRE(workbookXml.find("fullCalcOnLoad=\"true\"") != std::string::npos);

            // 3. Verify Node Ordering: workbookPr -> workbookProtection -> bookViews -> sheets -> calcPr
            size_t posWorkbookPr      = workbookXml.find("<workbookPr");
            size_t posWorkbookProtect = workbookXml.find("<workbookProtection");
            size_t posBookViews       = workbookXml.find("<bookViews");
            size_t posSheets          = workbookXml.find("<sheets");
            size_t posCalcPr          = workbookXml.find("<calcPr");

            REQUIRE(posWorkbookPr != std::string::npos);
            REQUIRE(posWorkbookProtect != std::string::npos);
            REQUIRE(posBookViews != std::string::npos);
            REQUIRE(posSheets != std::string::npos);
            REQUIRE(posCalcPr != std::string::npos);

            REQUIRE(posWorkbookPr < posWorkbookProtect);
            REQUIRE(posWorkbookProtect < posBookViews);
            REQUIRE(posBookViews < posSheets);
            REQUIRE(posSheets < posCalcPr);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify setSheetIndex updates definedNames localSheetId in OOXML")
    {
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            doc.workbook().addWorksheet("DataSheet");     // Index 2, localSheetId 1
            doc.workbook().addWorksheet("AuditSheet");    // Index 3, localSheetId 2

            auto dn = doc.workbook().definedNames();
            dn.append("LocalName1", "DataSheet!$A$1", 1);     // Scoped to DataSheet
            dn.append("LocalName2", "AuditSheet!$B$2", 2);    // Scoped to AuditSheet

            // Move AuditSheet to index 1
            doc.workbook().setSheetIndex("AuditSheet", 1);

            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            // Verify workbook.xml for definedNames
            std::string workbookXml = testDoc.getRawXml("xl/workbook.xml");

            // After moving AuditSheet (was index 3, localSheetId 2) to index 1:
            // New order: AuditSheet (localSheetId 0), Sheet1 (localSheetId 1), DataSheet (localSheetId 2)

            // AuditSheet was localSheetId 2, now it should be 0
            REQUIRE(workbookXml.find("name=\"LocalName2\" localSheetId=\"0\"") != std::string::npos);

            // DataSheet was localSheetId 1, now it should be 2
            REQUIRE(workbookXml.find("name=\"LocalName1\" localSheetId=\"2\"") != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify Merge Cells OOXML structure")
    {
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            wks.mergeCells("A1:B2");
            wks.mergeCells("C3:D4");

            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            std::string sheetXml = testDoc.getRawXml("xl/worksheets/sheet1.xml");

            // 1. Verify mergeCells container and count attribute
            REQUIRE(sheetXml.find("<mergeCells count=\"2\">") != std::string::npos);

            // 2. Verify individual mergeCell elements
            REQUIRE(sheetXml.find("<mergeCell ref=\"A1:B2\" />") != std::string::npos);
            REQUIRE(sheetXml.find("<mergeCell ref=\"C3:D4\" />") != std::string::npos);

            testDoc.close();
        }

        // Test deletion and empty container cleanup
        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");

            wks.unmergeCells("A1:B2");
            wks.unmergeCells("C3:D4");

            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            std::string sheetXml = testDoc.getRawXml("xl/worksheets/sheet1.xml");

            // 1. Verify mergeCells container is GONE (Excel doesn't allow empty mergeCells tag)
            REQUIRE(sheetXml.find("<mergeCells") == std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }

    SECTION("Verify Range Style Inheritance and XML Integrity")
    {
        std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // 1. Create a custom style to get an index != 0
            auto styles       = doc.styles();
            auto newFontIndex = styles.fonts().create();
            styles.fonts()[newFontIndex].setFontName("Arial Black");

            auto newXfIndex = styles.cellFormats().create();
            styles.cellFormats()[newXfIndex].setFontIndex(newFontIndex);

            XLStyleIndex customStyleIndex = newXfIndex;

            // Ensure our style index is indeed not 0
            REQUIRE(customStyleIndex > 0);

            // 2. Set column style using our custom style
            wks.column(2).setFormat(customStyleIndex);

            // 3. Create a range and iterate to fill values
            // The iterator should pick up customStyleIndex for column 2 and
            // explicitly set it on cell B2 because it's not the default 0.
            auto rng = wks.range(XLCellReference("B2"), XLCellReference("B2"));
            for (auto& cell : rng) { cell.value() = "Inherited"; }

            doc.save();
            doc.close();
        }

        {
            XLDocumentTest testDoc;
            testDoc.open(filename);

            std::string sheetXml = testDoc.getRawXml("xl/worksheets/sheet1.xml");

            // Verify column definition exists with correct style
            REQUIRE(sheetXml.find("<col min=\"2\" max=\"2\"") != std::string::npos);

            size_t colStylePos = sheetXml.find("style=\"");
            REQUIRE(colStylePos != std::string::npos);

            // Extract the style index string until the closing quote
            size_t      styleEndPos = sheetXml.find("\"", colStylePos + 7);
            std::string styleStr    = sheetXml.substr(colStylePos + 7, styleEndPos - (colStylePos + 7));

            // Verify cell B2 exists and explicitly inherited the non-default style
            std::string expectedCell = "<c r=\"B2\" s=\"" + styleStr + "\"";
            REQUIRE(sheetXml.find(expectedCell) != std::string::npos);

            testDoc.close();
        }
        std::filesystem::remove(filename);
    }
}

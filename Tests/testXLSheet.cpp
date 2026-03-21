#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>

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

TEST_CASE("XLSheet Tests", "[XLSheet]")
{
    SECTION("XLSheet Visibility")
    {
        XLDocument doc;
        doc.create("./testXLSheet1.xlsx", XLForceOverwrite);

        auto wks1 = doc.workbook().sheet(1);
        wks1.setName("VeryHidden");
        REQUIRE(wks1.name() == "VeryHidden");

        doc.workbook().addWorksheet("Hidden");
        auto wks2 = doc.workbook().sheet("Hidden");
        REQUIRE(wks2.name() == "Hidden");

        doc.workbook().addWorksheet("Visible");
        auto wks3 = doc.workbook().sheet("Visible");
        REQUIRE(wks3.name() == "Visible");

        REQUIRE(wks1.visibility() == XLSheetState::Visible);
        REQUIRE(wks2.visibility() == XLSheetState::Visible);
        REQUIRE(wks3.visibility() == XLSheetState::Visible);

        wks1.setVisibility(XLSheetState::VeryHidden);
        REQUIRE(wks1.visibility() == XLSheetState::VeryHidden);

        wks2.setVisibility(XLSheetState::Hidden);
        REQUIRE(wks2.visibility() == XLSheetState::Hidden);

        REQUIRE_THROWS(wks3.setVisibility(XLSheetState::Hidden));
        wks3.setVisibility(XLSheetState::Visible);

        doc.save();
    }

    SECTION("XLSheet Tab Color")
    {
        XLDocument doc;
        doc.create("./testXLSheet2.xlsx", XLForceOverwrite);

        auto wks1 = doc.workbook().sheet(1);
        wks1.setName("Sheet1");
        REQUIRE(wks1.name() == "Sheet1");

        doc.workbook().addWorksheet("Sheet2");
        auto wks2 = doc.workbook().sheet("Sheet2");
        REQUIRE(wks2.name() == "Sheet2");

        doc.workbook().addWorksheet("Sheet3");
        auto wks3 = doc.workbook().sheet("Sheet3");
        REQUIRE(wks3.name() == "Sheet3");

        wks1.setColor(XLColor(255, 0, 0));
        REQUIRE(wks1.color() == XLColor(255, 0, 0));
        REQUIRE_FALSE(wks1.color() == XLColor(0, 0, 0));

        wks2.setColor(XLColor(0, 255, 0));
        REQUIRE(wks2.color() == XLColor(0, 255, 0));
        REQUIRE_FALSE(wks2.color() == XLColor(0, 0, 0));

        wks3.setColor(XLColor(0, 0, 255));
        REQUIRE(wks3.color() == XLColor(0, 0, 255));
        REQUIRE_FALSE(wks3.color() == XLColor(0, 0, 0));

        doc.save();
    }

    SECTION("XLSheet AutoFilter")
    {
        XLDocument doc;
        doc.create("./testXLSheet3.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        REQUIRE_FALSE(wks.hasAutoFilter());
        REQUIRE(wks.autoFilter().empty());

        wks.setAutoFilter(wks.range("A1", "C10"));
        REQUIRE(wks.hasAutoFilter());
        REQUIRE(wks.autoFilter() == "A1:C10");

        wks.clearAutoFilter();
        REQUIRE_FALSE(wks.hasAutoFilter());
        REQUIRE(wks.autoFilter().empty());

        doc.save();
        }

        SECTION("XLSheet Panes (Freeze & Split)")
        {
            XLDocument doc;
            doc.create("./testXLSheetPanes.xlsx", XLForceOverwrite);

            auto wks = doc.workbook().worksheet("Sheet1");

            // 1. Freeze first row
            wks.freezePanes(0, 1);
            REQUIRE(wks.hasPanes());
            std::ostringstream ss1;
            wks.xmlDocument().document_element().print(ss1);
            std::string xml = ss1.str();
            REQUIRE(xml.find("ySplit=\"1\"") != std::string::npos);
            REQUIRE(xml.find("state=\"frozen\"") != std::string::npos);
            REQUIRE(xml.find("topLeftCell=\"A2\"") != std::string::npos);
            REQUIRE(xml.find("<selection pane=\"bottomLeft\"") != std::string::npos);

            // 2. Freeze first column
            wks.freezePanes(1, 0);
            REQUIRE(wks.hasPanes());
            std::ostringstream ss2;
            wks.xmlDocument().document_element().print(ss2);
            xml = ss2.str();
            REQUIRE(xml.find("xSplit=\"1\"") != std::string::npos);
            REQUIRE(xml.find("state=\"frozen\"") != std::string::npos);
            REQUIRE(xml.find("activePane=\"topRight\"") != std::string::npos);
            REQUIRE(xml.find("<selection pane=\"topRight\"") != std::string::npos);

            // 3. Freeze first row and first column
            wks.freezePanes(1, 1);
            REQUIRE(wks.hasPanes());
            std::ostringstream ss3;
            wks.xmlDocument().document_element().print(ss3);
            xml = ss3.str();
            REQUIRE(xml.find("xSplit=\"1\"") != std::string::npos);
            REQUIRE(xml.find("ySplit=\"1\"") != std::string::npos);
            REQUIRE(xml.find("state=\"frozen\"") != std::string::npos);
            REQUIRE(xml.find("activePane=\"bottomRight\"") != std::string::npos);
            REQUIRE(xml.find("<selection pane=\"topRight\"") != std::string::npos);
            REQUIRE(xml.find("<selection pane=\"bottomLeft\"") != std::string::npos);
            REQUIRE(xml.find("<selection pane=\"bottomRight\"") != std::string::npos);

            // 4. Split panes
            wks.splitPanes(1000, 2000, "C3", XLPane::BottomRight);
            REQUIRE(wks.hasPanes());
            std::ostringstream ss4;
            wks.xmlDocument().document_element().print(ss4);
            xml = ss4.str();
            REQUIRE(xml.find("xSplit=\"1000\"") != std::string::npos);
            REQUIRE(xml.find("ySplit=\"2000\"") != std::string::npos);
            REQUIRE(xml.find("state=\"split\"") != std::string::npos);
            REQUIRE(xml.find("activePane=\"bottomRight\"") != std::string::npos);

            // 5. Clear panes
            wks.clearPanes();
            REQUIRE_FALSE(wks.hasPanes());
            std::ostringstream ss5;
            wks.xmlDocument().document_element().print(ss5);
            xml = ss5.str();
            REQUIRE(xml.find("<pane") == std::string::npos);
            REQUIRE(xml.find("<selection") == std::string::npos);

            // 6. Freeze using string cell reference overload
            wks.freezePanes("B2");
            REQUIRE(wks.hasPanes());
            std::ostringstream ss6;
            wks.xmlDocument().document_element().print(ss6);
            xml = ss6.str();
            REQUIRE(xml.find("xSplit=\"1\"") != std::string::npos);
            REQUIRE(xml.find("ySplit=\"1\"") != std::string::npos);
            REQUIRE(xml.find("state=\"frozen\"") != std::string::npos);
            REQUIRE(xml.find("activePane=\"bottomRight\"") != std::string::npos);
            doc.save();
        }

        SECTION("XLSheet Grouping (Outline)")
        {
            XLDocument doc;
            doc.create("./testXLSheetGrouping.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // Group rows 2-5 (Level 1), not collapsed
            wks.groupRows(2, 5, 1, false);

            // Group rows 3-4 (Level 2), collapsed
            wks.groupRows(3, 4, 2, true);

            // Group columns B-D (2-4, Level 1), not collapsed
            wks.groupColumns(2, 4, 1, false);

            // Check Rows
            REQUIRE(wks.row(2).outlineLevel() == 1);
            REQUIRE_FALSE(wks.row(2).isHidden());
            
            REQUIRE(wks.row(3).outlineLevel() == 2);
            REQUIRE(wks.row(3).isHidden());
            
            REQUIRE(wks.row(4).outlineLevel() == 2);
            REQUIRE(wks.row(4).isHidden());
            
            REQUIRE(wks.row(5).outlineLevel() == 1);
            REQUIRE(wks.row(5).isCollapsed() == true); // Because level 2 above it was collapsed
            
            // Check Columns
            REQUIRE(wks.column(2).outlineLevel() == 1);
            REQUIRE(wks.column(3).outlineLevel() == 1);
            REQUIRE(wks.column(4).outlineLevel() == 1);
            
            doc.save();
        }

        SECTION("XLSheet Column Formatting")

    {
        XLDocument doc;
        doc.create("./testXLSheet4.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // Test Page Margins
        auto margins = wks.pageMargins();
        margins.setLeft(1.0);
        margins.setRight(1.0);
        margins.setTop(1.5);
        margins.setBottom(1.5);
        margins.setHeader(0.5);
        margins.setFooter(0.5);

        REQUIRE(margins.left() == 1.0);
        REQUIRE(margins.right() == 1.0);
        REQUIRE(margins.top() == 1.5);
        REQUIRE(margins.bottom() == 1.5);
        REQUIRE(margins.header() == 0.5);
        REQUIRE(margins.footer() == 0.5);

        // Test Print Options
        auto printOpts = wks.printOptions();
        printOpts.setGridLines(true);
        printOpts.setHeadings(true);
        printOpts.setHorizontalCentered(true);
        printOpts.setVerticalCentered(false);

        REQUIRE(printOpts.gridLines() == true);
        REQUIRE(printOpts.headings() == true);
        REQUIRE(printOpts.horizontalCentered() == true);
        REQUIRE(printOpts.verticalCentered() == false);

        // Test Page Setup
        auto setup = wks.pageSetup();
        setup.setPaperSize(9);    // A4
        setup.setOrientation(XLPageOrientation::Landscape);
        setup.setScale(85);
        setup.setBlackAndWhite(true);

        REQUIRE(setup.paperSize() == 9);
        REQUIRE(setup.orientation() == XLPageOrientation::Landscape);
        REQUIRE(setup.scale() == 85);
        REQUIRE(setup.blackAndWhite() == true);

        setup.setPageOrder("overThenDown");
        setup.setUseFirstPageNumber(true);
        setup.setFirstPageNumber(2);
        
        REQUIRE(setup.pageOrder() == "overThenDown");
        REQUIRE(setup.useFirstPageNumber() == true);
        REQUIRE(setup.firstPageNumber() == 2);

        // Test Header Footer
        auto hf = wks.headerFooter();
        hf.setDifferentFirst(true);
        hf.setDifferentOddEven(true);
        hf.setAlignWithMargins(false);
        hf.setOddHeader("&L&D&T&RPage &P of &N");
        hf.setOddFooter("&CConfidential");
        hf.setEvenHeader("&R&D");
        hf.setFirstHeader("&CFirst Page Header");

        REQUIRE(hf.differentFirst() == true);
        REQUIRE(hf.differentOddEven() == true);
        REQUIRE(hf.alignWithMargins() == false);
        REQUIRE(hf.oddHeader() == "&L&D&T&RPage &P of &N");
        REQUIRE(hf.oddFooter() == "&CConfidential");
        REQUIRE(hf.evenHeader() == "&R&D");
        REQUIRE(hf.firstHeader() == "&CFirst Page Header");

        doc.save();
        doc.close();

        // Re-open and verify
        XLDocument doc2;
        doc2.open("./testXLSheet4.xlsx");
        auto wks2 = doc2.workbook().worksheet("Sheet1");

        REQUIRE(wks2.pageMargins().left() == 1.0);
        REQUIRE(wks2.printOptions().gridLines() == true);
        REQUIRE(wks2.pageSetup().orientation() == XLPageOrientation::Landscape);

        REQUIRE(wks2.headerFooter().oddHeader() == "&L&D&T&RPage &P of &N");
        REQUIRE(wks2.headerFooter().differentFirst() == true);

        // Verify XML structure
        std::string sheetXml = getRawXml(doc2, "xl/worksheets/sheet1.xml");

        // 1. Check margins
        REQUIRE(sheetXml.find("<pageMargins") != std::string::npos);
        REQUIRE(sheetXml.find("left=\"1.000000\"") != std::string::npos);

        // 2. Check print options
        REQUIRE(sheetXml.find("<printOptions") != std::string::npos);
        REQUIRE(sheetXml.find("gridLines=\"1\"") != std::string::npos);
        REQUIRE(sheetXml.find("headings=\"1\"") != std::string::npos);
        REQUIRE(sheetXml.find("horizontalCentered=\"1\"") != std::string::npos);

        // 3. Check page setup
        REQUIRE(sheetXml.find("<pageSetup") != std::string::npos);
        REQUIRE(sheetXml.find("orientation=\"landscape\"") != std::string::npos);
        REQUIRE(sheetXml.find("paperSize=\"9\"") != std::string::npos);
        REQUIRE(sheetXml.find("scale=\"85\"") != std::string::npos);
        REQUIRE(sheetXml.find("blackAndWhite=\"1\"") != std::string::npos);

        // 4. Check ordering: printOptions < pageMargins < pageSetup
        size_t printPos  = sheetXml.find("<printOptions");
        size_t marginPos = sheetXml.find("<pageMargins");
        size_t setupPos  = sheetXml.find("<pageSetup");

        REQUIRE(printPos != std::string::npos);
        REQUIRE(marginPos != std::string::npos);
        REQUIRE(setupPos != std::string::npos);
        REQUIRE(printPos < marginPos);
        REQUIRE(marginPos < setupPos);

        // 5. Check headerFooter
        REQUIRE(sheetXml.find("<headerFooter") != std::string::npos);
        REQUIRE(sheetXml.find("differentFirst=\"1\"") != std::string::npos);
        REQUIRE(sheetXml.find("differentOddEven=\"1\"") != std::string::npos);
        REQUIRE(sheetXml.find("alignWithMargins=\"0\"") != std::string::npos);
        REQUIRE(sheetXml.find("<oddHeader>&amp;L&amp;D&amp;T&amp;RPage &amp;P of &amp;N</oddHeader>") != std::string::npos);
        REQUIRE(sheetXml.find("<oddFooter>&amp;CConfidential</oddFooter>") != std::string::npos);
        REQUIRE(sheetXml.find("<evenHeader>&amp;R&amp;D</evenHeader>") != std::string::npos);
        REQUIRE(sheetXml.find("<firstHeader>&amp;CFirst Page Header</firstHeader>") != std::string::npos);
        
        size_t hfPos = sheetXml.find("<headerFooter");
        REQUIRE(setupPos < hfPos);

        doc2.close();
        }


        SECTION("XLSheet HeaderFooter OOXML Strict Node Order")
        {
            XLDocument doc;
            doc.create("./testXLSheet_HF_Order.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // Intentionally set in reverse/random order to trigger the NodeOrder re-arrangement
            auto hf = wks.headerFooter();
            hf.setFirstFooter("FirstFooter");
            hf.setFirstHeader("FirstHeader");
            hf.setEvenFooter("EvenFooter");
            hf.setEvenHeader("EvenHeader");
            hf.setOddFooter("OddFooter");
            hf.setOddHeader("OddHeader");

            doc.save();
            doc.close();

            // Re-open and verify functionality
            XLDocument doc2;
            doc2.open("./testXLSheet_HF_Order.xlsx");
            auto wks2 = doc2.workbook().worksheet("Sheet1");
            auto hf2 = wks2.headerFooter();

            REQUIRE(hf2.oddHeader() == "OddHeader");
            REQUIRE(hf2.oddFooter() == "OddFooter");
            REQUIRE(hf2.evenHeader() == "EvenHeader");
            REQUIRE(hf2.evenFooter() == "EvenFooter");
            REQUIRE(hf2.firstHeader() == "FirstHeader");
            REQUIRE(hf2.firstFooter() == "FirstFooter");

            // Verify strict OOXML ordering in XML output
            std::string sheetXml = getRawXml(doc2, "xl/worksheets/sheet1.xml");

            size_t oddH  = sheetXml.find("<oddHeader>");
            size_t oddF  = sheetXml.find("<oddFooter>");
            size_t evenH = sheetXml.find("<evenHeader>");
            size_t evenF = sheetXml.find("<evenFooter>");
            size_t firstH= sheetXml.find("<firstHeader>");
            size_t firstF= sheetXml.find("<firstFooter>");

            REQUIRE(oddH != std::string::npos);
            REQUIRE(oddF != std::string::npos);
            REQUIRE(evenH != std::string::npos);
            REQUIRE(evenF != std::string::npos);
            REQUIRE(firstH != std::string::npos);
            REQUIRE(firstF != std::string::npos);

            // Strict OOXML sequence requirement: oddHeader, oddFooter, evenHeader, evenFooter, firstHeader, firstFooter
            REQUIRE(oddH < oddF);
            REQUIRE(oddF < evenH);
            REQUIRE(evenH < evenF);
            REQUIRE(evenF < firstH);
            REQUIRE(firstH < firstF);
            
            doc2.close();
        }

        SECTION("XLSheet Protection")
        {
            XLDocument doc;
            doc.create("./testXLSheet5.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // 1. Test safety: should not crash and should return defaults (true = allowed) when node doesn't exist
            REQUIRE_FALSE(wks.sheetProtected());
            REQUIRE(wks.insertColumnsAllowed());
            REQUIRE(wks.formatCellsAllowed());
            REQUIRE(wks.passwordHash().empty());

            // 2. Test basic protection
            wks.protectSheet(true);
            REQUIRE(wks.sheetProtected() == true);

            // 3. Test detailed settings
            wks.allowInsertColumns(true);
            REQUIRE(wks.insertColumnsAllowed() == true);

            wks.denyInsertRows();
            REQUIRE(wks.insertRowsAllowed() == false);

            wks.allowFormatCells(true);
            REQUIRE(wks.formatCellsAllowed() == true);

            wks.denyFormatColumns();
            REQUIRE(wks.formatColumnsAllowed() == false);

            wks.allowAutoFilter(true);
            REQUIRE(wks.autoFilterAllowed() == true);

            wks.denySort();
            REQUIRE(wks.sortAllowed() == false);

            // 4. Test password
            wks.setPassword("OpenXLSX");
            REQUIRE(wks.passwordIsSet() == true);
            REQUIRE(wks.passwordHash() == "a355");

            doc.save();

            // 5. Verify XML structure
            std::string sheetXml = getRawXml(doc, "xl/worksheets/sheet1.xml");
            REQUIRE(sheetXml.find("<sheetProtection") != std::string::npos);
            REQUIRE(sheetXml.find("sheet=\"true\"") != std::string::npos);
            REQUIRE(sheetXml.find("insertColumns=\"false\"") != std::string::npos); // allow = true -> XML "false" (not protected)
            REQUIRE(sheetXml.find("insertRows=\"true\"") != std::string::npos);    // allow = false -> XML "true" (protected)
            REQUIRE(sheetXml.find("formatCells=\"false\"") != std::string::npos);
            REQUIRE(sheetXml.find("formatColumns=\"true\"") != std::string::npos);
            REQUIRE(sheetXml.find("autoFilter=\"false\"") != std::string::npos);
            REQUIRE(sheetXml.find("sort=\"true\"") != std::string::npos);
            REQUIRE(sheetXml.find("password=\"a355\"") != std::string::npos);

            // 6. Test clear functions
            wks.clearPassword();
            REQUIRE_FALSE(wks.passwordIsSet());

            wks.clearSheetProtection();
            REQUIRE_FALSE(wks.sheetProtected());
            REQUIRE(wks.insertRowsAllowed()); // Should revert to default (allowed)

            doc.close();
        }

        SECTION("XLSheet Views and Page Breaks")
        {
            XLDocument doc;
            doc.create("./testXLSheet6.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // Test Zoom
            REQUIRE(wks.zoom() == 100); // Default
            wks.setZoom(150);
            REQUIRE(wks.zoom() == 150);

            // Test Page Breaks
            wks.insertRowBreak(5);
            wks.insertRowBreak(10);
            
            doc.save();
            doc.close();

            // Re-open and verify
            XLDocument doc2;
            doc2.open("./testXLSheet6.xlsx");
            auto wks2 = doc2.workbook().worksheet("Sheet1");

            REQUIRE(wks2.zoom() == 150);

            // Verify XML structure
            std::string sheetXml = getRawXml(doc2, "xl/worksheets/sheet1.xml");
            
            // Zoom
            REQUIRE(sheetXml.find("zoomScale=\"150\"") != std::string::npos);
            REQUIRE(sheetXml.find("zoomScaleNormal=\"150\"") != std::string::npos);

            // Page Breaks
            REQUIRE(sheetXml.find("<rowBreaks") != std::string::npos);
            REQUIRE(sheetXml.find("count=\"2\"") != std::string::npos);
            REQUIRE(sheetXml.find("manualBreakCount=\"2\"") != std::string::npos);
            REQUIRE(sheetXml.find("<brk id=\"5\"") != std::string::npos);
            REQUIRE(sheetXml.find("<brk id=\"10\"") != std::string::npos);

            // Test Removal
            wks2.removeRowBreak(5);
            doc2.save();
            doc2.close();

            XLDocument doc3;
            doc3.open("./testXLSheet6.xlsx");
            std::string sheetXml3 = getRawXml(doc3, "xl/worksheets/sheet1.xml");
            REQUIRE(sheetXml3.find("count=\"1\"") != std::string::npos);
            REQUIRE(sheetXml3.find("id=\"5\"") == std::string::npos);
            REQUIRE(sheetXml3.find("id=\"10\"") != std::string::npos);
            doc3.close();
        }

}
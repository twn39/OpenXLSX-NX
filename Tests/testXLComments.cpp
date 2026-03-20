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
    std::string (XLDocument::*get_impl(XLDocument_extractXmlFromArchive))(std::string_view);

    // Function to call the protected member
    std::string getRawXml(XLDocument& doc, const std::string& path)
    {
        static auto fn = get_impl(XLDocument_extractXmlFromArchive());
        return (doc.*fn)(path);
    }
}    // namespace

TEST_CASE("XLComments Tests", "[XLComments]")
{
    SECTION("Basic Comment Operations")
    {
        XLDocument doc;
        doc.create("./testXLComments.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        REQUIRE(wks.hasComments() == false);

        // Add a comment
        wks.comments().addAuthor("Author1");
        wks.comments().set("A1", "This is a comment", 0);

        REQUIRE(wks.hasComments() == true);
        REQUIRE(wks.comments().count() == 1);
        REQUIRE(wks.comments().get("A1") == "This is a comment");
        REQUIRE(wks.comments().authorCount() == 1);
        REQUIRE(wks.comments().author(0) == "Author1");

        // Delete comment
        wks.comments().deleteComment("A1");
        REQUIRE(wks.comments().count() == 0);
        REQUIRE(wks.comments().get("A1") == "");

        doc.close();
    }

    SECTION("Multiple Authors and Comments")
    {
        XLDocument doc;
        doc.create("./testXLCommentsMultiple.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        auto&    comments = wks.comments();
        uint16_t auth1    = comments.addAuthor("Author One");
        uint16_t auth2    = comments.addAuthor("Author Two");

        comments.set("B2", "Comment by Auth 1", auth1);
        comments.set("C3", "Comment by Auth 2", auth2);

        REQUIRE(comments.count() == 2);
        REQUIRE(comments.authorId("B2") == auth1);
        REQUIRE(comments.authorId("C3") == auth2);

        REQUIRE(comments.get(0).ref() == "B2");
        REQUIRE(comments.get(1).ref() == "C3");

        doc.close();
    }

    SECTION("Comment Shape Properties")
    {
        XLDocument doc;
        doc.create("./testXLCommentsShapes.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.comments().set("D4", "Shape test");
        auto shape = wks.comments().shape("D4");

        // Test shape visibility
        shape.style().show();
        REQUIRE(shape.style().visible() == true);
        shape.style().hide();
        REQUIRE(shape.style().visible() == false);

        doc.close();
    }

    SECTION("Rich Text Comment Operations")
    {
        XLDocument doc;
        doc.create("./testXLCommentsRichText.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        XLRichText rt;
        XLRichTextRun r1("Bold"); r1.setBold(true);
        XLRichTextRun r2(" Red"); r2.setFontColor(XLColor(255, 0, 0));
        rt.addRun(r1);
        rt.addRun(r2);

        wks.comments().setRichText("A1", rt);
        
        REQUIRE(wks.comments().count() == 1);
        REQUIRE(wks.comments().get("A1") == "Bold Red");

        doc.save();
        
        // Verify XML structure
        std::string commentsXml = getRawXml(doc, "xl/comments1.xml");
        REQUIRE(commentsXml.find("<r>") != std::string::npos);
        REQUIRE(commentsXml.find("<rPr>") != std::string::npos);
        REQUIRE(commentsXml.find("<b />") != std::string::npos);
        REQUIRE(commentsXml.find("<t>Bold</t>") != std::string::npos);
        REQUIRE(commentsXml.find("<color rgb=\"ffff0000\" />") != std::string::npos);
        REQUIRE(commentsXml.find("<t xml:space=\"preserve\"> Red</t>") != std::string::npos);

        doc.close();

        // Re-open and verify formatting
        XLDocument doc2;
        doc2.open("./testXLCommentsRichText.xlsx");
        auto wks2 = doc2.workbook().worksheet("Sheet1");
        
        auto rt2 = wks2.comments().get(0).richText();
        REQUIRE(rt2.runs().size() == 2);
        REQUIRE(rt2.runs()[0].text() == "Bold");
        REQUIRE(rt2.runs()[0].bold() == true);
        REQUIRE(rt2.runs()[1].text() == " Red");
        REQUIRE(rt2.runs()[1].fontColor()->hex() == "ffff0000");

        doc2.close();
    }

    SECTION("Optimized Comments Overloads, Deduplication and Custom Size")
    {
        XLDocument doc;
        doc.create("./testXLCommentsOptimizations.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // 1. Author Deduplication Test
        auto& comments = wks.comments();
        uint16_t id1 = comments.addAuthor("Dedupe Tester");
        uint16_t id2 = comments.addAuthor("Dedupe Tester"); // Should return same ID
        REQUIRE(id1 == id2);
        REQUIRE(comments.authorCount() == 1);

        // 2. String Author API & Size Dimensions Overload
        comments.set("A1", "Custom Size Comment", "Dedupe Tester", 5, 8);
        REQUIRE(comments.authorId("A1") == id1); // should match deduplicated ID
        REQUIRE(comments.count() == 1);
        
        // 3. Visibility toggling via explicit API
        comments.setVisible("A1", true);
        REQUIRE(comments.shape("A1").style().visible() == true);
        comments.setVisible("A1", false);
        REQUIRE(comments.shape("A1").style().visible() == false);

        doc.save();

        // 4. Verify OOXML VML format
        std::string vmlXml = getRawXml(doc, "xl/drawings/vmlDrawing1.vml");
        // We know A1 is column 0, width=5 => right column=5. height=8 => bottom row=8. 
        // Our anchor logic generates coordinate string like "5,10,0,5,9,10,8,5" or similar (Left,Top,Right,Bottom depends on bounds checks).
        // Since destCol=1, destRow=1: 
        // LeftCol = (1-1)+1 = 1. RightCol = (1-1)+1+5 = 6.
        // TopRow = (1-1)+1 = 1. BottomRow = (1-1)+1+8 = 9.
        // The expected string contains "1,10,1,5,6,10,9,5"
        REQUIRE(vmlXml.find("1,10,1,5,6,10,9,5") != std::string::npos);
        
        // Verify style visibility
        REQUIRE(vmlXml.find("visibility:hidden") != std::string::npos);

        doc.close();
    }
}

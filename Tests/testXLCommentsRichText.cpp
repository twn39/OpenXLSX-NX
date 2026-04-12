#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("CommentsRichTextValidation", "[Comments][RichText]")
{
    XLDocument doc;
    doc.create("Comments_RichText_Test.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Hover me";

    XLRichText    rt;
    XLRichTextRun run1("Warning: ");
    run1.setBold(true);
    run1.setFontColor(XLColor("FFFF0000"));    // Red
    rt.addRun(run1);

    XLRichTextRun run2("This is a rich text comment!");
    run2.setItalic(true);
    run2.setFontColor(XLColor("FF0000FF"));    // Blue
    rt.addRun(run2);

    wks.comments().setRichText("A1", rt, "Author Name", 4, 6);

    // Save and reload to verify
    doc.save();
    doc.close();

    doc.open("Comments_RichText_Test.xlsx");
    auto wks2 = doc.workbook().worksheet("Sheet1");

    // get comment by index since get(cellRef) returns std::string
    auto comment = wks2.comments().get(0);
    REQUIRE(comment.valid() == true);
    REQUIRE(comment.ref() == "A1");

    auto parsedRt = comment.richText();
    REQUIRE(parsedRt.runs().size() == 2);
    REQUIRE(parsedRt.runs()[0].text() == "Warning: ");
    REQUIRE(parsedRt.runs()[0].bold().value_or(false) == true);
    REQUIRE(parsedRt.runs()[0].fontColor()->hex() == "FFFF0000");

    REQUIRE(parsedRt.runs()[1].text() == "This is a rich text comment!");
    REQUIRE(parsedRt.runs()[1].italic().value_or(false) == true);
    REQUIRE(parsedRt.runs()[1].fontColor()->hex() == "FF0000FF");

    doc.close();
}

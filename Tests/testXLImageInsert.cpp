#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Image Insert API Tests", "[XLImageInsert]")
{
    XLDocument doc;
    doc.create("./testXLImageInsert.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Default insert (100% scale)
    wks.insertImage("B2", "Tests/test.png");

    // Scaled insert
    XLImageOptions opts;
    opts.scaleX = 0.5;
    opts.scaleY = 0.5;
    wks.insertImage("D5", "Tests/test.png", opts);

    doc.save();
    doc.close();

    // Verify it opened correctly
    XLDocument doc2;
    doc2.open("./testXLImageInsert.xlsx");
    REQUIRE(doc2.workbook().worksheet("Sheet1").images().size() == 2);
    doc2.close();
}

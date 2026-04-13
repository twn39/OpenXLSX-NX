#include "OpenXLSX.hpp"
#include <catch2/catch_test_macros.hpp>
#include <fstream>
#include <iostream>
#include <string>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__TableSlicerFullTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__TableSlicerFullTest_SaveAs_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("TableSlicerAPIandOOXMLValidation", "[XLSlicer]")
{
    // 1. Create a document with a table and two slicers
    XLDocument doc;
    doc.create(__global_unique_file_0(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Product";
    wks.cell("C1").value() = "Sales";

    wks.cell("A2").value() = "North";
    wks.cell("B2").value() = "Apple";
    wks.cell("C2").value() = 100;

    wks.cell("A3").value() = "South";
    wks.cell("B3").value() = "Banana";
    wks.cell("C3").value() = 200;

    auto tables = wks.tables();
    auto table  = tables.add("SalesTable", "A1:C3");

    XLSlicerOptions opts1;
    opts1.name    = "SlicerRegion";
    opts1.caption = "Filter Region";
    wks.addTableSlicer("E2", table, "Region", opts1);

    XLSlicerOptions opts2;
    opts2.name    = "SlicerProduct";
    opts2.caption = "Filter Product";
    wks.addTableSlicer("G2", table, "Product", opts2);

    doc.save();

    // 2. Functional Reload and Relationship Verification
    XLDocument doc2;
    doc2.open(__global_unique_file_0());
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    // Verify worksheet has drawing (since slicers are injected into drawings)
    REQUIRE(wks2.hasDrawing() == true);

    // Verify relations and content types via accessing raw XML
    // In OpenXLSX we can access raw XML node trees
    auto sheetNode = wks2.xmlDocument().document_element();
    auto extLst    = sheetNode.child("extLst");
    REQUIRE(!extLst.empty());

    // Check for slicerList extension
    auto ext = extLst.find_child_by_attribute("uri", "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}");
    REQUIRE(!ext.empty());

    auto slicerList = ext.child("x14:slicerList");
    REQUIRE(!slicerList.empty());

    // We added 2 slicers, so there should be 2 x14:slicer relationships
    int slicerRefCount = 0;
    for (auto s : slicerList.children("x14:slicer")) {
        slicerRefCount++;
        REQUIRE(!s.attribute("r:id").empty());
    }
    REQUIRE(slicerRefCount == 2);

    // Verify Drawing XML structure
    auto drwRoot           = wks2.drawing().xmlDocument().document_element();
    int  anchorCount       = 0;
    int  graphicFrameCount = 0;
    int  fallbackCount     = 0;
    for (auto anchor : drwRoot.children("xdr:twoCellAnchor")) {
        anchorCount++;
        auto altContent = anchor.child("mc:AlternateContent");
        REQUIRE(!altContent.empty());

        auto choice = altContent.child("mc:Choice");
        REQUIRE(!choice.empty());
        REQUIRE(std::string(choice.attribute("Requires").value()) == "sle15");

        auto graphicFrame = choice.child("xdr:graphicFrame");
        if (!graphicFrame.empty()) {
            graphicFrameCount++;
            // Check that graphic uri is correct
            auto graphicData = graphicFrame.child("a:graphic").child("a:graphicData");
            REQUIRE(std::string(graphicData.attribute("uri").value()) == "http://schemas.microsoft.com/office/drawing/2010/slicer");
            auto sleSlicer = graphicData.child("sle:slicer");
            REQUIRE(!sleSlicer.empty());
        }

        auto fallback = altContent.child("mc:Fallback");
        if (!fallback.empty()) {
            fallbackCount++;
            REQUIRE(!fallback.child("xdr:sp").empty());
        }
    }

    // 2 slicers should generate 2 anchors, 2 graphic frames, and 2 fallbacks
    REQUIRE(anchorCount == 2);
    REQUIRE(graphicFrameCount == 2);
    REQUIRE(fallbackCount == 2);

    // 3. Document Level Validation
    auto wbkNode   = doc2.workbook().xmlDocument().document_element();
    auto wbkExtLst = wbkNode.child("extLst");
    REQUIRE(!wbkExtLst.empty());

    auto wbkExt = wbkExtLst.find_child_by_attribute("uri", "{46BE6895-7355-4a93-B00E-2C351335B9C9}");
    REQUIRE(!wbkExt.empty());

    auto slicerCaches = wbkExt.child("x15:slicerCaches");
    REQUIRE(!slicerCaches.empty());

    int cacheRefCount = 0;
    for (auto c : slicerCaches.children("x14:slicerCache")) {
        cacheRefCount++;
        REQUIRE(!c.attribute("r:id").empty());
    }
    REQUIRE(cacheRefCount == 2);

    // 4. Double Save to ensure memory stream doesn't break dependencies
    doc2.saveAs(__global_unique_file_1(), XLForceOverwrite);
    doc2.close();
    XLDocument doc3;
    doc3.open(__global_unique_file_1());
    REQUIRE(doc3.workbook().worksheet("Sheet1").hasDrawing() == true);

    // Ensure table parsing still works
    auto tb = doc3.workbook().worksheet("Sheet1").tables().table("SalesTable");
    REQUIRE(tb.valid() == true);
    REQUIRE(tb.name() == "SalesTable");
}

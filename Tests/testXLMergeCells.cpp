#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>

using namespace OpenXLSX;

TEST_CASE("XLMergeCellsTests", "[XLMergeCells]")
{
    SECTION("Basic Merge and Unmerge")
    {
        XLDocument doc;
        doc.create("./testXLMergeCells.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("A1:B2");
        REQUIRE(wks.merges().count() == 1);
        REQUIRE(std::string(wks.merges().merge(0)) == "A1:B2");
        REQUIRE(wks.merges().mergeExists("A1:B2") == true);

        wks.unmergeCells("A1:B2");
        REQUIRE(wks.merges().count() == 0);
        REQUIRE(wks.merges().mergeExists("A1:B2") == false);

        doc.close();
    }

    SECTION("Multiple Merges")
    {
        XLDocument doc;
        doc.create("./testXLMultipleMerges.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("A1:A2");
        wks.mergeCells("C1:C2");
        REQUIRE(wks.merges().count() == 2);

        REQUIRE(wks.merges().findMerge("A1:A2") == 0);
        REQUIRE(wks.merges().findMerge("C1:C2") == 1);

        doc.close();
    }

    SECTION("Find Merge by Cell")
    {
        XLDocument doc;
        doc.create("./testXLMergeByCell.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("B2:C3");

        // Test with strings
        REQUIRE(wks.merges().findMergeByCell("B2") == 0);
        REQUIRE(wks.merges().findMergeByCell("C3") == 0);
        REQUIRE(wks.merges().findMergeByCell("B3") == 0);
        REQUIRE(wks.merges().findMergeByCell("C2") == 0);
        REQUIRE(wks.merges().findMergeByCell("A1") == XLMergeNotFound);

        // Test with XLCellReference (Numerical path)
        REQUIRE(wks.merges().findMergeByCell(XLCellReference(2, 2)) == 0);
        REQUIRE(wks.merges().findMergeByCell(XLCellReference(3, 3)) == 0);
        REQUIRE(wks.merges().findMergeByCell(XLCellReference(1, 1)) == XLMergeNotFound);

        doc.close();
    }

    SECTION("Numerical Cache Consistency after Deletion")
    {
        XLDocument doc;
        doc.create("./testXLMergeCacheConsistency.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("A1:B2");
        wks.mergeCells("D4:E5");
        wks.mergeCells("G7:H8");

        REQUIRE(wks.merges().count() == 3);
        REQUIRE(wks.merges().findMergeByCell("E5") == 1);

        // Delete middle one
        wks.unmergeCells("D4:E5");
        REQUIRE(wks.merges().count() == 2);

        // Check that subsequent index shifted correctly in cache
        REQUIRE(wks.merges().findMerge("G7:H8") == 1);
        REQUIRE(wks.merges().findMergeByCell("H8") == 1);
        REQUIRE(wks.merges().findMergeByCell("E5") == XLMergeNotFound);

        doc.close();
    }

    SECTION("Overlap Detection (Optimized Path)")
    {
        XLDocument doc;
        doc.create("./testXLMergeOverlap.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("B2:D4");

        // Overlapping cases (triggers XLRect::overlaps)
        REQUIRE_THROWS_AS(wks.mergeCells("C3:E5"), XLInputError);    // Partial overlap
        REQUIRE_THROWS_AS(wks.mergeCells("B2:D4"), XLInputError);    // Exact same
        REQUIRE_THROWS_AS(wks.mergeCells("A1:E5"), XLInputError);    // Encompassing

        // Touch but not overlap
        REQUIRE_NOTHROW(wks.mergeCells("E2:F4"));    // Right touch
        REQUIRE_NOTHROW(wks.mergeCells("B5:D6"));    // Bottom touch

        doc.close();
    }

    SECTION("Empty Hidden Cells")
    {
        XLDocument doc;
        doc.create("./testXLEmptyHiddenCells.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value() = 1;
        wks.cell("A2").value() = 2;
        wks.cell("B1").value() = 3;
        wks.cell("B2").value() = 4;

        wks.mergeCells("A1:B2", XLEmptyHiddenCells);

        REQUIRE(wks.cell("A1").value().get<int>() == 1);
        REQUIRE(wks.cell("A2").value().type() == XLValueType::Empty);
        REQUIRE(wks.cell("B1").value().type() == XLValueType::Empty);
        REQUIRE(wks.cell("B2").value().type() == XLValueType::Empty);

        doc.close();
    }
    SECTION("Get Value of Merged Cell")
    {
        XLDocument doc;
        doc.create("./testXLMergeGetValue.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("B2").value() = "Merged Title";
        wks.mergeCells("B2:D4", true);

        // The top-left cell should retain the value
        REQUIRE(wks.cell("B2").value().get<std::string>() == "Merged Title");

        // The rest should be empty
        REQUIRE(wks.cell("C3").value().type() == XLValueType::Empty);
        REQUIRE(wks.cell("D4").value().type() == XLValueType::Empty);

        // However, a good API should allow us to retrieve the top-left value
        // if we query ANY cell within the merged range.
        // Currently OpenXLSX returns Empty for C3. Let's see if there's a helper.
        // If not, we should test the finding mechanism.
        
        auto mergeIdx = wks.merges().findMergeByCell("C3");
        REQUIRE(mergeIdx != XLMergeNotFound);
        
        std::string mergeRange = std::string(wks.merges().merge(mergeIdx));
        REQUIRE(mergeRange == "B2:D4");
        
        // Extract top left cell from range
        XLCellReference topLeft(mergeRange.substr(0, mergeRange.find(':')));
        REQUIRE(wks.cell(topLeft).value().get<std::string>() == "Merged Title");

        doc.close();
        std::remove("./testXLMergeGetValue.xlsx");
    }

    SECTION("Boundary Overflow Tests")
    {
        XLDocument doc;
        doc.create("./testXLMergeBounds.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // Excel has max limits 1048576 x 16384
        REQUIRE_NOTHROW(wks.mergeCells("A1:XFD1048576"));
        REQUIRE(wks.merges().count() == 1);
        
        // Find merge in huge encompassing bounds
        REQUIRE(wks.merges().findMergeByCell("ZZ999") == 0);

        doc.close();
        std::remove("./testXLMergeBounds.xlsx");
    }
}

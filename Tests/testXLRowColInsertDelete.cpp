#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <string>

using namespace OpenXLSX;

// =============================================================================
// Row Insert / Delete Tests
// =============================================================================

TEST_CASE("XLRowColInsertDeleteTests", "[XLRowColInsertDelete]")
{
    SECTION("deleteRow basic — subsequent rows shift up")
    {
        XLDocument doc;
        doc.create("./testInsDelRow_basic.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value() = "Row1";
        wks.cell("A2").value() = "Row2";
        wks.cell("A3").value() = "Row3";
        wks.cell("A4").value() = "Row4";

        // Delete row 2 — rows 3,4 should slide to 2,3
        wks.deleteRow(2, 1);

        REQUIRE(wks.cell("A1").value().getString() == "Row1");
        REQUIRE(wks.cell("A2").value().getString() == "Row3");
        REQUIRE(wks.cell("A3").value().getString() == "Row4");
        REQUIRE(wks.cell("A4").value().type() == XLValueType::Empty);

        doc.close();
    }

    SECTION("insertRow basic — existing rows shift down")
    {
        XLDocument doc;
        doc.create("./testInsDelRow_insert.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value() = "Row1";
        wks.cell("A2").value() = "Row2";
        wks.cell("A3").value() = "Row3";

        // Insert one empty row before row 2 — rows 2,3 should become 3,4
        wks.insertRow(2, 1);

        REQUIRE(wks.cell("A1").value().getString() == "Row1");
        REQUIRE(wks.cell("A2").value().type() == XLValueType::Empty);    // new empty row
        REQUIRE(wks.cell("A3").value().getString() == "Row2");
        REQUIRE(wks.cell("A4").value().getString() == "Row3");

        doc.close();
    }

    SECTION("deleteRow multi-count — removes multiple rows and shifts")
    {
        XLDocument doc;
        doc.create("./testInsDelRow_multi.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        for (int i = 1; i <= 5; ++i) wks.cell(static_cast<uint32_t>(i), 1).value() = i * 10;

        // Delete rows 2 and 3 — rows 4,5 should become rows 2,3
        wks.deleteRow(2, 2);

        REQUIRE(wks.cell("A1").value().get<int>() == 10);
        REQUIRE(wks.cell("A2").value().get<int>() == 40);
        REQUIRE(wks.cell("A3").value().get<int>() == 50);
        REQUIRE(wks.cell("A4").value().type() == XLValueType::Empty);

        doc.close();
    }

    SECTION("insertRow multi-count — inserts multiple empty rows")
    {
        XLDocument doc;
        doc.create("./testInsDelRow_multiInsert.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value() = 100;
        wks.cell("A2").value() = 200;

        // Insert 3 rows before row 2
        wks.insertRow(2, 3);

        REQUIRE(wks.cell("A1").value().get<int>() == 100);
        REQUIRE(wks.cell("A2").value().type() == XLValueType::Empty);
        REQUIRE(wks.cell("A3").value().type() == XLValueType::Empty);
        REQUIRE(wks.cell("A4").value().type() == XLValueType::Empty);
        REQUIRE(wks.cell("A5").value().get<int>() == 200);

        doc.close();
    }

    SECTION("deleteRow updates formulas in cells below")
    {
        XLDocument doc;
        doc.create("./testInsDelRow_formula.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value()   = 1;
        wks.cell("A2").value()   = 2;
        wks.cell("A3").value()   = 3;
        wks.cell("A4").value()   = 4;
        wks.cell("A5").formula() = "=SUM(A1:A4)";    // formula in row 5, above result

        // Delete row 2 — formula range A1:A4 should become A1:A3
        wks.deleteRow(2, 1);

        std::string formulaAfter = wks.cell("A4").formula();
        // After delete: row 5 became row 4, formula must reference A1:A3
        REQUIRE(formulaAfter.find("A3") != std::string::npos);
        REQUIRE(formulaAfter.find("A4") == std::string::npos);

        doc.close();
    }

    SECTION("insertRow updates formula references")
    {
        XLDocument doc;
        doc.create("./testInsDelRow_formulaInsert.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value()   = 10;
        wks.cell("A2").value()   = 20;
        wks.cell("A3").formula() = "=A1+A2";    // references A1 and A2

        // Insert row before row 2 — A2 reference in the formula on (original) row 3
        // should become A3 (row 3 is pushed to row 4); A1 stay unchanged.
        wks.insertRow(2, 1);

        // Original row 3 is now row 4
        std::string formulaAfter = wks.cell("A4").formula();
        REQUIRE(formulaAfter.find("A1") != std::string::npos);
        REQUIRE(formulaAfter.find("A3") != std::string::npos);    // A2 → A3

        doc.close();
    }

    SECTION("deleteRow updates mergeCells below")
    {
        XLDocument doc;
        doc.create("./testInsDelRow_merge.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // Merge in row 1 (above delete) — should be unchanged
        wks.mergeCells("A1:B1");
        // Merge in row 4 (below delete) — should shift up by 2
        wks.mergeCells("A4:B5");

        // Delete rows 2 and 3
        wks.deleteRow(2, 2);

        REQUIRE(wks.merges().mergeExists("A1:B1"));    // unchanged
        REQUIRE(wks.merges().mergeExists("A2:B3"));    // shifted up by 2

        doc.close();
    }

    SECTION("deleteRow edge cases: collapsing and destruction of mergeCells")
    {
        XLDocument doc;
        doc.create("./testInsDelRow_mergeEdgeCases.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // 1. Merge cell that will be partially deleted (shrink/collapse)
        wks.mergeCells("A1:A5");    // 5 rows tall

        // 2. Merge cell that will be entirely engulfed and destroyed
        wks.mergeCells("C3:D4");    // 2 rows tall, exists entirely within the blast radius

        // 3. Merge cell overlapping the bottom edge of the deletion
        wks.mergeCells("F4:G6");    // 3 rows tall

        // Delete row 3 and 4 (count = 2)
        wks.deleteRow(3, 2);

        // A1:A5 should collapse by 2 rows -> A1:A3
        REQUIRE(wks.merges().mergeExists("A1:A3"));
        REQUIRE_FALSE(wks.merges().mergeExists("A1:A5"));

        // C3:D4 was entirely inside rows 3-4, it should be annihilated
        REQUIRE_FALSE(wks.merges().mergeExists("C3:D4"));
        // Ensure it didn't mutate into a 0-height or negative-height invalid merge
        REQUIRE_FALSE(wks.merges().mergeExists("C3:D2"));
        REQUIRE_FALSE(wks.merges().mergeExists("C3:D3"));

        // F4:G6 had its top half deleted and bottom half shifted up.
        // Original rows: 4, 5, 6. Deleted 3, 4.
        // Row 4 is gone. Row 5 becomes 3. Row 6 becomes 4.
        // It should collapse into F3:G4.
        REQUIRE(wks.merges().mergeExists("F3:G4"));
        REQUIRE_FALSE(wks.merges().mergeExists("F4:G6"));

        doc.close();
    }

    SECTION("insertRow updates mergeCells at/below insertion point")
    {
        XLDocument doc;
        doc.create("./testInsDelRow_mergeInsert.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("A1:B1");    // above insertion point — should not move
        wks.mergeCells("A3:B4");    // at/below insertion point — should shift down

        wks.insertRow(2, 1);    // insert 1 row before row 2

        REQUIRE(wks.merges().mergeExists("A1:B1"));    // unchanged
        REQUIRE(wks.merges().mergeExists("A4:B5"));    // shifted from A3:B4 to A4:B5

        doc.close();
    }

    // =============================================================================
    // Column Insert / Delete Tests
    // =============================================================================

    SECTION("deleteColumn basic — subsequent columns shift left")
    {
        XLDocument doc;
        doc.create("./testInsDelCol_basic.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value() = "ColA";
        wks.cell("B1").value() = "ColB";
        wks.cell("C1").value() = "ColC";
        wks.cell("D1").value() = "ColD";

        // Delete column B (col 2)
        wks.deleteColumn(2, 1);

        REQUIRE(wks.cell("A1").value().getString() == "ColA");
        REQUIRE(wks.cell("B1").value().getString() == "ColC");
        REQUIRE(wks.cell("C1").value().getString() == "ColD");
        REQUIRE(wks.cell("D1").value().type() == XLValueType::Empty);

        doc.close();
    }

    SECTION("insertColumn basic — existing columns shift right")
    {
        XLDocument doc;
        doc.create("./testInsDelCol_insert.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value() = "ColA";
        wks.cell("B1").value() = "ColB";
        wks.cell("C1").value() = "ColC";

        // Insert one empty column before column B (col 2)
        wks.insertColumn(2, 1);

        REQUIRE(wks.cell("A1").value().getString() == "ColA");
        REQUIRE(wks.cell("B1").value().type() == XLValueType::Empty);    // new empty column
        REQUIRE(wks.cell("C1").value().getString() == "ColB");
        REQUIRE(wks.cell("D1").value().getString() == "ColC");

        doc.close();
    }

    SECTION("deleteColumn updates formulas")
    {
        XLDocument doc;
        doc.create("./testInsDelCol_formula.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value()   = 1;
        wks.cell("B1").value()   = 2;
        wks.cell("C1").value()   = 3;
        wks.cell("D1").formula() = "=SUM(A1:C1)";

        // Delete column B — range should become A1:B1, D1→C1
        wks.deleteColumn(2, 1);

        std::string formulaAfter = wks.cell("C1").formula();    // D1 → C1
        REQUIRE(formulaAfter.find("B1") != std::string::npos);
        REQUIRE(formulaAfter.find("C1") == std::string::npos);

        doc.close();
    }
}

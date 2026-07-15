/**
 * testXLSlicerOOXMLFixes.cpp
 *
 * Regression tests for the three core Slicer OOXML bug fixes:
 *
 * FIX-1 (XLTables.cpp L130):
 *   autoFilter xr:uid last segment must be exactly 12 hex digits.
 *   Bad:  {00000000-0009-0000-0100-0000000001}  (10 digits)
 *   Good: {00000000-0009-0000-0100-000001000000} (12 digits)
 *
 * FIX-2 (XLDocument.cpp L1422):
 *   workbook.xml extLst node order when both Table and Pivot Slicers exist.
 *   Required: {BBE1A952}(Pivot) → {79F54976}(workbookPr) → {46BE6895}(Table)
 *   Bad:      {79F54976} → {46BE6895} → {BBE1A952}
 *
 * FIX-3 (XLDocument.cpp L1381-1383):
 *   Pivot SlicerCache XML must use "0"/"1" for OOXML boolean attributes,
 *   not "false"/"true".
 *   Bad:  showMissing="false"  s="true"
 *   Good: showMissing="0"      s="1"
 *
 * Run:
 *   ./OpenXLSXTests "[SlicerOOXMLFixes]" --reporter compact
 */

#include "TestHelpers.hpp"
#include "OpenXLSX.hpp"
#include "XLSlicer.hpp"
#include "XLSlicerCollection.hpp"
#include "XLTables.hpp"
#include "XLPivotTable.hpp"

#include <catch2/catch_test_macros.hpp>
#include <regex>
#include <string>
#include <vector>

using namespace OpenXLSX;

// ─── Helpers ────────────────────────────────────────────────────────────────

namespace {

// Fill a standard 4-column dataset into a worksheet (A1:D9)
void fillStandardData(XLWorksheet& wks)
{
    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Product";
    wks.cell("C1").value() = "Quarter";
    wks.cell("D1").value() = "Sales";

    const std::vector<std::tuple<std::string,std::string,std::string,int>> rows = {
        {"North","Widget","Q1",1200}, {"North","Gadget","Q2",900},
        {"South","Widget","Q1",850},  {"South","Gadget","Q3",1100},
        {"East", "Widget","Q2",700},  {"East", "Gadget","Q4",1300},
        {"West", "Widget","Q3",600},  {"West", "Gadget","Q1",1050},
    };
    for (size_t i = 0; i < rows.size(); ++i) {
        uint32_t r = 2 + static_cast<uint32_t>(i);
        wks.cell(r, 1).value() = std::get<0>(rows[i]);
        wks.cell(r, 2).value() = std::get<1>(rows[i]);
        wks.cell(r, 3).value() = std::get<2>(rows[i]);
        wks.cell(r, 4).value() = std::get<3>(rows[i]);
    }
}

// Validate the UUID format: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
// Returns true if the last segment is exactly 12 hex characters.
bool isValidUUID12(const std::string& uid)
{
    // Pattern: {8-4-4-4-12}
    static const std::regex uuidRe(
        R"(\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-([0-9A-Fa-f]{12})\})");
    std::smatch m;
    return std::regex_match(uid, m, uuidRe);
}

// Collect ext URIs in document order from a workbook extLst node
std::vector<std::string> getExtLstURIOrder(const XMLNode& extLst)
{
    std::vector<std::string> uris;
    for (auto ext = extLst.first_child(); ext; ext = ext.next_sibling()) {
        auto uriAttr = ext.attribute("uri");
        if (uriAttr)
            uris.push_back(uriAttr.value());
    }
    return uris;
}

} // namespace


// ════════════════════════════════════════════════════════════════════════════
// FIX-1: autoFilter xr:uid must have exactly 12 hex digits in last segment
// ════════════════════════════════════════════════════════════════════════════
TEST_CASE("SlicerFix1_AutoFilterUUIDFormat", "[SlicerOOXMLFixes]")
{
    const std::string fname = TestHelpers::getUniqueFilename("slicer_fix1_uuid") + ".xlsx";

    // --- Create: table with Table Slicer ---
    {
        XLDocument doc;
        doc.create(fname, XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        fillStandardData(wks);

        auto table = wks.tables().add("T_UUID", "A1:D9");
        XLSlicerOptions opts;
        opts.name = "Slicer_Region";
        wks.addTableSlicer("F1", table, "Region", opts);
        doc.save();
        doc.close();
    }

    // --- Verify: table1.xml autoFilter xr:uid last segment is 12 hex digits ---
    {
        XLDocument doc;
        doc.open(fname);
        auto wks = doc.workbook().worksheet("Sheet1");

        // Access the raw table XML
        auto& tables = wks.tables();
        REQUIRE(tables.count() >= 1);
        auto table = tables.table("T_UUID");
        REQUIRE(table.valid());

        // Read table XML directly through xmlDocument
        auto tableRoot = table.xmlDocument().document_element();
        auto autoFilter = tableRoot.child("autoFilter");
        REQUIRE_FALSE(autoFilter.empty());

        auto uidAttr = autoFilter.attribute("xr:uid");
        REQUIRE_FALSE(uidAttr.empty());

        std::string uid = uidAttr.value();
        INFO("autoFilter xr:uid = " << uid);

        // FIX-1 regression: last segment must be exactly 12 hex chars
        CHECK(isValidUUID12(uid));

        // Additionally verify the exact pattern we generate:
        // {00000000-0009-0000-0100-XXXXXX000000} where XXXXXX = tableId in hex
        CHECK(uid.length() == 38);  // total chars: {8-4-4-4-12} + 4 dashes + 2 braces = 38

        doc.close();
    }
}

TEST_CASE("SlicerFix1_MultipleTablesUniqueUIDs", "[SlicerOOXMLFixes]")
{
    // When multiple tables exist, each table's autoFilter xr:uid must be unique
    const std::string fname = TestHelpers::getUniqueFilename("slicer_fix1_multitable") + ".xlsx";

    {
        XLDocument doc;
        doc.create(fname, XLForceOverwrite);

        // Sheet1: table T1 + slicer
        auto wks1 = doc.workbook().worksheet("Sheet1");
        fillStandardData(wks1);
        auto t1 = wks1.tables().add("T1", "A1:D9");
        wks1.addTableSlicer("F1", t1, "Region",
                            XLSlicerOptions{"Slicer_Region", "Region"});

        // Sheet2: table T2 + slicer
        doc.workbook().addWorksheet("Sheet2");
        auto wks2 = doc.workbook().worksheet("Sheet2");
        fillStandardData(wks2);
        auto t2 = wks2.tables().add("T2", "A1:D9");
        wks2.addTableSlicer("F1", t2, "Product",
                            XLSlicerOptions{"Slicer_Product", "Product"});

        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(fname);

        auto wks1 = doc.workbook().worksheet("Sheet1");
        auto wks2 = doc.workbook().worksheet("Sheet2");

        auto root1 = wks1.tables().table("T1").xmlDocument().document_element();
        auto root2 = wks2.tables().table("T2").xmlDocument().document_element();

        auto uid1 = root1.child("autoFilter").attribute("xr:uid").value();
        auto uid2 = root2.child("autoFilter").attribute("xr:uid").value();

        INFO("T1 autoFilter uid = " << uid1);
        INFO("T2 autoFilter uid = " << uid2);

        // Both must be valid
        CHECK(isValidUUID12(uid1));
        CHECK(isValidUUID12(uid2));

        // Both must be different
        CHECK(uid1 != uid2);

        doc.close();
    }
}


// ════════════════════════════════════════════════════════════════════════════
// FIX-2: workbook.xml extLst order when Table + Pivot Slicers coexist
// Required order: {BBE1A952}(Pivot) → {79F54976}(workbookPr) → {46BE6895}(Table)
// ════════════════════════════════════════════════════════════════════════════
TEST_CASE("SlicerFix2_WorkbookExtLstOrder_TableOnly", "[SlicerOOXMLFixes]")
{
    // When only Table Slicers exist: {79F54976} → {46BE6895}
    const std::string fname = TestHelpers::getUniqueFilename("slicer_fix2_tableonly") + ".xlsx";

    {
        XLDocument doc;
        doc.create(fname, XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        fillStandardData(wks);
        auto table = wks.tables().add("T_TableOnly", "A1:D9");
        wks.addTableSlicer("F1", table, "Region",
                            XLSlicerOptions{"Slicer_Region", "Region"});
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(fname);
        auto wbkNode = doc.workbook().xmlDocument().document_element();
        auto extLst  = wbkNode.child("extLst");
        REQUIRE_FALSE(extLst.empty());

        auto uris = getExtLstURIOrder(extLst);

        // Must have workbookPr and Table cache URIs
        auto hasPivotCache = std::find(uris.begin(), uris.end(),
            "{BBE1A952-AA13-448e-AADC-164F8A28A991}") != uris.end();
        auto hasWbkPr = std::find(uris.begin(), uris.end(),
            "{79F54976-1DA5-4618-B147-4CDE4B953A38}") != uris.end();
        auto hasTableCache = std::find(uris.begin(), uris.end(),
            "{46BE6895-7355-4a93-B00E-2C351335B9C9}") != uris.end();

        // Table-only: no pivot cache URI
        CHECK_FALSE(hasPivotCache);
        // Must have workbookPr and table caches
        CHECK(hasWbkPr);
        CHECK(hasTableCache);

        // workbookPr must appear BEFORE table caches
        auto idxWbkPr = std::find(uris.begin(), uris.end(),
            "{79F54976-1DA5-4618-B147-4CDE4B953A38}") - uris.begin();
        auto idxTableCache = std::find(uris.begin(), uris.end(),
            "{46BE6895-7355-4a93-B00E-2C351335B9C9}") - uris.begin();
        INFO("workbookPr index=" << idxWbkPr << " tableCache index=" << idxTableCache);
        CHECK(idxWbkPr < idxTableCache);

        doc.close();
    }
}

TEST_CASE("SlicerFix2_WorkbookExtLstOrder_Mixed", "[SlicerOOXMLFixes]")
{
    // FIX-2 core: when BOTH Table and Pivot Slicers exist,
    // extLst order MUST be: {BBE1A952} → {79F54976} → {46BE6895}
    const std::string fname = TestHelpers::getUniqueFilename("slicer_fix2_mixed") + ".xlsx";

    {
        XLDocument doc;
        doc.create(fname, XLForceOverwrite);

        // Sheet1: Table + Table Slicers (created FIRST)
        auto wks1 = doc.workbook().worksheet("Sheet1");
        fillStandardData(wks1);
        auto t1 = wks1.tables().add("T_Mixed1", "A1:D9");
        wks1.addTableSlicer("F1", t1, "Region",
                            XLSlicerOptions{"Slicer_Region", "Region"});
        wks1.addTableSlicer("H1", t1, "Product",
                            XLSlicerOptions{"Slicer_Product", "Product"});

        // Sheet2: Pivot Table + Pivot Slicer (created SECOND)
        doc.workbook().addWorksheet("Sheet2");
        auto wks2 = doc.workbook().worksheet("Sheet2");
        fillStandardData(wks2);
        auto pt = wks2.addPivotTable(
            XLPivotTableOptions("PT_Mixed", "Sheet1!A1:D9", "F2")
                .addRowField("Region")
                .addDataField("Sales", "Total")
        );
        if (pt.valid()) {
            wks2.addPivotSlicer("K2", pt, "Region",
                                XLSlicerOptions{"Region", "Region Filter"});
        }

        // Sheet3: Another Table Slicer (created LAST)
        doc.workbook().addWorksheet("Sheet3");
        auto wks3 = doc.workbook().worksheet("Sheet3");
        fillStandardData(wks3);
        auto t3 = wks3.tables().add("T_Mixed3", "A1:D9");
        wks3.addTableSlicer("F1", t3, "Quarter",
                            XLSlicerOptions{"Slicer_Quarter", "Quarter"});

        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(fname);
        auto wbkNode = doc.workbook().xmlDocument().document_element();
        auto extLst  = wbkNode.child("extLst");
        REQUIRE_FALSE(extLst.empty());

        auto uris = getExtLstURIOrder(extLst);

        INFO("extLst URI order:");
        for (size_t i = 0; i < uris.size(); ++i)
            INFO("  [" << i << "] " << uris[i]);

        // All three must exist
        auto itPivot = std::find(uris.begin(), uris.end(),
            "{BBE1A952-AA13-448e-AADC-164F8A28A991}");
        auto itWbkPr = std::find(uris.begin(), uris.end(),
            "{79F54976-1DA5-4618-B147-4CDE4B953A38}");
        auto itTable = std::find(uris.begin(), uris.end(),
            "{46BE6895-7355-4a93-B00E-2C351335B9C9}");

        REQUIRE(itPivot != uris.end());
        REQUIRE(itWbkPr != uris.end());
        REQUIRE(itTable != uris.end());

        auto idxPivot = itPivot - uris.begin();
        auto idxWbkPr = itWbkPr - uris.begin();
        auto idxTable = itTable - uris.begin();

        // FIX-2: {BBE1A952} MUST come before {79F54976}
        CHECK(idxPivot < idxWbkPr);

        // {79F54976} MUST come before {46BE6895}
        CHECK(idxWbkPr < idxTable);

        // Combined: strict order Pivot < workbookPr < Table
        CHECK(idxPivot < idxTable);

        doc.close();
    }
}

TEST_CASE("SlicerFix2_WorkbookExtLstOrder_PivotThenTable", "[SlicerOOXMLFixes]")
{
    // Edge case: Pivot Slicer created BEFORE Table Slicer
    // (order should still be correct)
    const std::string fname = TestHelpers::getUniqueFilename("slicer_fix2_pivot_first") + ".xlsx";

    {
        XLDocument doc;
        doc.create(fname, XLForceOverwrite);

        // Sheet1: Pivot first
        auto wks1 = doc.workbook().worksheet("Sheet1");
        fillStandardData(wks1);
        auto pt = wks1.addPivotTable(
            XLPivotTableOptions("PT_First", "Sheet1!A1:D9", "F2")
                .addRowField("Region")
                .addDataField("Sales", "Total")
        );
        if (pt.valid()) {
            wks1.addPivotSlicer("K2", pt, "Region",
                                XLSlicerOptions{"PivotRegion", "Pivot Region"});
        }

        // Sheet2: Table second
        doc.workbook().addWorksheet("Sheet2");
        auto wks2 = doc.workbook().worksheet("Sheet2");
        fillStandardData(wks2);
        auto t2 = wks2.tables().add("T_Second", "A1:D9");
        wks2.addTableSlicer("F1", t2, "Product",
                            XLSlicerOptions{"Slicer_Product2", "Product"});

        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(fname);
        auto wbkNode = doc.workbook().xmlDocument().document_element();
        auto extLst  = wbkNode.child("extLst");
        REQUIRE_FALSE(extLst.empty());

        auto uris = getExtLstURIOrder(extLst);

        auto idxPivot = std::find(uris.begin(), uris.end(),
            "{BBE1A952-AA13-448e-AADC-164F8A28A991}") - uris.begin();
        auto idxWbkPr = std::find(uris.begin(), uris.end(),
            "{79F54976-1DA5-4618-B147-4CDE4B953A38}") - uris.begin();
        auto idxTable = std::find(uris.begin(), uris.end(),
            "{46BE6895-7355-4a93-B00E-2C351335B9C9}") - uris.begin();

        INFO("Pivot idx=" << idxPivot << " workbookPr idx=" << idxWbkPr
             << " Table idx=" << idxTable);

        CHECK(idxPivot < idxWbkPr);
        CHECK(idxWbkPr < idxTable);

        doc.close();
    }
}


// ════════════════════════════════════════════════════════════════════════════
// FIX-3: Pivot SlicerCache must use "0"/"1" not "false"/"true"
// ════════════════════════════════════════════════════════════════════════════
TEST_CASE("SlicerFix3_PivotCacheBooleanValues", "[SlicerOOXMLFixes]")
{
    const std::string fname = TestHelpers::getUniqueFilename("slicer_fix3_bool") + ".xlsx";

    {
        XLDocument doc;
        doc.create(fname, XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        fillStandardData(wks);

        auto pt = wks.addPivotTable(
            XLPivotTableOptions("PT_Bool", "Sheet1!A1:D9", "F2")
                .addRowField("Region")
                .addDataField("Sales", "Total")
        );
        REQUIRE(pt.valid());

        wks.addPivotSlicer("K2", pt, "Region",
                           XLSlicerOptions{"Region", "Region Filter"});
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(fname);

        // Locate the pivot slicer cache (contains pivotTables node, not extLst)
        bool foundPivotCache   = false;
        bool showMissingValid  = false;  // must be "0", not "false"
        bool noStringFalseTrueValues = true;

        for (const auto& entryName : doc.archive().entryNames()) {
            if (entryName.find("xl/slicerCaches/") == std::string::npos) continue;

            // Access the raw XML for this slicer cache
            std::string xmlStr = doc.extractXmlFromArchive(entryName);
            if (xmlStr.empty()) continue;

            XMLDocument xmlDoc;
            xmlDoc.load_string(xmlStr.c_str());
            auto root = xmlDoc.document_element();
            // Pivot slicer cache has <pivotTables> child
            if (root.child("pivotTables").empty()) continue;

            foundPivotCache = true;

            auto tabular = root.child("data").child("tabular");
            if (tabular.empty()) continue;

            // FIX-3a: showMissing must be "0" not "false"
            auto showMissingAttr = tabular.attribute("showMissing");
            if (!showMissingAttr.empty()) {
                std::string val = showMissingAttr.value();
                INFO("showMissing value = '" << val << "'");
                showMissingValid = (val == "0" || val == "1");
                CHECK(val != "false");
                CHECK(val != "true");
            }

            // FIX-3b: <i s=...> must use "0" or "1", not "false"/"true"
            auto items = tabular.child("items");
            for (auto item = items.first_child(); item; item = item.next_sibling()) {
                auto sAttr = item.attribute("s");
                if (!sAttr.empty()) {
                    std::string sVal = sAttr.value();
                    if (sVal == "false" || sVal == "true") {
                        noStringFalseTrueValues = false;
                        INFO("Item s attribute has invalid bool value: '" << sVal << "'");
                    }
                }
            }
            break;  // Only need to check one pivot slicer cache
        }

        // If we have a pivot slicer, the cache must exist and be correct
        REQUIRE(foundPivotCache);
        CHECK(showMissingValid);
        CHECK(noStringFalseTrueValues);

        doc.close();
    }
}

TEST_CASE("SlicerFix3_PivotCacheBooleanValues_MultipleSlicers", "[SlicerOOXMLFixes]")
{
    // Verify that ALL pivot slicer caches use correct boolean values
    const std::string fname = TestHelpers::getUniqueFilename("slicer_fix3_multipivot") + ".xlsx";

    {
        XLDocument doc;
        doc.create(fname, XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        fillStandardData(wks);

        auto pt = wks.addPivotTable(
            XLPivotTableOptions("PT_MultiBool", "Sheet1!A1:D9", "F2")
                .addRowField("Region")
                .addDataField("Sales", "Total")
        );
        REQUIRE(pt.valid());

        // Two pivot slicers on same table
        wks.addPivotSlicer("K2", pt, "Region",
                           XLSlicerOptions{"PivotRegion2", "Region"});
        wks.addPivotSlicer("M2", pt, "Product",
                           XLSlicerOptions{"PivotProduct2", "Product"});

        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(fname);

        int  pivotCacheCount   = 0;
        bool allBooleansValid  = true;

        for (const auto& entryName : doc.archive().entryNames()) {
            if (entryName.find("xl/slicerCaches/") == std::string::npos) continue;

            std::string xmlStr = doc.extractXmlFromArchive(entryName);
            if (xmlStr.empty()) continue;

            XMLDocument xmlDoc;
            xmlDoc.load_string(xmlStr.c_str());
            auto root = xmlDoc.document_element();
            if (root.child("pivotTables").empty()) continue;

            ++pivotCacheCount;

            auto tabular = root.child("data").child("tabular");
            if (tabular.empty()) continue;

            auto showMissingAttr = tabular.attribute("showMissing");
            if (!showMissingAttr.empty()) {
                std::string val = showMissingAttr.value();
                if (val == "false" || val == "true") {
                    allBooleansValid = false;
                    INFO("Cache " << entryName << ": showMissing='" << val << "' (must be 0 or 1)");
                }
            }

            auto items = tabular.child("items");
            for (auto item = items.first_child(); item; item = item.next_sibling()) {
                auto sAttr = item.attribute("s");
                if (!sAttr.empty()) {
                    std::string sVal = sAttr.value();
                    if (sVal == "false" || sVal == "true") {
                        allBooleansValid = false;
                        INFO("Cache " << entryName << ": item s='" << sVal << "' (must be 0 or 1)");
                    }
                }
            }
        }

        CHECK(pivotCacheCount >= 1);
        CHECK(allBooleansValid);

        doc.close();
    }
}


// ════════════════════════════════════════════════════════════════════════════
// COMBINED: All three fixes work together in one mixed workbook
// ════════════════════════════════════════════════════════════════════════════
TEST_CASE("SlicerFix_AllThreeFixes_CombinedRegression", "[SlicerOOXMLFixes]")
{
    // Creates the comprehensive mixed scenario that was originally failing:
    // - Sheet1: Table + 2 Table Slicers
    // - Sheet2: Pivot Table + Pivot Slicer
    // - Sheet3: Table + 1 Table Slicer
    // All three fixes must pass simultaneously.
    const std::string fname = TestHelpers::getUniqueFilename("slicer_fix_all3") + ".xlsx";

    {
        XLDocument doc;
        doc.create(fname, XLForceOverwrite);

        // Sheet1
        auto wks1 = doc.workbook().worksheet("Sheet1");
        fillStandardData(wks1);
        auto t1 = wks1.tables().add("T_Combined1", "A1:D9");
        wks1.addTableSlicer("F1", t1, "Region",
                            XLSlicerOptions{"Slicer_Region", "Region"});
        wks1.addTableSlicer("H1", t1, "Product",
                            XLSlicerOptions{"Slicer_Product", "Product"});

        // Sheet2
        doc.workbook().addWorksheet("Sheet2");
        auto wks2 = doc.workbook().worksheet("Sheet2");
        fillStandardData(wks2);
        auto pt = wks2.addPivotTable(
            XLPivotTableOptions("PT_Combined", "Sheet1!A1:D9", "F2")
                .addRowField("Region")
                .addDataField("Sales", "Total")
        );
        if (pt.valid()) {
            wks2.addPivotSlicer("K2", pt, "Region",
                                XLSlicerOptions{"Region1", "Region"});
        }

        // Sheet3
        doc.workbook().addWorksheet("Sheet3");
        auto wks3 = doc.workbook().worksheet("Sheet3");
        fillStandardData(wks3);
        auto t3 = wks3.tables().add("T_Combined3", "A1:D9");
        wks3.addTableSlicer("F1", t3, "Quarter",
                            XLSlicerOptions{"Slicer_Quarter", "Quarter"});

        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(fname);

        // ── FIX-1: Verify all table autoFilter UIDs are valid 12-digit UUIDs ──
        SECTION("FIX-1: autoFilter xr:uid format") {
            for (const auto& sheetName : {"Sheet1", "Sheet3"}) {
                auto wks = doc.workbook().worksheet(sheetName);
                auto tables = wks.tables();
                for (size_t i = 0; i < tables.count(); ++i) {
                    auto table = tables[i];
                    auto root = table.xmlDocument().document_element();
                    auto af   = root.child("autoFilter");
                    if (af.empty()) continue;

                    auto uid = std::string(af.attribute("xr:uid").value());
                    INFO("Sheet=" << sheetName << " table=" << table.name()
                         << " uid=" << uid);
                    CHECK(isValidUUID12(uid));
                    CHECK(uid.length() == 38);
                }
            }
        }

        // ── FIX-2: Verify workbook extLst node order ───────────────────────
        SECTION("FIX-2: workbook extLst order") {
            auto wbkNode = doc.workbook().xmlDocument().document_element();
            auto extLst  = wbkNode.child("extLst");
            REQUIRE_FALSE(extLst.empty());

            auto uris = getExtLstURIOrder(extLst);
            INFO("extLst order: [" << uris.size() << " nodes]");

            auto idxPivot = std::find(uris.begin(), uris.end(),
                "{BBE1A952-AA13-448e-AADC-164F8A28A991}") - uris.begin();
            auto idxWbkPr = std::find(uris.begin(), uris.end(),
                "{79F54976-1DA5-4618-B147-4CDE4B953A38}") - uris.begin();
            auto idxTable = std::find(uris.begin(), uris.end(),
                "{46BE6895-7355-4a93-B00E-2C351335B9C9}") - uris.begin();

            CHECK(static_cast<size_t>(idxPivot) < uris.size());
            CHECK(static_cast<size_t>(idxWbkPr) < uris.size());
            CHECK(static_cast<size_t>(idxTable) < uris.size());
            CHECK(idxPivot < idxWbkPr);
            CHECK(idxWbkPr < idxTable);
        }

        // ── FIX-3: Verify pivot slicer cache boolean values ────────────────
        SECTION("FIX-3: pivot cache boolean values") {
            bool foundPivotCache = false;

            for (const auto& entryName : doc.archive().entryNames()) {
                if (entryName.find("xl/slicerCaches/") == std::string::npos) continue;

                std::string xmlStr = doc.extractXmlFromArchive(entryName);
                if (xmlStr.empty()) continue;

                XMLDocument xmlDoc;
                xmlDoc.load_string(xmlStr.c_str());
                auto root = xmlDoc.document_element();
                if (root.child("pivotTables").empty()) continue;

                foundPivotCache = true;

                auto tabular = root.child("data").child("tabular");
                if (!tabular.empty()) {
                    auto showMissing = std::string(tabular.attribute("showMissing").value());
                    INFO("showMissing = '" << showMissing << "'");
                    CHECK(showMissing != "false");
                    CHECK(showMissing != "true");

                    for (auto item = tabular.child("items").first_child();
                         item; item = item.next_sibling()) {
                        auto sVal = std::string(item.attribute("s").value());
                        if (!sVal.empty()) {
                            CHECK(sVal != "false");
                            CHECK(sVal != "true");
                        }
                    }
                }
                break;
            }

            CHECK(foundPivotCache);
        }

        doc.close();
    }
}

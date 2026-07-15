#include "TestHelpers.hpp"
#include "OpenXLSX.hpp"
#include "XLPivotTable.hpp"

#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_testXLPivotTable_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotSlicerTest_SaveAs_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_testXLPivotTable_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotSlicerTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_testXLPivotTable_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__PivotTest_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("DynamicPivotTableGeneration", "[XLPivotTable]")
{
    // === Functionality Setup ===
    XLDocument doc;
    doc.create(__global_unique_testXLPivotTable_2(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Add some sample data
    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Product";
    wks.cell("C1").value() = "Sales";

    wks.cell("A2").value() = "North";
    wks.cell("B2").value() = "Apples";
    wks.cell("C2").value() = 100;

    wks.cell("A3").value() = "South";
    wks.cell("B3").value() = "Bananas";
    wks.cell("C3").value() = 300;

    wks.cell("A4").value() = "North";
    wks.cell("B4").value() = "Oranges";
    wks.cell("C4").value() = 150;

    auto options = XLPivotTableOptions("Pivot1", "Sheet1!A1:C4", "E1")
        .addRowField("Region")
        .addColumnField("Product")
        .addDataField("Sales", "Total Sales", XLPivotSubtotal::Sum);

    // Execute dynamic construction
    wks.addPivotTable(options);

    doc.save();
    doc.close();

    // === Read-back and OOXML Verification ===
    XLDocument doc2;
    doc2.open(__global_unique_testXLPivotTable_2());

    // Verify XML structure in Pivot Cache Definition
    std::string cacheDefXmlStr = doc2.extractXmlFromArchive("xl/pivotCache/pivotCacheDefinition1.xml");

    // Verify sharedItems has actual distinct values (containsString/Number present)
    bool hasCacheData = cacheDefXmlStr.find("containsString") != std::string::npos ||
                        cacheDefXmlStr.find("containsNumber") != std::string::npos;
    REQUIRE(hasCacheData);
    REQUIRE(cacheDefXmlStr.find("refreshOnLoad=\"1\"") != std::string::npos);

    // Verify sharedItems has actual distinct values for valid x-index references
    REQUIRE(cacheDefXmlStr.find("<cacheField name=\"Product\" numFmtId=\"0\">") != std::string::npos);
    REQUIRE(cacheDefXmlStr.find("<cacheField name=\"Sales\" numFmtId=\"0\">") != std::string::npos);
    // Check that sharedItems are populated (containsString or containsNumber)
    bool hasSharedData = cacheDefXmlStr.find("containsString") != std::string::npos ||
                         cacheDefXmlStr.find("containsNumber") != std::string::npos;
    REQUIRE(hasSharedData);

    // Verify XML structure in Pivot Table Definition
    std::string ptDefXmlStr = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");

    // Location and layout
    REQUIRE(ptDefXmlStr.find("<location ref=\"E1:") != std::string::npos);

    // Check axis mapping and specific element omission formatting
    REQUIRE(ptDefXmlStr.find("<pivotField axis=\"axisRow\"") != std::string::npos);    // Region
    REQUIRE(ptDefXmlStr.find("<pivotField axis=\"axisCol\"") != std::string::npos);    // Product
    bool hasDataField = ptDefXmlStr.find("<pivotField dataField=\"1\" showAll=\"0\" />") != std::string::npos ||
                        ptDefXmlStr.find("<pivotField dataField=\"1\" showAll=\"0\"/>") != std::string::npos;
    REQUIRE(hasDataField);    // Sales

    // Check proper field mappings
    REQUIRE(ptDefXmlStr.find("<rowFields count=\"1\">") != std::string::npos);
    bool hasRowField =
        ptDefXmlStr.find("<field x=\"0\" />") != std::string::npos || ptDefXmlStr.find("<field x=\"0\"/>") != std::string::npos;
    REQUIRE(hasRowField);

    REQUIRE(ptDefXmlStr.find("<colFields count=\"1\">") != std::string::npos);
    bool hasColField =
        ptDefXmlStr.find("<field x=\"1\" />") != std::string::npos || ptDefXmlStr.find("<field x=\"1\"/>") != std::string::npos;
    REQUIRE(hasColField);

    // Check colItems is populated (count > 0), exact count depends on data
    REQUIRE(ptDefXmlStr.find("<colItems count=") != std::string::npos);

    // Check custom data name overriding
    bool hasTotalSales = ptDefXmlStr.find("name=\"Total Sales\"") != std::string::npos;
    
    REQUIRE(hasTotalSales);

    doc2.close();
}

TEST_CASE("PivotTableAdvancedSlicersandRefreshOnLoad", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create(__global_unique_testXLPivotTable_1(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Sales";
    wks.cell("A2").value() = "North";
    wks.cell("B2").value() = 100;
    wks.cell("A3").value() = "South";
    wks.cell("B3").value() = 200;

    doc.workbook().addWorksheet("Pivot");
    auto pivotWks = doc.workbook().worksheet("Pivot");

    auto pt = pivotWks.addPivotTable(XLPivotTableOptions("MyPivot", "Sheet1!A1:B3", "A1")
                                     .addRowField("Region")
                                     .addDataField("Sales", "Sum of Sales", XLPivotSubtotal::Sum));

    // Set refresh on load
    pt.setRefreshOnLoad(true);

    // Add a Slicer bound to the Pivot Table
    XLSlicerOptions sOpts;
    sOpts.name    = "Region";
    sOpts.caption = "Region Filter";
    pivotWks.addPivotSlicer("D1", pt, "Region", sOpts);

    doc.save();
    doc.close();

    // Verify
    XLDocument doc2;
    doc2.open(__global_unique_testXLPivotTable_1());
    auto ptWks2 = doc2.workbook().worksheet("Pivot");
    REQUIRE(ptWks2.hasDrawing() == true);

    // Deep Validation 1: Worksheet extLst must have A8765BA9 for Pivot Slicers
    auto wksNode   = ptWks2.xmlDocument().document_element();
    auto wksExtLst = wksNode.child("extLst");
    REQUIRE(!wksExtLst.empty());
    auto wksExt = wksExtLst.find_child_by_attribute("uri", "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}");
    REQUIRE(!wksExt.empty());
    REQUIRE(!wksExt.child("x14:slicerList").empty());

    // Deep Validation 2: Workbook extLst must have BBE1A952 for Pivot Slicer Caches
    auto wbkNode   = doc2.workbook().xmlDocument().document_element();
    auto wbkExtLst = wbkNode.child("extLst");
    REQUIRE(!wbkExtLst.empty());
    auto wbkExt = wbkExtLst.find_child_by_attribute("uri", "{BBE1A952-AA13-448e-AADC-164F8A28A991}");
    REQUIRE(!wbkExt.empty());
    REQUIRE(!wbkExt.child("x14:slicerCaches").empty());

    // Verify both definedNames exist in the workbook (both widget name and cache name)
    auto definedNames = doc2.workbook().definedNames();
    REQUIRE(definedNames.exists("Region"));
    REQUIRE(definedNames.get("Region").refersTo() == "#N/A");
    REQUIRE(definedNames.exists("Slicer_Region"));
    REQUIRE(definedNames.get("Slicer_Region").refersTo() == "#N/A");

    // --- NEW: Deep Archive XML content format validation ---

    // 1. Validate PivotCacheDefinition contains the required x14:pivotCacheDefinition extLst (The root cause of the bug)
    auto pcNode = wbkNode.child("pivotCaches").child("pivotCache");
    REQUIRE(!pcNode.empty());
    std::string pcRId  = pcNode.attribute("r:id").value();
    std::string pcPath = doc2.workbookRelationships().relationshipById(pcRId).target();
    if (pcPath[0] != '/') pcPath = "/xl/" + pcPath;

    std::string pcDefStr = doc2.extractXmlFromArchive(pcPath.substr(1));    // drop leading slash
    REQUIRE(!pcDefStr.empty());
    XMLDocument pcDefDoc;
    pcDefDoc.load_string(pcDefStr.c_str());
    auto pcDefRoot = pcDefDoc.document_element();
    REQUIRE(std::string(pcDefRoot.name()) == "pivotCacheDefinition");

    auto pcDefExtLst = pcDefRoot.child("extLst");
    REQUIRE(!pcDefExtLst.empty());
    auto pcDefExt = pcDefExtLst.find_child_by_attribute("uri", "{725AE2AE-9491-48be-B2B4-4EB974FC3084}");
    REQUIRE(!pcDefExt.empty());
    REQUIRE(!pcDefExt.child("x14:pivotCacheDefinition").empty());
    REQUIRE(std::string(pcDefExt.child("x14:pivotCacheDefinition").attribute("pivotCacheId").value()) ==
            pcNode.attribute("cacheId").value());

    // 2. Validate SlicerCache XML formatting
    std::string slicerCacheStr = doc2.extractXmlFromArchive("xl/slicerCaches/slicerCache1.xml");
    REQUIRE(!slicerCacheStr.empty());
    XMLDocument scDoc;
    scDoc.load_string(slicerCacheStr.c_str());
    auto scRoot = scDoc.document_element();
    REQUIRE(std::string(scRoot.name()) == "slicerCacheDefinition");
    REQUIRE(std::string(scRoot.attribute("name").value()) == "Slicer_Region");
    REQUIRE(std::string(scRoot.attribute("sourceName").value()) == "Region");
    REQUIRE(std::string(scRoot.attribute("xr10:uid").value()) == "{00000000-0013-0000-FFFF-FFFF01000000}");
    auto scPivotTables = scRoot.child("pivotTables");
    REQUIRE(!scPivotTables.empty());
    REQUIRE(std::string(scPivotTables.child("pivotTable").attribute("name").value()) == "MyPivot");

    // 3. Validate Slicer XML formatting
    std::string slicerStr = doc2.extractXmlFromArchive("xl/slicers/slicer1.xml");
    REQUIRE(!slicerStr.empty());
    XMLDocument sDoc;
    sDoc.load_string(slicerStr.c_str());
    auto sRoot = sDoc.document_element();
    REQUIRE(std::string(sRoot.name()) == "slicers");
    auto slicerNode = sRoot.child("slicer");
    REQUIRE(!slicerNode.empty());
    REQUIRE(std::string(slicerNode.attribute("name").value()) == "Region");
    REQUIRE(std::string(slicerNode.attribute("cache").value()) == "Slicer_Region");
    REQUIRE(std::string(slicerNode.attribute("caption").value()) == "Region Filter");

    // 4. Validate Drawing XML formatting (Must use a14 for Pivot Slicers!)
    // We need to find the drawing file path for the Pivot worksheet
    std::string drwRId = ptWks2.xmlDocument().document_element().child("drawing").attribute("r:id").value();
    // We know it's one of the drawing files, let's just search the archive
    std::string drawingStr;
    for (auto name : doc2.archive().entryNames()) {
        if (name.find("xl/drawings/drawing") == 0) {
            std::string temp = doc2.extractXmlFromArchive(name);
            if (temp.find("RegionSlicer") != std::string::npos || temp.find("sle") != std::string::npos) {
                drawingStr = temp;
                break;
            }
        }
    }
    REQUIRE(!drawingStr.empty());
    XMLDocument drwDoc;
    drwDoc.load_string(drawingStr.c_str());
    auto drwRoot = drwDoc.document_element();

    bool foundSleChoice = false;
    for (auto anchor : drwRoot.children()) {
        std::string anchorName = anchor.raw_name();
        if (anchorName != "xdr:twoCellAnchor" && anchorName != "xdr:oneCellAnchor") {
            continue;
        }
        auto altContent = anchor.child("mc:AlternateContent");
        if (!altContent.empty()) {
            auto choice = altContent.child("mc:Choice");
            if (!choice.empty() && std::string(choice.attribute("Requires").value()) == "sle") {
                foundSleChoice    = true;
                auto graphicFrame = choice.child("xdr:graphicFrame");
                REQUIRE(!graphicFrame.empty());
                auto sleSlicer = graphicFrame.child("a:graphic").child("a:graphicData").child("sle:slicer");
                REQUIRE(!sleSlicer.empty());
                REQUIRE(std::string(sleSlicer.attribute("name").value()) == "Region");
            }
        }
    }
    REQUIRE(foundSleChoice == true);

    // Test double save for robustness
    doc2.saveAs(__global_unique_testXLPivotTable_0(), XLForceOverwrite);
    doc2.close();
    
    XLDocument doc3;
    doc3.open(__global_unique_testXLPivotTable_0());
    REQUIRE(doc3.workbook().worksheet("Pivot").hasDrawing() == true);
    
    // --- NEW: Test CRUD operations ---
    auto crudWks = doc3.workbook().worksheet("Pivot");
    
    // 1. Read: Verify the pivot table can be read and its properties matched
    auto ptList = crudWks.pivotTables();
    REQUIRE(ptList.size() == 1);
    REQUIRE(ptList[0].name() == "MyPivot");
    REQUIRE(ptList[0].targetCell() == "A1");
    REQUIRE(ptList[0].sourceRange() == "Sheet1!A1:B3");
    
    // 2. Update: Change source range and verify
    ptList[0].changeSourceRange("Sheet1!A1:B10");
    REQUIRE(ptList[0].sourceRange() == "Sheet1!A1:B10");
    
    // 3. Delete: Remove the pivot table (+ package part / orphan cache GC)
    bool deleted = crudWks.deletePivotTable("MyPivot");
    REQUIRE(deleted == true);
    REQUIRE(crudWks.pivotTables().empty() == true);
    // Orphan cache must be removed when last consumer is deleted.
    {
        auto cacheListAfter = doc3.workbook().xmlDocument().document_element().child("pivotCaches");
        REQUIRE(cacheListAfter.empty());
    }
    REQUIRE_FALSE(doc3.archive().hasEntry("xl/pivotTables/pivotTable1.xml"));

    // Deleting non-existent
    REQUIRE(crudWks.deletePivotTable("NonExistent") == false);

    // --- Pivot cache sharing (two tables, one cache) ---
    auto pt2 = crudWks.addPivotTable(XLPivotTableOptions("AnotherPivot", "Sheet1!A1:B3", "D1")
                                     .addRowField("Region")
                                     .addDataField("Sales"));
    auto pt3 = crudWks.addPivotTable(XLPivotTableOptions("SharedCachePivot", "Sheet1!A1:B3", "G1")
                                     .addRowField("Region")
                                     .addDataField("Sales"));
    (void)pt2;
    (void)pt3;

    auto cacheList = doc3.workbook().xmlDocument().document_element().child("pivotCaches");
    int cacheCount = 0;
    for (auto c : cacheList.children("pivotCache")) { cacheCount++; }
    REQUIRE(cacheCount == 1); // Same resolved source → one cache

    // Deleting one pivot must keep the shared cache.
    REQUIRE(crudWks.deletePivotTable("AnotherPivot") == true);
    cacheCount = 0;
    for (auto c : cacheList.children("pivotCache")) { cacheCount++; }
    REQUIRE(cacheCount == 1);
    REQUIRE(crudWks.pivotTables().size() == 1);

    // Deleting the last pivot drops the cache.
    REQUIRE(crudWks.deletePivotTable("SharedCachePivot") == true);
    REQUIRE(doc3.workbook().xmlDocument().document_element().child("pivotCaches").empty());
    doc3.close();
}
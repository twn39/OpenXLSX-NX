#include "OpenXLSX.hpp"
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("DynamicPivotTableGeneration", "[XLPivotTable]")
{
    // === Functionality Setup ===
    XLDocument doc;
    doc.create("./PivotTest.xlsx", XLForceOverwrite);
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
    doc2.open("./PivotTest.xlsx");

    // Verify XML structure in Pivot Cache Definition
    std::string cacheDefXmlStr = doc2.extractXmlFromArchive("xl/pivotCache/pivotCacheDefinition1.xml");

    // Check root node attributes indicating dynamic refresh
    REQUIRE(cacheDefXmlStr.find("saveData=\"0\"") != std::string::npos);
    REQUIRE(cacheDefXmlStr.find("refreshOnLoad=\"1\"") != std::string::npos);

    // Check dynamically analyzed fields from worksheet headers
    REQUIRE(cacheDefXmlStr.find("<cacheFields count=\"3\">") != std::string::npos);
    REQUIRE(cacheDefXmlStr.find("<cacheField name=\"Region\" numFmtId=\"0\">") != std::string::npos);
    REQUIRE(cacheDefXmlStr.find("<cacheField name=\"Product\" numFmtId=\"0\">") != std::string::npos);
    REQUIRE(cacheDefXmlStr.find("<cacheField name=\"Sales\" numFmtId=\"0\">") != std::string::npos);

    // Check proper missing item placeholders to avoid Excel corruption
    bool hasMissingItems = cacheDefXmlStr.find("<sharedItems containsBlank=\"1\" count=\"0\"><m/></sharedItems>") != std::string::npos ||
                           cacheDefXmlStr.find("<m />") != std::string::npos || cacheDefXmlStr.find("<m/>") != std::string::npos;
    REQUIRE(hasMissingItems);

    // Verify XML structure in Pivot Table Definition
    std::string ptDefXmlStr = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");

    // Location and layout
    REQUIRE(ptDefXmlStr.find("<location ref=\"E1\"") != std::string::npos);

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

    // Check correct empty items node handling (fixes Excel warning)
    REQUIRE(ptDefXmlStr.find("<colItems count=\"1\">") != std::string::npos);
    bool hasEmptyItem = ptDefXmlStr.find("<i />") != std::string::npos || ptDefXmlStr.find("<i/>") != std::string::npos;
    REQUIRE(hasEmptyItem);

    // Check custom data name overriding
    bool hasTotalSales = ptDefXmlStr.find("name=\"Total Sales\"") != std::string::npos;
    
    REQUIRE(hasTotalSales);

    doc2.close();
}

TEST_CASE("PivotTableAdvancedSlicersandRefreshOnLoad", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create("./PivotSlicerTest.xlsx", XLForceOverwrite);
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

    // Verify
    XLDocument doc2;
    doc2.open("./PivotSlicerTest.xlsx");
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

    // --- NEW: Deep Archive XML content format validation ---

    // 1. Validate PivotCacheDefinition contains the required x14:pivotCacheDefinition extLst (The root cause of the bug)
    auto pcNode = wbkNode.child("pivotCaches").child("pivotCache");
    REQUIRE(!pcNode.empty());
    std::string pcRId  = pcNode.attribute("r:id").value();
    std::string pcPath = doc2.workbookRelationships().relationshipById(pcRId).target();
    if (pcPath[0] != '/') pcPath = "/xl/" + pcPath;

    std::string pcDefStr = doc2.archive().getEntry(pcPath.substr(1));    // drop leading slash
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
    std::string slicerCacheStr = doc2.archive().getEntry("xl/slicerCaches/slicerCache1.xml");
    REQUIRE(!slicerCacheStr.empty());
    XMLDocument scDoc;
    scDoc.load_string(slicerCacheStr.c_str());
    auto scRoot = scDoc.document_element();
    REQUIRE(std::string(scRoot.name()) == "slicerCacheDefinition");
    REQUIRE(std::string(scRoot.attribute("name").value()) == "Slicer_Region");
    REQUIRE(std::string(scRoot.attribute("sourceName").value()) == "Region");
    auto scPivotTables = scRoot.child("pivotTables");
    REQUIRE(!scPivotTables.empty());
    REQUIRE(std::string(scPivotTables.child("pivotTable").attribute("name").value()) == "MyPivot");

    // 3. Validate Slicer XML formatting
    std::string slicerStr = doc2.archive().getEntry("xl/slicers/slicer1.xml");
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
            std::string temp = doc2.archive().getEntry(name);
            if (temp.find("RegionSlicer") != std::string::npos || temp.find("a14") != std::string::npos) {
                drawingStr = temp;
                break;
            }
        }
    }
    REQUIRE(!drawingStr.empty());
    XMLDocument drwDoc;
    drwDoc.load_string(drawingStr.c_str());
    auto drwRoot = drwDoc.document_element();

    bool foundA14Choice = false;
    for (auto anchor : drwRoot.children("xdr:twoCellAnchor")) {
        auto altContent = anchor.child("mc:AlternateContent");
        if (!altContent.empty()) {
            auto choice = altContent.child("mc:Choice");
            if (!choice.empty() && std::string(choice.attribute("Requires").value()) == "a14") {
                foundA14Choice    = true;
                auto graphicFrame = choice.child("xdr:graphicFrame");
                REQUIRE(!graphicFrame.empty());
                auto sleSlicer = graphicFrame.child("a:graphic").child("a:graphicData").child("sle:slicer");
                REQUIRE(!sleSlicer.empty());
                REQUIRE(std::string(sleSlicer.attribute("name").value()) == "Region");
            }
        }
    }
    REQUIRE(foundA14Choice == true);

    // Test double save for robustness
    doc2.saveAs("./PivotSlicerTest_SaveAs.xlsx", XLForceOverwrite);
    XLDocument doc3;
    doc3.open("./PivotSlicerTest_SaveAs.xlsx");
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
    
    // 3. Delete: Remove the pivot table
    bool deleted = crudWks.deletePivotTable("MyPivot");
    REQUIRE(deleted == true);
    REQUIRE(crudWks.pivotTables().empty() == true);
    
    // Deleting non-existent
    REQUIRE(crudWks.deletePivotTable("NonExistent") == false);

    // --- NEW: Test Pivot Cache Sharing ---
    auto pt2 = crudWks.addPivotTable(XLPivotTableOptions("AnotherPivot", "Sheet1!A1:B10", "D1")
                                     .addRowField("Region"));

    auto cacheList = doc3.workbook().xmlDocument().document_element().child("pivotCaches");
    int cacheCount = 0;
    for (auto c : cacheList.children("pivotCache")) { cacheCount++; }
    REQUIRE(cacheCount == 1); // Should reuse existing cache since range matches
}
#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <fstream>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Shapes and Form Controls functionality and OOXML verification") {
    XLDocument doc;
    doc.create("TestShapes.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Add a simple vector shape
    XLVectorShapeOptions opts1;
    opts1.type = XLVectorShapeType::Ellipse;
    opts1.name = "My Ellipse";
    opts1.text = "Hello Vector";
    opts1.fillColor = "FF0000"; // Red
    opts1.width = 150;
    opts1.height = 80;
    opts1.offsetX = 10;
    opts1.offsetY = 10;
    wks.addShape("B2", opts1);

    // Add a button (Form Control)
    XLVectorShapeOptions opts2;
    opts2.type = XLVectorShapeType::Rectangle;
    opts2.name = "Button 1";
    opts2.text = "Click Me";
    opts2.fillColor = "E0E0E0"; 
    opts2.lineColor = "808080";
    opts2.width = 100;
    opts2.height = 30;
    opts2.macro = "Button1_Click";
    wks.addShape("D5", opts2);

    doc.save();

    // 1. Functionality Validation (Reload the document)
    XLDocument doc2;
    doc2.open("TestShapes.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");
    
    // Check if the drawing relationship is correctly created
    REQUIRE(wks2.hasDrawing() == true);

    auto& drawing = wks2.drawing();
    auto rootNode = drawing.xmlDocument().document_element();
    
    // Count xdr:oneCellAnchor elements
    int shapeCount = 0;
    int buttonCount = 0;
    int ellipseCount = 0;

    for (const auto& anchor : rootNode.children("xdr:oneCellAnchor")) {
        shapeCount++;
        auto sp = anchor.child("xdr:sp");
        if (!sp.empty()) {
            // Verify Macro
            if (std::string(sp.attribute("macro").value()) == "Button1_Click") {
                buttonCount++;
            }
            
            // Verify Preset Geometry (Ellipse)
            auto prstGeom = sp.child("xdr:spPr").child("a:prstGeom");
            if (!prstGeom.empty() && std::string(prstGeom.attribute("prst").value()) == "ellipse") {
                ellipseCount++;
            }
        }
    }

    REQUIRE(shapeCount == 2);
    REQUIRE(buttonCount == 1);
    REQUIRE(ellipseCount == 1);

    // 2. Deep OOXML Validation
    // Validate that the XML tags strictly conform to the DrawingML specification we implemented
    bool hasSpPr = false;
    bool hasTxBody = false;

    for (const auto& anchor : rootNode.children("xdr:oneCellAnchor")) {
        auto sp = anchor.child("xdr:sp");
        REQUIRE(!sp.empty()); // Every inserted vector should be an sp node
        
        auto spPr = sp.child("xdr:spPr");
        if (!spPr.empty()) {
            hasSpPr = true;
            // Validate bounding box properties
            auto xfrm = spPr.child("a:xfrm");
            REQUIRE(!xfrm.empty());
            REQUIRE(!xfrm.child("a:ext").empty());
        }

        auto txBody = sp.child("xdr:txBody");
        if (!txBody.empty()) {
            hasTxBody = true;
            // Validate valid rich text structure
            auto bodyPr = txBody.child("a:bodyPr");
            REQUIRE(!bodyPr.empty());
            auto p = txBody.child("a:p");
            REQUIRE(!p.empty());
            auto r = p.child("a:r");
            REQUIRE(!r.empty());
            auto t = r.child("a:t");
            REQUIRE(!t.empty());
            // Make sure the text matches what we set
            std::string text = t.text().get();
            REQUIRE((text == "Hello Vector" || text == "Click Me"));
        }
    }

    REQUIRE(hasSpPr == true);
    REQUIRE(hasTxBody == true);
}

#include "OpenXLSX.hpp"
#include <catch2/catch_test_macros.hpp>
#include <fstream>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Shapes and Form Controls functionality and OOXML verification")
{
    XLDocument doc;
    doc.create("TestShapes.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Add a simple vector shape
    XLVectorShapeOptions opts1;
    opts1.type      = XLVectorShapeType::Ellipse;
    opts1.name      = "My Ellipse";
    opts1.text      = "Hello Vector";
    opts1.fillColor = "FF0000";    // Red
    opts1.width     = 150;
    opts1.height    = 80;
    opts1.offsetX   = 10;
    opts1.offsetY   = 10;
    wks.addShape("B2", opts1);

    // Add a button (Form Control)
    XLVectorShapeOptions opts2;
    opts2.type      = XLVectorShapeType::Rectangle;
    opts2.name      = "Button 1";
    opts2.text      = "Click Me";
    opts2.fillColor = "E0E0E0";
    opts2.lineColor = "808080";
    opts2.width     = 100;
    opts2.height    = 30;
    opts2.macro     = "Button1_Click";
    wks.addShape("D5", opts2);

    doc.save();

    // 1. Functionality Validation (Reload the document)
    XLDocument doc2;
    doc2.open("TestShapes.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    // Check if the drawing relationship is correctly created
    REQUIRE(wks2.hasDrawing() == true);

    auto& drawing  = wks2.drawing();
    auto  rootNode = drawing.xmlDocument().document_element();

    // Count xdr:oneCellAnchor elements
    int shapeCount   = 0;
    int buttonCount  = 0;
    int ellipseCount = 0;

    for (const auto& anchor : rootNode.children("xdr:oneCellAnchor")) {
        shapeCount++;
        auto sp = anchor.child("xdr:sp");
        if (!sp.empty()) {
            // Verify Macro
            if (std::string(sp.attribute("macro").value()) == "Button1_Click") { buttonCount++; }

            // Verify Preset Geometry (Ellipse)
            auto prstGeom = sp.child("xdr:spPr").child("a:prstGeom");
            if (!prstGeom.empty() && std::string(prstGeom.attribute("prst").value()) == "ellipse") { ellipseCount++; }
        }
    }

    REQUIRE(shapeCount == 2);
    REQUIRE(buttonCount == 1);
    REQUIRE(ellipseCount == 1);

    // 2. Deep OOXML Validation
    // Validate that the XML tags strictly conform to the DrawingML specification we implemented
    bool hasSpPr   = false;
    bool hasTxBody = false;

    for (const auto& anchor : rootNode.children("xdr:oneCellAnchor")) {
        auto sp = anchor.child("xdr:sp");
        REQUIRE(!sp.empty());    // Every inserted vector should be an sp node

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

TEST_CASE("Vector Shape Enhancements Validation")
{
    XLDocument doc;
    doc.create("TestShapeEnhancements.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Shape 1: TwoCellAnchor with Advanced Types (FlowChartInputOutput) and ARGB RichText
    XLVectorShapeOptions opts1;
    opts1.type = XLVectorShapeType::FlowChartData; // Will output as flowChartInputOutput internally
    opts1.endRow = 5;
    opts1.endCol = 5;
    opts1.endOffsetX = 10;
    opts1.endOffsetY = 10;
    
    // Pass 8-character ARGB to ensure stripping mechanism drops the alpha channel down to 6-char RGB.
    XLRichText rt("Data: ");
    rt += XLRichTextRun("1234").setFontColor(XLColor(128, 255, 0, 0)); // alpha=128 (80 in hex)
    opts1.richText = rt;
    
    opts1.vertAlign = "ctr";
    opts1.horzAlign = "r";

    wks.addShape("A1", opts1);

    // Shape 2: Transformations and Outline properties
    XLVectorShapeOptions opts2;
    opts2.type = XLVectorShapeType::Star24;
    opts2.rotation = 90;
    opts2.flipH = true;
    opts2.flipV = true;
    opts2.lineDash = "dashDot";
    opts2.arrowStart = "diamond";
    opts2.arrowEnd = "triangle";

    wks.addShape("B2", opts2);

    doc.save();

    // 1. Functionality Validation (Reload the document)
    XLDocument doc2;
    doc2.open("TestShapeEnhancements.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    REQUIRE(wks2.hasDrawing() == true);
    auto& drawing  = wks2.drawing();
    auto  rootNode = drawing.xmlDocument().document_element();

    // Verification logic
    bool foundTwoCellAnchor = false;
    bool foundValidFlowchartData = false;
    bool foundValidRGBStrip = false;
    bool foundValidAlignment = false;

    bool foundOneCellAnchor = false;
    bool foundTransforms = false;
    bool foundDashes = false;

    for (const auto& anchor : rootNode.children()) {
        std::string anchorName = anchor.name();
        
        if (anchorName == "xdr:twoCellAnchor") {
            foundTwoCellAnchor = true;
            
            auto sp = anchor.child("xdr:sp");
            auto prstGeom = sp.child("xdr:spPr").child("a:prstGeom");
            if (std::string(prstGeom.attribute("prst").value()) == "flowChartInputOutput") {
                foundValidFlowchartData = true;
            }

            auto txBody = sp.child("xdr:txBody");
            auto bodyPr = txBody.child("a:bodyPr");
            if (std::string(bodyPr.attribute("anchor").value()) == "ctr") {
                auto p = txBody.child("a:p");
                if (std::string(p.child("a:pPr").attribute("algn").value()) == "r") {
                    foundValidAlignment = true;
                }
                
                // Verify the 2nd run's ARGB stripping. (128 = 80 in hex. Original was 80FF0000, must be FF0000)
                auto rNodes = p.children("a:r");
                auto it = rNodes.begin();
                std::advance(it, 1); // Get second run ("1234")
                if (it != rNodes.end()) {
                    auto srgbClr = it->child("a:rPr").child("a:solidFill").child("a:srgbClr");
                    if (std::string(srgbClr.attribute("val").value()) == "FF0000") {
                        foundValidRGBStrip = true;
                    }
                }
            }
        } 
        else if (anchorName == "xdr:oneCellAnchor") {
            foundOneCellAnchor = true;
            
            auto sp = anchor.child("xdr:sp");
            auto spPr = sp.child("xdr:spPr");
            auto xfrm = spPr.child("a:xfrm");

            if (std::string(xfrm.attribute("rot").value()) == "5400000" && 
                std::string(xfrm.attribute("flipH").value()) == "1" &&
                std::string(xfrm.attribute("flipV").value()) == "1") {
                foundTransforms = true;
            }

            auto ln = spPr.child("a:ln");
            if (std::string(ln.child("a:prstDash").attribute("val").value()) == "dashDot" &&
                std::string(ln.child("a:headEnd").attribute("type").value()) == "diamond" &&
                std::string(ln.child("a:tailEnd").attribute("type").value()) == "triangle") {
                foundDashes = true;
            }
        }
    }

    REQUIRE(foundTwoCellAnchor == true);
    REQUIRE(foundValidFlowchartData == true);
    REQUIRE(foundValidRGBStrip == true);
    REQUIRE(foundValidAlignment == true);
    
    REQUIRE(foundOneCellAnchor == true);
    REQUIRE(foundTransforms == true);
    REQUIRE(foundDashes == true);
}

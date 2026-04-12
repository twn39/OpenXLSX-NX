#include "OpenXLSX.hpp"
#include <catch2/catch_all.hpp>
#include <filesystem>
#include <iostream>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("ImageInsertAdvancedAPITests", "[XLImageInsert]")
{
    XLDocument doc;
    doc.create("testXLImageInsert.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    SECTION("OneCell Anchor (Default)")
    {
        // Should anchor to top-left of B2, auto-sizing to natural image dimensions.
        REQUIRE_NOTHROW(wks.insertImage("B2", "Tests/test.png"));

        doc.save();
        doc.close();

        XLDocument doc2;
        REQUIRE_NOTHROW(doc2.open("testXLImageInsert.xlsx"));
        auto wks2   = doc2.workbook().worksheet("Sheet1");
        auto images = wks2.images();
        REQUIRE(images.size() == 1);

        // Verify underlying XML
        std::string drawingStr = doc2.extractXmlFromArchive("xl/drawings/drawing1.xml");
        REQUIRE(drawingStr.find("<xdr:oneCellAnchor>") != std::string::npos);
        // B2 is col=1, row=1
        REQUIRE(drawingStr.find("<xdr:col>1</xdr:col>") != std::string::npos);
        REQUIRE(drawingStr.find("<xdr:row>1</xdr:row>") != std::string::npos);

        bool hasPng = false;
        for (auto name : doc2.archive().entryNames()) {
            if (name.find("xl/media/image") != std::string::npos && name.find(".png") != std::string::npos) { hasPng = true; }
        }
        REQUIRE(hasPng);
        
        // Test Image Extraction
        auto images2 = wks2.images();
        REQUIRE(images2.size() == 1);
        auto imgBinary = images2[0].imageBinary();
        REQUIRE(imgBinary.size() > 0);
        // Save to disk to verify 
        std::ofstream out("extracted_test.png", std::ios::binary);
        out.write(reinterpret_cast<const char*>(imgBinary.data()), imgBinary.size());
        out.close();
        REQUIRE(std::filesystem::exists("extracted_test.png"));
        REQUIRE(std::filesystem::file_size("extracted_test.png") == imgBinary.size());
        
        doc2.close();
    }

    SECTION("TwoCell Anchor with Scaling")
    {
        XLImageOptions opts;
        opts.positioning = XLImagePositioning::TwoCell;
        opts.scaleX      = 2.0;
        opts.scaleY      = 2.0;
        REQUIRE_NOTHROW(wks.insertImage("C5", "Tests/test.jpg", opts));

        doc.save();
        doc.close();

        XLDocument doc2;
        REQUIRE_NOTHROW(doc2.open("testXLImageInsert.xlsx"));

        std::string drawingStr = doc2.extractXmlFromArchive("xl/drawings/drawing1.xml");
        REQUIRE(drawingStr.find("<xdr:twoCellAnchor>") != std::string::npos);
        // C5 is col=2, row=4 (0-based in drawing XML)
        REQUIRE(drawingStr.find("<xdr:col>2</xdr:col>") != std::string::npos);
        REQUIRE(drawingStr.find("<xdr:row>4</xdr:row>") != std::string::npos);

        bool hasJpg = false;
        for (auto name : doc2.archive().entryNames()) {
            if (name.find("xl/media/image") != std::string::npos &&
                (name.find(".jpg") != std::string::npos || name.find(".jpeg") != std::string::npos))
            {
                hasJpg = true;
            }
        }
        REQUIRE(hasJpg);
        doc2.close();
    }

    SECTION("Absolute Anchor")
    {
        XLImageOptions opts;
        opts.positioning = XLImagePositioning::Absolute;
        opts.offsetX     = 50;    // pixels
        opts.offsetY     = 60;    // pixels
        REQUIRE_NOTHROW(wks.insertImage("A1", "Tests/test.png", opts));

        doc.save();
        doc.close();

        XLDocument doc2;
        REQUIRE_NOTHROW(doc2.open("testXLImageInsert.xlsx"));
        std::string drawingStr = doc2.extractXmlFromArchive("xl/drawings/drawing1.xml");
        REQUIRE(drawingStr.find("<xdr:absoluteAnchor>") != std::string::npos);
        // 50 * 9525 = 476250 EMUs
        // 60 * 9525 = 571500 EMUs
        bool hasOffset = (drawingStr.find("x=\"476250\" y=\"571500\"") != std::string::npos) ||
                         (drawingStr.find("<xdr:pos x=\"476250\" y=\"571500\"") != std::string::npos);
        REQUIRE(hasOffset);
        doc2.close();
    }

    std::filesystem::remove("testXLImageInsert.xlsx");
}
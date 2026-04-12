#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <iostream>
#include <sstream>

using namespace OpenXLSX;

TEST_CASE("XLDrawingVMLandXLTablesTests", "[Drawing][Tables]")
{
    const std::string filename = "DrawingTablesIntegration.xlsx";

    SECTION("XLVmlDrawing Shape Round-trip")
    {
        // 1. Create a document and add a comment which triggers VML drawing
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            wks.comments().addAuthor("Author");
            wks.comments().set("A1", "Test Comment", 0);
            REQUIRE(wks.hasVmlDrawing());

            doc.save();
            doc.close();
        }

        // 2. Open again to modify VML properties
        {
            XLDocument doc;
            doc.open(filename);
            auto  wks = doc.workbook().worksheet("Sheet1");
            auto& vml = wks.vmlDrawing();
            REQUIRE(vml.shapeCount() >= 1);

            auto shape = vml.shape(0);
            shape.setFillColor("#ff0000");    // Red

            auto style = shape.style();
            style.setWidth(200);
            style.setHeight(100);
            shape.setStyle(style);

            auto clientData = shape.clientData();
            clientData.setTextVAlign(XLShapeTextVAlign::Center);
            clientData.setTextHAlign(XLShapeTextHAlign::Center);

            doc.save();
            doc.close();
        }

        // 3. Re-open and verify final state
        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            REQUIRE(wks.hasVmlDrawing());
            REQUIRE(wks.comments().get("A1") == "Test Comment");

            auto& vml = wks.vmlDrawing();
            REQUIRE(vml.shapeCount() >= 1);

            auto shape = vml.shape(0);
            REQUIRE(shape.fillColor() == "#ff0000");

            auto style = shape.style();
            REQUIRE(style.width() == 200);
            REQUIRE(style.height() == 100);

            auto clientData = shape.clientData();
            REQUIRE(clientData.textVAlign() == XLShapeTextVAlign::Center);
            REQUIRE(clientData.textHAlign() == XLShapeTextHAlign::Center);

            doc.close();
        }
    }

    SECTION("XLTables Basics")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // Trigger table creation
            auto& tables = wks.tables();
            tables.add("MyTable", "A1:B2");
            REQUIRE(tables.valid());

            doc.save();
            doc.close();
        }

        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            REQUIRE(wks.hasTables());
            doc.close();
        }
    }
}

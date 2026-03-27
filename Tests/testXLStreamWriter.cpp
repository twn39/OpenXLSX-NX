#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <pugixml.hpp>
#include <string>

using namespace OpenXLSX;

TEST_CASE("Stream Writer Functional and OOXML Tests", "[XLStreamWriter]")
{
    SECTION("Functional and Type Mapping")
    {
        XLDocument doc;
        doc.create("./testXLStreamWriter_Functional.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        {
            auto writer = wks.streamWriter();
            
            // Header
            writer.appendRow({"ID", "Name", "Score", "Active", "Remarks (<, > & \")"});

            // Basic Data
            writer.appendRow({1, "Alice", 95.5, true, "Excellent"});

            // Special XML Escaped Chars
            writer.appendRow({2, "Bob & Charlie", 88.0, false, "<tag> \"quote\" 'single'"});

            // Empty Cells
            writer.appendRow({3, XLCellValue(), XLCellValue(), true, "Empty Test"});
            
            // Long Text
            std::string longText(1000, 'A');
            writer.appendRow({4, "LongText", 100.0, true, longText});
        }

        doc.save();
        doc.close();

        // Reopen and Verify Functionality
        XLDocument doc2;
        doc2.open("./testXLStreamWriter_Functional.xlsx");
        auto wks2 = doc2.workbook().worksheet("Sheet1");

        // Row 1
        REQUIRE(wks2.cell("A1").value().get<std::string>() == "ID");
        REQUIRE(wks2.cell("B1").value().get<std::string>() == "Name");
        REQUIRE(wks2.cell("E1").value().get<std::string>() == "Remarks (<, > & \")");

        // Row 2
        REQUIRE(wks2.cell("A2").value().get<int>() == 1);
        REQUIRE(wks2.cell("B2").value().get<std::string>() == "Alice");
        REQUIRE(wks2.cell("C2").value().get<double>() == 95.5);
        REQUIRE(wks2.cell("D2").value().get<bool>() == true);

        // Row 3 (Escaping)
        REQUIRE(wks2.cell("A3").value().get<int>() == 2);
        REQUIRE(wks2.cell("B3").value().get<std::string>() == "Bob & Charlie");
        REQUIRE(wks2.cell("E3").value().get<std::string>() == "<tag> \"quote\" 'single'");

        // Row 4 (Empty)
        REQUIRE(wks2.cell("A4").value().get<int>() == 3);
        REQUIRE(wks2.cell("B4").value().type() == XLValueType::Empty);
        REQUIRE(wks2.cell("C4").value().type() == XLValueType::Empty);
        REQUIRE(wks2.cell("E4").value().get<std::string>() == "Empty Test");
        
        // Row 5 (Long Text)
        REQUIRE(wks2.cell("E5").value().get<std::string>() == std::string(1000, 'A'));
        
        doc2.close();
    }

    SECTION("OOXML Content Validation")
    {
        XLDocument doc;
        doc.create("./testXLStreamWriter_OOXML.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        {
            auto writer = wks.streamWriter();
            writer.appendRow({"StringData", 42, true});
        }
        
        doc.save();
        doc.close();

        // Validate OOXML structure
        XLZipArchive archive;
        archive.open("./testXLStreamWriter_OOXML.xlsx");
        std::string sheetData = archive.getEntry("xl/worksheets/sheet1.xml");
        archive.close();

        pugi::xml_document xmlDoc;
        xmlDoc.load_string(sheetData.c_str());

        auto root = xmlDoc.document_element();
        auto sdNode = root.child("sheetData");
        
        REQUIRE(!sdNode.empty());

        auto rowNode = sdNode.child("row");
        REQUIRE(!rowNode.empty());
        REQUIRE(std::string(rowNode.attribute("r").value()) == "1");

        // Cell A1: String (t="inlineStr")
        auto cA1 = rowNode.first_child();
        REQUIRE(std::string(cA1.attribute("r").value()) == "A1");
        REQUIRE(std::string(cA1.attribute("t").value()) == "inlineStr");
        auto isA1 = cA1.child("is");
        REQUIRE(!isA1.empty());
        auto tA1 = isA1.child("t");
        REQUIRE(!tA1.empty());
        REQUIRE(std::string(tA1.text().get()) == "StringData");
        REQUIRE(std::string(tA1.attribute("xml:space").value()) == "preserve");

        // Cell B1: Integer (t="n")
        auto cB1 = cA1.next_sibling();
        REQUIRE(std::string(cB1.attribute("r").value()) == "B1");
        REQUIRE(std::string(cB1.attribute("t").value()) == "n");
        auto vB1 = cB1.child("v");
        REQUIRE(std::string(vB1.text().get()) == "42");

        // Cell C1: Boolean (t="b")
        auto cC1 = cB1.next_sibling();
        REQUIRE(std::string(cC1.attribute("r").value()) == "C1");
        REQUIRE(std::string(cC1.attribute("t").value()) == "b");
        auto vC1 = cC1.child("v");
        REQUIRE(std::string(vC1.text().get()) == "1");
    }
}

TEST_CASE("Stream Writer Buffering – Large Write", "[XLStreamWriter]")
{
    // Write 100K rows; the write buffer must coalesce many small writes into
    // efficient fstream writes and still produce a correct file on close.
    constexpr int kRows = 100000;

    {
        XLDocument doc;
        doc.create("./testXLStreamWriter_Large.xlsx", XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto writer = wks.streamWriter();

        for (int i = 0; i < kRows; ++i) {
            writer.appendRow({i, "Item", i * 1.5, (i % 2 == 0)});
        }
        writer.close();
        doc.save();
        doc.close();
    }

    // Read back and verify row count + spot-check values
    XLDocument doc;
    doc.open("./testXLStreamWriter_Large.xlsx");
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto reader = wks.streamReader();

    int count = 0;
    while (reader.hasNext()) {
        auto row = reader.nextRow();
        REQUIRE(row.size() >= 2);
        REQUIRE(row[0].get<int>() == count);
        ++count;
    }
    REQUIRE(count == kRows);
    doc.close();
}
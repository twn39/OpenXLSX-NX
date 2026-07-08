#include <OpenXLSX.hpp>
#include "XLStreamWriter.hpp"
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
#include <sstream>

using namespace OpenXLSX;

namespace {
inline const std::string& getTestFilename() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("stream_shared_strings") + ".xlsx";
    return name;
}
} // namespace

TEST_CASE("XLStreamWriterSharedStringsTests", "[XLStreamWriter]")
{
    SECTION("Default Inline Mode")
    {
        std::string filename = getTestFilename() + "_default.xlsx";
        
        {
            XLDocument doc;
            doc.create(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            
            // Default streaming mode: useSharedStrings = false
            auto writer = wks.streamWriter();
            writer.appendRow({XLCellValue("Apple"), XLCellValue("Orange"), XLCellValue("Apple")});
            writer.close();
            doc.save();
        }
        
        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            
            // Verify values
            REQUIRE(wks.cell(1, 1).value().get<std::string>() == "Apple");
            REQUIRE(wks.cell(1, 2).value().get<std::string>() == "Orange");
            REQUIRE(wks.cell(1, 3).value().get<std::string>() == "Apple");
            
            // Verify XML representation is inlineStr
            std::stringstream ss1, ss2;
            wks.cell(1, 1).print(ss1);
            wks.cell(1, 2).print(ss2);
            
            REQUIRE(ss1.str().find("t=\"inlineStr\"") != std::string::npos);
            REQUIRE(ss2.str().find("t=\"inlineStr\"") != std::string::npos);
            
            doc.close();
            std::filesystem::remove(filename);
        }
    }

    SECTION("Shared Strings Mode with Cache")
    {
        std::string filename = getTestFilename() + "_shared.xlsx";
        
        {
            XLDocument doc;
            doc.create(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            
            // Enable shared strings
            auto writer = wks.streamWriter(true);
            writer.appendRow({XLCellValue("Banana"), XLCellValue("Cherry"), XLCellValue("Banana")});
            writer.close();
            doc.save();
        }
        
        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            
            // Verify values are read back correctly
            REQUIRE(wks.cell(1, 1).value().get<std::string>() == "Banana");
            REQUIRE(wks.cell(1, 2).value().get<std::string>() == "Cherry");
            REQUIRE(wks.cell(1, 3).value().get<std::string>() == "Banana");
            
            // Verify XML representation is shared string reference (t="s")
            std::stringstream ss1, ss2, ss3;
            wks.cell(1, 1).print(ss1);
            wks.cell(1, 2).print(ss2);
            wks.cell(1, 3).print(ss3);
            
            REQUIRE(ss1.str().find("t=\"s\"") != std::string::npos);
            REQUIRE(ss2.str().find("t=\"s\"") != std::string::npos);
            REQUIRE(ss3.str().find("t=\"s\"") != std::string::npos);
            
            // Verify Banana is deduplicated (cell 1 and cell 3 should refer to the same index)
            // The XML format is: <c r="A1" t="s"><v>0</v></c> and <c r="C1" t="s"><v>0</v></c>
            size_t v1 = ss1.str().find("<v>");
            size_t v3 = ss3.str().find("<v>");
            REQUIRE(v1 != std::string::npos);
            REQUIRE(v3 != std::string::npos);
            REQUIRE(ss1.str().substr(v1, 7) == ss3.str().substr(v3, 7)); // both should have same index value inside <v>...</v>
            
            doc.close();
            std::filesystem::remove(filename);
        }
    }

    SECTION("Safeguard Threshold and Fallback Mode")
    {
        std::string filename = getTestFilename() + "_fallback.xlsx";
        
        {
            XLDocument doc;
            doc.create(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            
            // Enable shared strings but set threshold to 2 unique strings
            auto writer = wks.streamWriter(true, 2);
            writer.appendRow({
                XLCellValue("Red"),    // Unique 1: Shared string (index 0)
                XLCellValue("Green"),  // Unique 2: Shared string (index 1)
                XLCellValue("Red"),    // Duplicate: Shared string (uses index 0)
                XLCellValue("Blue"),   // Unique 3: Exceeds threshold! Should fall back to inlineStr
                XLCellValue("Green"),  // Duplicate: Should still be written as shared string (uses index 1)
                XLCellValue("Yellow")  // Unique 4: Exceeds threshold! Should fall back to inlineStr
            });
            writer.close();
            doc.save();
        }
        
        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            
            // Verify values
            REQUIRE(wks.cell(1, 1).value().get<std::string>() == "Red");
            REQUIRE(wks.cell(1, 2).value().get<std::string>() == "Green");
            REQUIRE(wks.cell(1, 3).value().get<std::string>() == "Red");
            REQUIRE(wks.cell(1, 4).value().get<std::string>() == "Blue");
            REQUIRE(wks.cell(1, 5).value().get<std::string>() == "Green");
            REQUIRE(wks.cell(1, 6).value().get<std::string>() == "Yellow");
            
            // Verify XML representation of each cell
            std::stringstream ss1, ss2, ss3, ss4, ss5, ss6;
            wks.cell(1, 1).print(ss1);
            wks.cell(1, 2).print(ss2);
            wks.cell(1, 3).print(ss3);
            wks.cell(1, 4).print(ss4);
            wks.cell(1, 5).print(ss5);
            wks.cell(1, 6).print(ss6);
            
            // Red, Green, Red (dup), Green (dup) should be shared strings
            REQUIRE(ss1.str().find("t=\"s\"") != std::string::npos);
            REQUIRE(ss2.str().find("t=\"s\"") != std::string::npos);
            REQUIRE(ss3.str().find("t=\"s\"") != std::string::npos);
            REQUIRE(ss5.str().find("t=\"s\"") != std::string::npos);
            
            // Blue and Yellow (exceeded threshold of 2 unique strings) should be inlineStr
            REQUIRE(ss4.str().find("t=\"inlineStr\"") != std::string::npos);
            REQUIRE(ss6.str().find("t=\"inlineStr\"") != std::string::npos);
            
            doc.close();
            std::filesystem::remove(filename);
        }
    }
}

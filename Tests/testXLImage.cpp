#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>
#include <fstream>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testImage_xlsx") + ".xlsx";
    return name;
}
} // namespace


// Helper to read binary file
std::string readBinaryFile(const std::string& path)
{
    std::ifstream file(path, std::ios::binary);
    if (!file) return "";
    return std::string((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
}

TEST_CASE("XLImageTests", "[XLImage]")
{
    SECTION("Add and read images")
    {
        XLDocument doc;
        doc.create(__global_unique_file_0(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        std::string imgData = readBinaryFile("./Tests/test.png");
        REQUIRE_FALSE(imgData.empty());

        wks.addImage("test.png", imgData, 2, 2, 200, 200);

        REQUIRE(wks.images().size() == 1);
        REQUIRE(wks.images()[0].name() == "test.png");
        REQUIRE(wks.images()[0].row() == 2);
        REQUIRE(wks.images()[0].col() == 2);

        doc.save();
        doc.close();

        // Re-open and verify
        XLDocument doc2;
        doc2.open(__global_unique_file_0());
        auto wks2 = doc2.workbook().worksheet("Sheet1");

        auto imgs = wks2.images();
        REQUIRE(imgs.size() == 1);
        REQUIRE(imgs[0].name() == "test.png");

        // Get raw data
        std::string relId   = imgs[0].relationshipId();
        std::string imgPath = wks2.drawing().relationships().relationshipById(relId).target();
        // Relationship target is relative to drawing file (xl/drawings/)
        // internal path should be xl/media/test.png

        // Resolve path (simplified)
        std::string fullPath      = "xl/" + imgPath.substr(3);    // remove "../"
        std::string retrievedData = doc2.getImage(fullPath);
        REQUIRE(retrievedData == imgData);

        doc2.close();
    }
}

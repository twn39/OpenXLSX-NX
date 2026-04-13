#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLColumnCreation_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLColumn_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLColumnByName_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("XLColumnTests", "[XLColumn]")
{
    SECTION("Basic Column Operations")
    {
        XLDocument doc;
        doc.create(__global_unique_file_1(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        auto col1 = wks.column(1);
        col1.setWidth(20.5f);
        REQUIRE(col1.width() == 20.5f);

        col1.setHidden(true);
        REQUIRE(col1.isHidden() == true);

        col1.setHidden(false);
        REQUIRE(col1.isHidden() == false);

        // Test setFormat (assuming style 1 exists or just setting a value)
        col1.setFormat(1);
        REQUIRE(col1.format() == 1);

        doc.save();
        doc.close();
    }

    SECTION("Column Access by Name")
    {
        XLDocument doc;
        doc.create(__global_unique_file_2(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        auto colB = wks.column("B");
        colB.setWidth(15.0f);
        REQUIRE(wks.column(2).width() == 15.0f);

        doc.close();
    }

    SECTION("Automatic Column Node Creation")
    {
        XLDocument doc;
        doc.create(__global_unique_file_0(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // Column 10 doesn't exist yet
        auto col10 = wks.column(10);
        REQUIRE(col10.width() > 0);    // Default width
        col10.setWidth(25.0f);
        REQUIRE(wks.column(10).width() == 25.0f);

        doc.close();
    }
}

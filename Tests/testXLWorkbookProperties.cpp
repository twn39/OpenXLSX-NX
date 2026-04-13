#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLAppProperties_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLWorkbookNames_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLWorkbook_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_3() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLProperties_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("XLPropertiesTests", "[XLProperties]")
{
    SECTION("Core Properties")
    {
        XLDocument doc;
        doc.create(__global_unique_file_3(), XLForceOverwrite);

        doc.setProperty(XLProperty::Title, "Test Title");
        doc.setProperty(XLProperty::Creator, "Test Creator");
        doc.setProperty(XLProperty::Subject, "Test Subject");

        REQUIRE(doc.property(XLProperty::Title) == "Test Title");
        REQUIRE(doc.property(XLProperty::Creator) == "Test Creator");
        REQUIRE(doc.property(XLProperty::Subject) == "Test Subject");

        doc.deleteProperty(XLProperty::Subject);
        // Depending on implementation, it might return empty or old value if not re-read correctly,
        // but property() should reflect current XML state.
        REQUIRE(doc.property(XLProperty::Subject) == "");

        doc.close();
    }

    SECTION("App Properties")
    {
        XLDocument doc;
        doc.create(__global_unique_file_0(), XLForceOverwrite);

        // App properties are usually set via the doc object as well,
        // though some are managed internally (like sheet names).

        doc.workbook().addWorksheet("MySheet");
        auto sheetNames = doc.workbook().sheetNames();
        bool found      = false;
        for (const auto& name : sheetNames) {
            if (name == "MySheet") found = true;
        }
        REQUIRE(found);

        doc.close();
    }
}

TEST_CASE("XLWorkbookTests", "[XLWorkbook]")
{
    SECTION("Sheet Management")
    {
        XLDocument doc;
        doc.create(__global_unique_file_2(), XLForceOverwrite);
        auto wb = doc.workbook();

        // Initial state
        REQUIRE(wb.sheetCount() == 1);
        REQUIRE(wb.sheetExists("Sheet1"));

        // Add worksheet
        wb.addWorksheet("Sheet2");
        REQUIRE(wb.sheetCount() == 2);
        REQUIRE(wb.sheetExists("Sheet2"));

        // Clone worksheet
        wb.cloneSheet("Sheet1", "Sheet1_Clone");
        REQUIRE(wb.sheetCount() == 3);
        REQUIRE(wb.sheetExists("Sheet1_Clone"));

        // Indexing
        REQUIRE(wb.indexOfSheet("Sheet1") == 1);
        REQUIRE(wb.indexOfSheet("Sheet2") == 2);

        wb.setSheetIndex("Sheet2", 1);
        REQUIRE(wb.indexOfSheet("Sheet2") == 1);
        REQUIRE(wb.indexOfSheet("Sheet1") == 2);

        // Delete worksheet
        wb.deleteSheet("Sheet1_Clone");
        REQUIRE(wb.sheetCount() == 2);
        REQUIRE(wb.sheetExists("Sheet1_Clone") == false);

        // Errors
        REQUIRE_THROWS_AS(wb.deleteSheet("NonExistent"), XLException);

        doc.close();
    }

    SECTION("Sheet Names and Types")
    {
        XLDocument doc;
        doc.create(__global_unique_file_1(), XLForceOverwrite);
        auto wb = doc.workbook();

        wb.addWorksheet("Data");

        auto names = wb.sheetNames();
        REQUIRE(names.size() == 2);
        REQUIRE(std::find(names.begin(), names.end(), "Data") != names.end());

        REQUIRE(wb.typeOfSheet("Sheet1") == XLSheetType::Worksheet);
        REQUIRE(wb.typeOfSheet(1) == XLSheetType::Worksheet);

        doc.close();
    }
}

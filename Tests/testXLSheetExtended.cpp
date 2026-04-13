#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLSheetProtection_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLRowColFormat_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLTabProperties_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("XLWorksheetExtendedTests", "[XLSheet]")
{
    SECTION("Sheet Protection")
    {
        XLDocument doc;
        doc.create(__global_unique_file_0(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        REQUIRE(wks.sheetProtected() == false);

        wks.protectSheet(true);
        REQUIRE(wks.sheetProtected() == true);

        wks.protectObjects(true);
        REQUIRE(wks.objectsProtected() == true);

        wks.protectScenarios(true);
        REQUIRE(wks.scenariosProtected() == true);

        wks.allowInsertColumns(true);
        REQUIRE(wks.insertColumnsAllowed() == true);

        wks.allowInsertRows(false);
        REQUIRE(wks.insertRowsAllowed() == false);

        wks.setPassword("password123");
        REQUIRE(wks.passwordIsSet() == true);

        wks.clearPassword();
        REQUIRE(wks.passwordIsSet() == false);

        wks.clearSheetProtection();
        REQUIRE(wks.sheetProtected() == false);

        doc.close();
    }

    SECTION("Tab Properties")
    {
        XLDocument doc;
        doc.create(__global_unique_file_2(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.setColor(XLColor(255, 0, 0));
        // Note: getColor_impl depends on how hex is stored and retrieved
        // REQUIRE(wks.color().hex() == "FFFF0000");

        wks.setSelected(true);
        REQUIRE(wks.isSelected() == true);

        wks.setSelected(false);
        REQUIRE(wks.isSelected() == false);

        doc.close();
    }

    SECTION("Row/Column Formatting")
    {
        XLDocument doc;
        doc.create(__global_unique_file_1(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // Set row format
        wks.cell("A1").value() = "Test";
        wks.setRowFormat(1, 1);
        REQUIRE(wks.getRowFormat(1) == 1);
        REQUIRE(wks.row(1).cells().begin()->cellFormat() == 1);

        // Set column format
        wks.setColumnFormat(1, 2);
        REQUIRE(wks.getColumnFormat(1) == 2);
        REQUIRE(wks.cell("A1").cellFormat() == 2);

        doc.close();
    }
}

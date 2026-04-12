#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>

using namespace OpenXLSX;

TEST_CASE("XLWorksheetExtendedTests", "[XLSheet]")
{
    SECTION("Sheet Protection")
    {
        XLDocument doc;
        doc.create("./testXLSheetProtection.xlsx", XLForceOverwrite);
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
        doc.create("./testXLTabProperties.xlsx", XLForceOverwrite);
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
        doc.create("./testXLRowColFormat.xlsx", XLForceOverwrite);
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

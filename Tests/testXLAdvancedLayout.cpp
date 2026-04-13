#include "TestHelpers.hpp"
#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("Layout_Test_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("AdvancedPrintLayout", "[PrintLayout]")
{
    XLDocument doc;
    doc.create(__global_unique_file_0(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    wks.setName("My Sheet");

    // Fill some data
    for (uint32_t r = 1; r <= 20; ++r) {
        for (uint16_t c = 1; c <= 10; ++c) {
            wks.cell(r, c).value() = r * c;
        }
    }

    // Set Print Area and Titles
    wks.setPrintArea("$A$1:$F$10");
    wks.setPrintTitleRows(1, 2);
    wks.setPrintTitleCols(1, 2);

    // Outline Grouping
    wks.groupRows(5, 10, 1, true);       // Group rows 5-10, initially collapsed
    wks.groupColumns(3, 5, 1, false);    // Group columns C-E, expanded

    doc.save();

    // Verify
    auto     definedNames = doc.workbook().definedNames();
    uint32_t localSheetId = wks.index() - 1;

    auto printArea = definedNames.get("_xlnm.Print_Area", localSheetId);
    REQUIRE(printArea.valid() == true);
    REQUIRE(printArea.refersTo() == "'My Sheet'!$A$1:$F$10");

    auto printTitles = definedNames.get("_xlnm.Print_Titles", localSheetId);
    REQUIRE(printTitles.valid() == true);
    REQUIRE(printTitles.refersTo() == "'My Sheet'!$1:$2,'My Sheet'!$A:$B");

    // Check Outline PR
    auto outlinePr = wks.xmlDocument().document_element().child("sheetPr").child("outlinePr");
    REQUIRE(std::string(outlinePr.attribute("summaryBelow").value()) == "1");
    REQUIRE(std::string(outlinePr.attribute("summaryRight").value()) == "1");

    doc.close();
}

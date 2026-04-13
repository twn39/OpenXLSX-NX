#include "TestHelpers.hpp"
#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("NamedStyles_Test_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("NamedStylesCreationandApplication", "[Styles][NamedStyle]")
{
    XLDocument doc;
    doc.create(__global_unique_file_0(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    auto styles = doc.styles();

    // Create a new font
    auto fontIdx = styles.fonts().create();
    auto font    = styles.fonts()[fontIdx];
    font.setBold(true);
    font.setFontColor(XLColor("FFFF0000"));    // Red
    font.setFontSize(20);

    // Create a new fill
    auto fillIdx = styles.fills().create();
    auto fill    = styles.fills()[fillIdx];
    fill.setFillType(XLPatternFill);
    fill.setPatternType(XLPatternSolid);
    fill.setColor(XLColor("FFFFFF00"));    // Yellow
    fill.setBackgroundColor(XLColor("FFFFFF00"));

    // Create a named style
    auto nsIdx = styles.addNamedStyle("Highlight", fontIdx, fillIdx);

    // Verify lookup
    REQUIRE(styles.namedStyle("Highlight") == nsIdx);
    REQUIRE(styles.namedStyle("NonExistent") == XLInvalidStyleIndex);

    // Apply to cells
    wks.cell("B2").value() = "Highlight Me";
    wks.cell("B2").setCellFormat(nsIdx);

    // Also apply using lookup
    wks.cell("C3").value() = "Me Too";
    wks.cell("C3").setCellFormat(styles.namedStyle("Highlight"));

    // Verify OOXML DOM
    auto stylesNode = doc.styles().xmlDocument().document_element();

    // cellStyles
    auto cellStylesNode = stylesNode.child("cellStyles");
    REQUIRE(cellStylesNode.child("cellStyle").next_sibling("cellStyle").attribute("name").value() == std::string("Highlight"));
    REQUIRE(cellStylesNode.child("cellStyle").next_sibling("cellStyle").attribute("builtinId").empty() ==
            true);    // Custom styles shouldn't have builtinId
    REQUIRE(cellStylesNode.child("cellStyle").next_sibling("cellStyle").attribute("xfId").as_uint() ==
            1);    // Should point to index 1 of cellStyleXfs

    // cellStyleXfs
    auto cellStyleXfsNode = stylesNode.child("cellStyleXfs");
    auto xfNode           = cellStyleXfsNode.child("xf").next_sibling("xf");
    REQUIRE(xfNode.attribute("fontId").as_uint() == fontIdx);
    REQUIRE(xfNode.attribute("fillId").as_uint() == fillIdx);

    // cellXfs
    auto cellXfsNode   = stylesNode.child("cellXfs");
    auto appliedXfNode = cellXfsNode.child("xf").next_sibling("xf");
    REQUIRE(appliedXfNode.attribute("xfId").as_uint() == 1);            // Should point to the named style format
    REQUIRE(appliedXfNode.attribute("fontId").as_uint() == fontIdx);    // Should inherit and duplicate font
    REQUIRE(appliedXfNode.attribute("fillId").as_uint() == fillIdx);

    doc.save();
    doc.close();
}

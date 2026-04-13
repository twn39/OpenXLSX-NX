#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
#include <filesystem>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("stream_style_test_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("StreamingStyledTest_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("StreamWriterAdvancedFeatures", "[XLStreamWriter][Styles]")
{
    XLDocument doc;
    doc.create(__global_unique_file_0(), XLForceOverwrite);

    XLStyle s1;
    s1.font.bold  = true;
    s1.font.color = XLColor("FF0000");
    auto style1   = doc.styles().findOrCreateStyle(s1);

    XLStyle s2;
    s2.font.italic = true;
    s2.font.color  = XLColor("0000FF");
    auto style2    = doc.styles().findOrCreateStyle(s2);

    auto wks    = doc.workbook().worksheet("Sheet1");
    auto stream = wks.streamWriter();

    // 1. Basic appendRow (Unstyled via vector<XLCellValue>)
    std::vector<XLCellValue> headerRow = {"Header 1", "Header 2", "Header 3"};
    stream.appendRow(headerRow);

    // 2. Styled appendRow (via vector<XLStreamCell>)
    stream.appendRow({
        XLStreamCell("Red Bold", style1),
        XLStreamCell("Blue Italic", style2),
        XLStreamCell("Unstyled Value")    // uses default
    });

    stream.close();
    doc.save();
    doc.close();

    // Verify
    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open(__global_unique_file_0()));
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    REQUIRE(wks2.cell("A1").value().get<std::string>() == "Header 1");

    REQUIRE(wks2.cell("A2").value().get<std::string>() == "Red Bold");
    REQUIRE(wks2.cell("A2").cellFormat() == style1);

    REQUIRE(wks2.cell("B2").value().get<std::string>() == "Blue Italic");
    REQUIRE(wks2.cell("B2").cellFormat() == style2);

    REQUIRE(wks2.cell("C2").value().get<std::string>() == "Unstyled Value");
    REQUIRE(wks2.cell("C2").cellFormat() == XLDefaultCellFormat);

    doc2.close();
    std::filesystem::remove(__global_unique_file_0());
}
TEST_CASE("GenerateUserReviewStyledStreamingFile", "[XLStreamWriter][User]")
{
    XLDocument doc;
    doc.create(__global_unique_file_1(), XLForceOverwrite);
    
    // Create styles
    XLStyle headerStyle;
    headerStyle.font.bold = true;
    headerStyle.font.color = XLColor("FFFFFF");
    headerStyle.fill.pattern = XLPatternSolid;
    headerStyle.fill.fgColor = XLColor("4F81BD");
    auto headerId = doc.styles().findOrCreateStyle(headerStyle);

    XLStyle currencyStyle;
    currencyStyle.numberFormat = "$#,##0.00";
    auto currencyId = doc.styles().findOrCreateStyle(currencyStyle);

    XLStyle percentageStyle;
    percentageStyle.numberFormat = "0.00%";
    auto percentId = doc.styles().findOrCreateStyle(percentageStyle);

    XLStyle altRowStyle;
    altRowStyle.fill.pattern = XLPatternSolid;
    altRowStyle.fill.fgColor = XLColor("DCE6F1");
    auto altRowId = doc.styles().findOrCreateStyle(altRowStyle);

    auto wks = doc.workbook().worksheet("Sheet1");
    wks.column(1).setWidth(15);
    wks.column(2).setWidth(12);
    wks.column(3).setWidth(12);

    auto stream = wks.streamWriter();

    // Headers
    stream.appendRow({
        XLStreamCell("Product", headerId),
        XLStreamCell("Price", headerId),
        XLStreamCell("Margin", headerId)
    });

    // Rows
    for (int i = 1; i <= 20; ++i) {
        uint32_t styleId = (i % 2 == 0) ? altRowId : 0; 
        
        XLStyle rowCurrencyStyle = currencyStyle;
        if (i % 2 == 0) {
            rowCurrencyStyle.fill.pattern = XLPatternSolid;
            rowCurrencyStyle.fill.fgColor = XLColor("DCE6F1");
        }
        auto rowCurrencyId = doc.styles().findOrCreateStyle(rowCurrencyStyle);

        XLStyle rowPercentStyle = percentageStyle;
        if (i % 2 == 0) {
            rowPercentStyle.fill.pattern = XLPatternSolid;
            rowPercentStyle.fill.fgColor = XLColor("DCE6F1");
        }
        auto rowPercentId = doc.styles().findOrCreateStyle(rowPercentStyle);

        stream.appendRow({
            XLStreamCell("Product " + std::to_string(i), styleId),
            XLStreamCell(10.5 * i, rowCurrencyId),
            XLStreamCell(0.15 + (i * 0.01), rowPercentId)
        });
    }

    stream.close();
    doc.save();
    doc.close();
}

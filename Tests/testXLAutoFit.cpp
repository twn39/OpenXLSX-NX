#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
#include <filesystem>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("autofit_test_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("AutoFitColumnWidthTests", "[XLWorksheet][AutoFit]")
{
    XLDocument doc;
    doc.create(__global_unique_file_0(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    SECTION("Basic ASCII AutoFit")
    {
        wks.cell("A1").value() = "A";
        wks.cell("A2").value() = "Very long string that needs more space";
        wks.cell("A3").value() = "Short";

        wks.autoFitColumn(1);
        
        doc.save();
        doc.close();

        XLDocument doc2;
        doc2.open(__global_unique_file_0());
        auto wks2 = doc2.workbook().worksheet("Sheet1");
        
        // "Very long string that needs more space" is 38 chars
        // 38 * 1.2 + 1.5 = 47.1
        REQUIRE(wks2.column(1).width() > 30.0f);
        doc2.close();
    }

    SECTION("UTF-8 Multi-byte AutoFit (CJK)")
    {
        // 5 CJK characters. In UTF-8, these are 3 bytes each.
        // Under the new algorithm: 5 chars * 2.1 = 10.5 width
        // Padding = 1.5, Total ~12.0
        wks.cell("B1").value() = "中文测试字"; 
        
        wks.autoFitColumn(2);
        
        doc.save();
        doc.close();

        XLDocument doc2;
        doc2.open(__global_unique_file_0());
        auto wks2 = doc2.workbook().worksheet("Sheet1");
        
        float width = wks2.column(2).width();
        REQUIRE(width > 11.5f);
        REQUIRE(width < 13.0f); // It shouldn't balloon to 18+ (which it did when treating as 15 ASCII bytes)
        doc2.close();
    }
    
    std::filesystem::remove(__global_unique_file_0());
}
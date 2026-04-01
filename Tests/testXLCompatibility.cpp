#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <filesystem>
#include <vector>
#include <string>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Cross-Engine Compatibility", "[Compatibility]")
{
    // These files are expected to be copied into the test run directory by CMake
    std::vector<std::string> testFiles = {
        "openpyxl_generated.xlsx",
        "excelize_generated.xlsx"
    };

    for (const auto& file : testFiles) {
        SECTION("Load, verify and re-save " + file) {
            std::string filepath = "Fixtures/" + file;
            if (!std::filesystem::exists(filepath)) {
                // Try current dir
                if (std::filesystem::exists(file)) {
                    filepath = file;
                } else {
                    // Try looking in the source tree if tests are run from build dir
                    filepath = "../Tests/Fixtures/" + file;
                    if (!std::filesystem::exists(filepath)) {
                        FAIL("Test fixture not found: " << file);
                    }
                }
            }

            INFO("Testing compatibility with: " << filepath);
            
            XLDocument doc;
            REQUIRE_NOTHROW(doc.open(filepath));

            auto wks = doc.workbook().worksheet("TestSheet");
            
            // 1. Verify content
            REQUIRE(wks.cell("A1").value().get<std::string>().find("Hello from") != std::string::npos);
            REQUIRE(wks.cell("A2").value().get<int64_t>() == 42);
            REQUIRE(wks.cell("A3").value().get<double>() == Catch::Approx(3.1415));
            REQUIRE(wks.cell("B1").value().get<std::string>() == "Styled");

            // 2. Modify content
            wks.cell("C1").value() = "Modified by OpenXLSX";

            // 3. Save as new file (this ensures we don't break the original but tests the write-path)
            std::string outFile = "modified_" + file;
            REQUIRE_NOTHROW(doc.saveAs(outFile, XLForceOverwrite));
            doc.close();

            // 4. Re-open to verify our write didn't corrupt it
            XLDocument doc2;
            REQUIRE_NOTHROW(doc2.open(outFile));
            auto wks2 = doc2.workbook().worksheet("TestSheet");
            REQUIRE(wks2.cell("C1").value().get<std::string>() == "Modified by OpenXLSX");
            doc2.close();
            
            std::filesystem::remove(outFile);
        }
    }
}
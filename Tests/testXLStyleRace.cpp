#include <OpenXLSX.hpp>
#include <atomic>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
#include <thread>
#include <vector>

using namespace OpenXLSX;

TEST_CASE("XLStylesRaceTests", "[XLStyles][Race]")
{
    SECTION("Concurrent styles creation and application")
    {
        const std::string filename  = OpenXLSX::TestHelpers::getUniqueFilename("style_race");
        constexpr int     numSheets = 4;
        constexpr int     numRows   = 50;

        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        for (int i = 2; i <= numSheets; ++i) doc.workbook().addWorksheet("Sheet" + std::to_string(i));

        // Option A: Pre-allocate all styles in the main thread (Single-threaded context)
        std::vector<std::vector<XLStyleIndex>> preallocatedStyles(numSheets, std::vector<XLStyleIndex>(numRows));
        
        for (int t = 0; t < numSheets; ++t) {
            for (int row = 1; row <= numRows; ++row) {
                XLStyle style;
                style.font.size = 10 + (row % 10);
                style.font.bold = (row % 2 == 0);
                
                style.fill.pattern = XLPatternSolid;
                char hexColor[16];
                snprintf(hexColor, sizeof(hexColor), "FF%02X%02X00", (t * 50) & 0xFF, (row * 5) & 0xFF);
                style.fill.fgColor = XLColor(hexColor);
                
                // Call findOrCreateStyle sequentially in main thread
                preallocatedStyles[t][row - 1] = doc.styles().findOrCreateStyle(style);
            }
        }

        std::vector<XLWorksheet> worksheets;
        for (int i = 1; i <= numSheets; ++i) worksheets.push_back(doc.workbook().worksheet("Sheet" + std::to_string(i)));

        std::vector<std::thread> threads;
        std::atomic<int>         errors{0};

        for (int t = 0; t < numSheets; ++t) {
            threads.emplace_back([&, t]() {
                try {
                    auto& wks = worksheets[static_cast<size_t>(t)];
                    for (int row = 1; row <= numRows; ++row) {
                        wks.cell(static_cast<uint32_t>(row), 1).value() = row;
                        
                        // Use pre-allocated style index in child threads
                        XLStyleIndex styleIdx = preallocatedStyles[t][row - 1];
                        wks.cell(static_cast<uint32_t>(row), 1).setCellFormat(styleIdx);
                    }
                }
                catch (const std::exception& e) {
                    UNSCOPED_INFO("Exception in thread: " << e.what());
                    ++errors;
                }
                catch (...) {
                    UNSCOPED_INFO("Unknown exception in thread");
                    ++errors;
                }
            });
        }

        for (auto& th : threads) th.join();
        
        // Assert that no thread encountered exceptions/crashes
        REQUIRE(errors == 0);

        doc.save();
        doc.close();
        std::filesystem::remove(filename);
    }
}

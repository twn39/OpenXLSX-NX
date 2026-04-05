#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <thread>
#include <vector>
#include <string>
#include <mutex>
#include <atomic>

using namespace OpenXLSX;

TEST_CASE("High Pressure Concurrent Writes", "[.hide][XLConcurrent][TSAN]")
{
    std::string filename = "test_high_pressure_concurrent.xlsx";
    
    // 1. Create a large workbook with 100 sheets
    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        for (int i = 2; i <= 100; ++i) {
            doc.workbook().addWorksheet("Sheet" + std::to_string(i));
        }
        doc.save();
        doc.close();
    }

    // 2. Open it and fire 100 threads, each writing to a DIFFERENT sheet concurrently
    // Note: OpenXLSX architecture explicitly states "Supports concurrent writes to different worksheets. Writing to the same worksheet is unsafe."
    {
        XLDocument doc;
        doc.open(filename);

        const int num_threads = 100;
        const int rows_per_thread = 1000;
        std::vector<std::thread> threads;
        std::atomic<int> completed_threads{0};

        for (int i = 1; i <= num_threads; ++i) {
            threads.emplace_back([&doc, i, &completed_threads, rows_per_thread]() {
                try {
                    // Each thread gets its own distinct worksheet
                    std::string sheetName = "Sheet" + std::to_string(i);
                    auto wks = doc.workbook().worksheet(sheetName);
                    
                    // Write intensive data
                    for (uint32_t r = 1; r <= rows_per_thread; ++r) {
                        wks.cell(r, 1).value() = "Thread-" + std::to_string(i);
                        wks.cell(r, 2).value() = r * i; // Numeric calculation
                        wks.cell(r, 3).value() = r % 2 == 0; // Boolean
                    }
                    completed_threads++;
                } catch (...) {
                    // Catch everything to avoid tearing down the test suite on a single thread fail
                }
            });
        }

        // Wait for all threads to finish their DOM manipulations
        for (auto& t : threads) {
            if (t.joinable()) {
                t.join();
            }
        }

        REQUIRE(completed_threads.load() == num_threads);

        doc.save();
        doc.close();
    }

    // 3. Verify the data consistency sequentially to ensure no XML trees were corrupted during concurrent saves
    {
        XLDocument doc;
        doc.open(filename);
        
        for (int i = 1; i <= 100; ++i) {
            std::string sheetName = "Sheet" + std::to_string(i);
            auto wks = doc.workbook().worksheet(sheetName);
            
            REQUIRE(wks.cell(1, 1).value().get<std::string>() == "Thread-" + std::to_string(i));
            REQUIRE(wks.cell(500, 2).value().get<int64_t>() == 500 * i);
            REQUIRE(wks.cell(999, 3).value().get<bool>() == false);
        }
        doc.close();
    }
    
    std::remove(filename.c_str());
}

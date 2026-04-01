#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <filesystem>
#include <vector>
#include <string>
#include <iostream>

#ifdef _WIN32
#include <windows.h>
#include <psapi.h>
#elif defined(__APPLE__) && defined(__MACH__)
#include <mach/mach.h>
#else
#include <fstream>
#include <string>
#include <unistd.h>
#endif

using namespace OpenXLSX;

size_t getCurrentRSS() {
#ifdef _WIN32
    PROCESS_MEMORY_COUNTERS info;
    if (GetProcessMemoryInfo(GetCurrentProcess(), &info, sizeof(info)))
        return (size_t)info.WorkingSetSize;
    return 0;
#elif defined(__APPLE__) && defined(__MACH__)
    struct mach_task_basic_info info;
    mach_msg_type_number_t infoCount = MACH_TASK_BASIC_INFO_COUNT;
    if (task_info(mach_task_self(), MACH_TASK_BASIC_INFO, (task_info_t)&info, &infoCount) != KERN_SUCCESS)
        return 0;
    return (size_t)info.resident_size;
#else
    long rss = 0L;
    std::ifstream statm("/proc/self/statm");
    if (statm) {
        long pages = 0;
        statm >> pages >> rss;
        return (size_t)rss * sysconf(_SC_PAGE_SIZE);
    }
    return 0;
#endif
}

TEST_CASE("Streaming Giant Document Memory Test", "[Streaming][OOM]")
{
    const std::string filename = "GiantStreamingTest.xlsx";
    // 1 Million cells. DOM parsing this would easily take 100-200MB natively. 
    // Streaming should theoretically keep this under 50MB (mostly string buffer overhead).
    const uint32_t numRows = 100000;
    const uint16_t numCols = 10;
    
    // NOTE: AddressSanitizer (ASan) adds massive overhead (redzones/quarantine) to millions of small string allocations.
    // We set a generous limit specifically when sanitizers are active to avoid CI false positives.
#if defined(__has_feature)
#if __has_feature(address_sanitizer)
    const size_t maxAllowedMemoryGrowthMB = 1500; // ASan enabled (up to 1.5GB overhead for quarantine)
#else
    const size_t maxAllowedMemoryGrowthMB = 200;  // Normal
#endif
#elif defined(__SANITIZE_ADDRESS__)
    const size_t maxAllowedMemoryGrowthMB = 1500; // ASan enabled
#else
    const size_t maxAllowedMemoryGrowthMB = 200;  // Normal
#endif

    SECTION("XLStreamWriter Memory Usage")
    {
        size_t initialMem = getCurrentRSS();
        INFO("Initial Memory: " << initialMem / (1024 * 1024) << " MB");

        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        auto stream = wks.streamWriter();
        
        std::vector<XLCellValue> rowData(numCols, XLCellValue("Streaming Data"));
        for (uint32_t r = 1; r <= numRows; ++r) {
            stream.appendRow(rowData);
            
            // Periodically check memory to catch streaming leaks
            if (r % 50000 == 0) {
                size_t currentMem = getCurrentRSS();
                size_t memIncreaseMB = (currentMem >= initialMem) ? (currentMem - initialMem) / (1024 * 1024) : 0;
                INFO("Memory after " << r << " rows: " << currentMem / (1024 * 1024) << " MB (+" << memIncreaseMB << " MB)");
                REQUIRE(memIncreaseMB < maxAllowedMemoryGrowthMB);
            }
        }
        
        stream.close();
        doc.save();
        doc.close();

        size_t finalMem = getCurrentRSS();
        size_t memIncreaseMB = (finalMem >= initialMem) ? (finalMem - initialMem) / (1024 * 1024) : 0;
        INFO("Final Memory after Writing & Saving: " << finalMem / (1024 * 1024) << " MB (+" << memIncreaseMB << " MB)");
        REQUIRE(memIncreaseMB < maxAllowedMemoryGrowthMB);
    }

    SECTION("XLStreamReader Memory Usage")
    {
        // Skip if writer test failed and file isn't there
        if (!std::filesystem::exists(filename)) return;

        size_t initialMem = getCurrentRSS();
        INFO("Initial Memory before Reading: " << initialMem / (1024 * 1024) << " MB");

        XLDocument doc;
        REQUIRE_NOTHROW(doc.open(filename));
        auto wks = doc.workbook().worksheet("Sheet1");

        auto reader = wks.streamReader();
        uint32_t readRows = 0;
        
        while (reader.hasNext()) {
            auto row = reader.nextRow();
            REQUIRE(row.size() == numCols);
            REQUIRE(row[0].type() == XLValueType::String);
            readRows++;
            
            if (readRows % 50000 == 0) {
                size_t currentMem = getCurrentRSS();
                size_t memIncreaseMB = (currentMem >= initialMem) ? (currentMem - initialMem) / (1024 * 1024) : 0;
                INFO("Memory after reading " << readRows << " rows: " << currentMem / (1024 * 1024) << " MB (+" << memIncreaseMB << " MB)");
                REQUIRE(memIncreaseMB < maxAllowedMemoryGrowthMB);
            }
        }
        REQUIRE(readRows == numRows);
        doc.close();
        std::filesystem::remove(filename);
    }
}

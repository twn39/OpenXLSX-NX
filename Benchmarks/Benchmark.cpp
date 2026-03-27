#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>
#include <catch2/benchmark/catch_benchmark.hpp>
#include <deque>
#include <fstream>
#include <numeric>
#include <vector>

using namespace OpenXLSX;

// Using a slightly smaller row count for regular benchmarking to avoid excessive runtimes, 
// but enough to show performance characteristics.
constexpr uint64_t rowCount = 100000; 
constexpr uint8_t  colCount = 8;

TEST_CASE("OpenXLSX Benchmarks", "[.benchmark]")
{
    SECTION("Write Operations")
    {
        BENCHMARK("Write Strings")
        {
            XLDocument doc;
            doc.create("./benchmark_strings.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            std::deque<XLCellValue> values(colCount, "OpenXLSX");

            for (auto& row : wks.rows(rowCount)) row.values() = values;

            doc.save();
            doc.close();
            return rowCount * colCount;
        };

        BENCHMARK("Write Integers")
        {
            XLDocument doc;
            doc.create("./benchmark_integers.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            std::vector<XLCellValue> values(colCount, 42);

            for (auto& row : wks.rows(rowCount)) row.values() = values;

            doc.save();
            doc.close();
            return rowCount * colCount;
        };

        BENCHMARK("Write Floats")
        {
            XLDocument doc;
            doc.create("./benchmark_floats.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            std::vector<XLCellValue> values(colCount, 3.14);

            for (auto& row : wks.rows(rowCount)) row.values() = values;

            doc.save();
            doc.close();
            return rowCount * colCount;
        };

        BENCHMARK("Write Bools")
        {
            XLDocument doc;
            doc.create("./benchmark_bools.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            std::vector<XLCellValue> values(colCount, true);

            for (auto& row : wks.rows(rowCount)) row.values() = values;

            doc.save();
            doc.close();
            return rowCount * colCount;
        };
    }

    SECTION("Read Operations")
    {
        // Prerequisites for read benchmarks (ensure files exist)
        {
            XLDocument doc;
            if (!std::ifstream("./benchmark_strings.xlsx")) {
                doc.create("./benchmark_strings.xlsx", XLForceOverwrite);
                auto wks = doc.workbook().worksheet("Sheet1");
                std::deque<XLCellValue> values(colCount, "OpenXLSX");
                for (auto& row : wks.rows(rowCount)) row.values() = values;
                doc.save();
            }
            if (!std::ifstream("./benchmark_integers.xlsx")) {
                doc.create("./benchmark_integers.xlsx", XLForceOverwrite);
                auto wks = doc.workbook().worksheet("Sheet1");
                std::vector<XLCellValue> values(colCount, 42);
                for (auto& row : wks.rows(rowCount)) row.values() = values;
                doc.save();
            }
        }

        BENCHMARK("Read Strings")
        {
            XLDocument doc;
            doc.open("./benchmark_strings.xlsx");
            auto     wks    = doc.workbook().worksheet("Sheet1");
            uint64_t result = 0;

            for (auto& row : wks.rows()) {
                std::vector<std::string> values = std::vector<std::string>(row.values());
                result += std::count_if(values.begin(), values.end(), [](const std::string& v) {
                    return !v.empty();
                });
            }
            doc.close();
            return result;
        };

        BENCHMARK("Read Integers")
        {
            XLDocument doc;
            doc.open("./benchmark_integers.xlsx");
            auto     wks    = doc.workbook().worksheet("Sheet1");
            uint64_t result = 0;

            for (auto& row : wks.rows()) {
                std::vector<XLCellValue> values = row.values();
                result += std::accumulate(values.begin(),
                                          values.end(),
                                          0ULL,
                                          [](uint64_t a, const XLCellValue& b) { return a + b.get<uint64_t>(); });
            }
            doc.close();
            return result;
        };
    }


    SECTION("New Features Operations")
    {
        BENCHMARK("Formula Engine - Parse & Eval")
        {
            XLDocument doc;
            doc.create("./benchmark_formula.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            
            // Setup simple data
            wks.cell("A1").value() = 100.5;
            wks.cell("A2").value() = 200.5;
            wks.cell("A3").value() = 300.5;
            
            XLFormulaEngine engine;
            auto resolver = XLFormulaEngine::makeResolver(wks);
            
            double dummy_sum = 0;
            // Benchmark parsing and evaluating formula strings repeatedly
            for(int i = 0; i < 10000; ++i) {
                XLCellValue result = engine.evaluate("SUM(A1:A3)", resolver);
                dummy_sum += result.get<double>();
            }
            
            doc.close();
            std::filesystem::remove("./benchmark_formula.xlsx");
            return dummy_sum;
        };

        BENCHMARK("Style Pool - Deduplication")
        {
            XLDocument doc;
            doc.create("./benchmark_styles.xlsx", XLForceOverwrite);
            auto styles = doc.styles();
            
            XLStyle s;
            s.font.name = "Arial";
            s.font.size = 12;
            s.font.bold = true;
            s.fill.pattern = XLPatternSolid;
            s.fill.fgColor = XLColor("FFFF0000");

            size_t calls = 50000;
            // The first call registers the style, the next 49999 calls should hit the O(1) cache
            for(size_t i = 0; i < calls; ++i) {
                styles.findOrCreateStyle(s);
            }
            
            doc.close();
            std::filesystem::remove("./benchmark_styles.xlsx");
            return calls;
        };
    }

}

// Override Catch2's default main to force a lower benchmark sample rate
#define CATCH_CONFIG_RUNNER
#include <catch2/catch_session.hpp>

int main(int argc, char* argv[]) {
    Catch::Session session;

    // Build the default configuration
    auto configData = session.configData();
    // Default benchmark samples is 100, which takes forever. Hardcode to 5.
    configData.benchmarkSamples = 5;
    
    session.useConfigData(configData);

    // Let the user override via command line if they really want to
    int returnCode = session.applyCommandLine(argc, argv);
    if(returnCode != 0) // Indicates a command line error
        return returnCode;

    return session.run();
}

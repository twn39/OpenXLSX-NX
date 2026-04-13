#ifndef OPENXLSX_TESTHELPERS_HPP
#define OPENXLSX_TESTHELPERS_HPP

#include <string>
#include <random>
#include <filesystem>
#include <atomic>

namespace OpenXLSX {
namespace TestHelpers {

    inline std::string getUniqueFilename(const std::string& prefix = "test") {
        static std::atomic<uint64_t> counter{0};
        static std::random_device rd;
        static std::mt19937_64 gen(rd());
        static std::uniform_int_distribution<uint64_t> dis;

        std::filesystem::path tmpDir = std::filesystem::current_path() / "tmp";
        std::error_code ec;
        std::filesystem::create_directories(tmpDir, ec);

        std::string filename = prefix + "_" + std::to_string(counter++) + "_" + std::to_string(dis(gen)) + ".xlsx";
        return (tmpDir / filename).string();
    }

} // namespace TestHelpers
} // namespace OpenXLSX

#endif // OPENXLSX_TESTHELPERS_HPP

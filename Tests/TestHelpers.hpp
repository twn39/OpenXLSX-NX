#ifndef OPENXLSX_TESTHELPERS_HPP
#define OPENXLSX_TESTHELPERS_HPP

#include <string>
#include <chrono>
#include <filesystem>
#include <atomic>

namespace OpenXLSX {
namespace TestHelpers {

    inline std::string getUniqueFilename(const std::string& prefix = "test") {
        static std::atomic<uint64_t> counter{0};
        auto now = std::chrono::high_resolution_clock::now().time_since_epoch().count();

        std::error_code ec;
        std::filesystem::path tmpDir = std::filesystem::current_path() / "tmp";
        std::filesystem::create_directories(tmpDir, ec);

        std::string filename = prefix + "_" + std::to_string(now) + "_" + std::to_string(counter++) + ".xlsx";
        return (tmpDir / filename).string();
    }

} // namespace TestHelpers
} // namespace OpenXLSX

#endif // OPENXLSX_TESTHELPERS_HPP

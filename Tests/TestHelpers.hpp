/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

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

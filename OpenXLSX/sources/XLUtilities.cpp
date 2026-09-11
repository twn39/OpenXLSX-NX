/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#include "XLUtilities.hpp"
#include "XLException.hpp"
#include <array>
#include <gsl/span>
#include <string>
#include <vector>

namespace OpenXLSX
{

    /**
     * @brief Return a hexadecimal digit as character that is the equivalent of value
     */
    [[nodiscard]] constexpr char hexDigit(unsigned int value)
    {
        if (value > 0xf) return 0;
        if (value < 0xa) return static_cast<char>(value + '0');
        return static_cast<char>((value - 0xa) + 'a');
    }

    /**
     * @details Converts raw byte arrays into hex-encoded strings, required for constructing valid OOXML password hashes.
     */
    std::string BinaryAsHexString(gsl::span<const std::byte> data)
    {
        if (data.empty()) return "";
        std::string strAssemble(data.size() * 2, 0);
        for (size_t i = 0; i < data.size(); ++i) {
            auto valueByte         = static_cast<uint8_t>(data[i]);
            strAssemble[i * 2]     = hexDigit(valueByte >> 4);
            strAssemble[i * 2 + 1] = hexDigit(valueByte & 0x0f);
        }
        return strAssemble;
    }

    /**
     * @details Implements the legacy Excel hash algorithm to verify or secure sheet protection features.
     */
    uint16_t ExcelPasswordHash(std::string_view password)
    {
        uint16_t       wPasswordHash = 0;
        const uint16_t cchPassword   = gsl::narrow_cast<uint16_t>(password.length());

        for (uint16_t pos = 0; pos < cchPassword; ++pos) {
            uint32_t byteHash = static_cast<uint32_t>(password[pos]) << ((pos + 1) % 15);
            byteHash          = (byteHash >> 15) | (byteHash & 0x7fff);
            wPasswordHash ^= static_cast<uint16_t>(byteHash);
        }
        wPasswordHash ^= cchPassword ^ 0xce4b;
        return wPasswordHash;
    }

    /**
     * @details Bridges the legacy integer hash to the required XML string format, simplifying integration with the DOM attributes.
     */
    std::string ExcelPasswordHashAsString(std::string_view password)
    {
        const uint16_t                 pw = ExcelPasswordHash(password);
        const std::array<std::byte, 2> hashData{static_cast<std::byte>(pw >> 8), static_cast<std::byte>(pw & 0xff)};
        return BinaryAsHexString(gsl::make_span(hashData.data(), 2));
    }

    /**
     * @details Split a path into its constituent entries (directories and file).
     * Rationale: Standard path decomposition used for relative path calculation.
     * Uses std::string_view for zero-copy parsing.
     */
    std::vector<std::string> disassemblePath(std::string_view path, bool eliminateDots = true)
    {
        std::vector<std::string> result;
        size_t                   startpos = (!path.empty() && path.front() == '/' ? 1 : 0);
        size_t                   pos;
        do {
            pos = path.find('/', startpos);
            if (pos == startpos) throw XLInternalError("path must not contain two subsequent forward slashes");

            std::string_view dirEntry = path.substr(startpos, pos - startpos);
            if (!dirEntry.empty()) {
                if (eliminateDots) {
                    if (dirEntry == ".") {}
                    else if (dirEntry == "..") {
                        if (!result.empty())
                            result.pop_back();
                        else
                            throw XLInternalError("no remaining directory to exit with ..");
                    }
                    else
                        result.emplace_back(dirEntry);
                }
                else
                    result.emplace_back(dirEntry);
            }
            startpos = pos + 1;
        }
        while (pos != std::string_view::npos);

        return result;
    }

    /**
     * @details Calculate the relative path from B to A.
     * Rationale: Resolves internal OOXML relationship paths, which are often stored
     * relative to the source part's directory.
     */
    std::string getPathARelativeToPathB(std::string_view pathA, std::string_view pathB)
    {
        size_t startpos = 0;
        while (startpos < pathA.length() && startpos < pathB.length() && pathA[startpos] == pathB[startpos]) ++startpos;
        while (startpos > 0 and pathA[startpos - 1] != '/') --startpos;
        if (startpos == 0) throw XLInternalError("getPathARelativeToPathB: pathA and pathB have no common beginning");

        std::vector<std::string> dirEntriesB = disassemblePath(pathB.substr(startpos));
        if (!dirEntriesB.empty() && (!pathB.empty() && pathB.back() != '/')) dirEntriesB.pop_back();

        std::string result;
        for (size_t i = 0; i < dirEntriesB.size(); ++i) result += "../";
        result += std::string(pathA.substr(startpos));

        return result;
    }

    /**
     * @details Standardize path by resolving '.' and '..' segments.
     * Rationale: Normalizes paths to ensure consistent internal ZIP archive lookups.
     */
    std::string eliminateDotAndDotDotFromPath(std::string_view path)
    {
        if (path.empty()) return "";
        std::vector<std::string> dirEntries = disassemblePath(path);

        std::string result = (!path.empty() && path.front() == '/') ? "/" : "";
        if (!dirEntries.empty()) {
            result += dirEntries[0];
            for (size_t i = 1; i < dirEntries.size(); ++i) {
                result += "/";
                result += dirEntries[i];
            }
        }

        if (((result.length() > 1 or result.front() != '/') && (!path.empty() && path.back() == '/'))) result += "/";
        return result;
    }

}    // namespace OpenXLSX

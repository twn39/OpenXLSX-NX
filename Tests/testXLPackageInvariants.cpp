/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"

#include <filesystem>
#include <string>

using namespace OpenXLSX;

TEST_CASE("validatePackageInvariants passes on fresh document", "[XLPackageInvariants]")
{
    const std::string path = TestHelpers::getUniqueFilename("inv_ok") + ".xlsx";
    XLDocument        doc;
    doc.create(path, XLForceOverwrite);
    REQUIRE_NOTHROW(doc.validatePackageInvariants());
    doc.save();
    doc.close();
    std::error_code ec;
    std::filesystem::remove(path, ec);
}

TEST_CASE("validatePackageInvariants runs during saveAs", "[XLPackageInvariants]")
{
    const std::string path = TestHelpers::getUniqueFilename("inv_save") + ".xlsx";
    {
        XLDocument doc;
        doc.create(path, XLForceOverwrite);
        doc.workbook().worksheet("Sheet1").cell("A1").value() = 1;
        REQUIRE_NOTHROW(doc.save());
        doc.close();
    }
    {
        XLDocument doc;
        doc.open(path);
        REQUIRE_NOTHROW(doc.validatePackageInvariants());
        doc.close();
    }
    std::error_code ec;
    std::filesystem::remove(path, ec);
}

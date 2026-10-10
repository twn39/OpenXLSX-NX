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

TEST_CASE("XLZipArchive replaced entry immediate readback invariant (addEntry, addEntryAllocated, addEntryFromFile)", "[XLPackageInvariants][XLZipArchive]")
{
    const std::string path = TestHelpers::getUniqueFilename("zip_replace_readback") + ".xlsx";

    // 1. Create a valid baseline archive where entries are saved on disk with standard compression
    {
        XLDocument doc;
        doc.create(path, XLForceOverwrite);
        doc.workbook().worksheet("Sheet1").cell("A1").value() = "baseline";
        doc.save();
        doc.close();
    }

    // 2. Open archive, mutate entries in-place, and verify immediate readback before saving
    {
        XLDocument doc;
        doc.open(path);
        const std::string entryName = "xl/worksheets/sheet1.xml";
        REQUIRE(doc.archive().hasEntry(entryName));

        // Case A: addEntry (std::string) replaces existing compressed entry
        const std::string testXml1 = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>111</v></c></row></sheetData></worksheet>";
        doc.archive().addEntry(entryName, testXml1);
        std::string read1;
        REQUIRE_NOTHROW(read1 = doc.archive().getEntry(entryName));
        REQUIRE(read1 == testXml1);

        // Verify openEntryStream streaming read on the modified entry
        void* stream = doc.archive().openEntryStream(entryName);
        REQUIRE(stream != nullptr);
        std::string streamBuf(testXml1.size(), '\0');
        int64_t bytesRead = doc.archive().readEntryStream(stream, streamBuf.data(), streamBuf.size());
        REQUIRE(bytesRead == static_cast<int64_t>(testXml1.size()));
        REQUIRE(streamBuf == testXml1);
        doc.archive().closeEntryStream(stream);

        // Case B: addEntryAllocated (raw heap pointer) replaces existing entry
        const std::string testXml2 = "<worksheet><sheetData><row r=\"2\"><c r=\"B2\"><v>222</v></c></row></sheetData></worksheet>";
        char* heapBuf = static_cast<char*>(std::malloc(testXml2.size()));
        REQUIRE(heapBuf != nullptr);
        std::memcpy(heapBuf, testXml2.data(), testXml2.size());
        doc.archive().addEntryAllocated(entryName, heapBuf, testXml2.size());
        std::string read2;
        REQUIRE_NOTHROW(read2 = doc.archive().getEntry(entryName));
        REQUIRE(read2 == testXml2);

        // Case C: addEntryFromFile (filesystem path) replaces existing entry
        const std::string tempFile = TestHelpers::getUniqueFilename("temp_sheet_entry") + ".xml";
        const std::string testXml3 = "<worksheet><sheetData><row r=\"3\"><c r=\"C3\"><v>333</v></c></row></sheetData></worksheet>";
        {
            std::FILE* fp = std::fopen(tempFile.c_str(), "wb");
            REQUIRE(fp != nullptr);
            std::fwrite(testXml3.data(), 1, testXml3.size(), fp);
            std::fclose(fp);
        }
        doc.archive().addEntryFromFile(entryName, tempFile);
        // Libzip by design returns ZIP_ER_CHANGED for uncommitted file-backed sources (which do not support re-open before commit)
        REQUIRE_THROWS_AS(doc.archive().getEntry(entryName), XLInternalError);

        // 3. Save modified archive and close cleanly
        REQUIRE_NOTHROW(doc.save());
        doc.close();

        std::error_code ec;
        std::filesystem::remove(tempFile, ec);
    }

    // 4. Reopen and verify data persisted accurately
    {
        XLDocument doc;
        doc.open(path);
        const std::string expectedPersisted = "<worksheet><sheetData><row r=\"3\"><c r=\"C3\"><v>333</v></c></row></sheetData></worksheet>";
        REQUIRE(doc.archive().getEntry("xl/worksheets/sheet1.xml") == expectedPersisted);
        doc.close();
    }

    std::error_code ec;
    std::filesystem::remove(path, ec);
}


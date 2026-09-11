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
#include <catch2/catch_test_macros.hpp>
#include <filesystem>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLIncrementalSaveAndDirtyTracking", "[IncrementalIO]")
{
    const std::string testFile = "incremental_test.xlsx";
    const std::string saveAsFile = "incremental_saveas_test.xlsx";

    // Clean up before run
    std::filesystem::remove(testFile);
    std::filesystem::remove(saveAsFile);

    SECTION("Untouched worksheets remain unmaterialized and untouched during save")
    {
        // 1. Create a workbook with 3 worksheets
        {
            XLDocument doc;
            doc.create(testFile, XLForceOverwrite);
            auto wbk = doc.workbook();
            wbk.addWorksheet("Sheet2");
            wbk.addWorksheet("Sheet3");

            auto ws1 = wbk.worksheet("Sheet1");
            ws1.cell("A1").value() = "Initial Sheet1";

            auto ws2 = wbk.worksheet("Sheet2");
            ws2.cell("A1").value() = "Initial Sheet2";

            auto ws3 = wbk.worksheet("Sheet3");
            ws3.cell("A1").value() = "Initial Sheet3";

            doc.save();
            doc.close();
        }

        // 2. Re-open document and ONLY mutate Sheet1
        {
            XLDocument doc;
            doc.open(testFile);

            // Verify Sheet2 and Sheet3 are not loaded into memory DOM initially via part query
            auto* part2 = doc.findXmlPart("xl/worksheets/sheet2.xml");
            auto* part3 = doc.findXmlPart("xl/worksheets/sheet3.xml");
            REQUIRE(part2 != nullptr);
            REQUIRE(part3 != nullptr);
            REQUIRE_FALSE(part2->isLoaded());
            REQUIRE_FALSE(part3->isLoaded());
            REQUIRE_FALSE(part2->isDirty());
            REQUIRE_FALSE(part3->isDirty());

            // Mutate ONLY Sheet1
            auto ws1 = doc.workbook().worksheet("Sheet1");
            ws1.cell("A1").value() = "Modified Sheet1";
            REQUIRE(ws1.xmlDataPart()->isDirty());

            // Save document
            doc.save();

            // Crucial assertion: Sheet2 and Sheet3 MUST NOT have been loaded into DOM during save!
            REQUIRE_FALSE(part2->isLoaded());
            REQUIRE_FALSE(part3->isLoaded());
            REQUIRE_FALSE(part2->isDirty());
            REQUIRE_FALSE(part3->isDirty());

            // Sheet1 should have been marked clean after save
            REQUIRE_FALSE(ws1.xmlDataPart()->isDirty());

            doc.close();
        }

        // 3. Verify on disk that all sheets still exist and contain correct data
        {
            XLDocument doc;
            doc.open(testFile);
            auto wbk = doc.workbook();
            REQUIRE(wbk.sheetExists("Sheet1"));
            REQUIRE(wbk.sheetExists("Sheet2"));
            REQUIRE(wbk.sheetExists("Sheet3"));

            auto ws1 = wbk.worksheet("Sheet1");
            REQUIRE(ws1.cell("A1").value().get<std::string>() == "Modified Sheet1");

            auto ws2 = wbk.worksheet("Sheet2");
            REQUIRE(ws2.cell("A1").value().get<std::string>() == "Initial Sheet2");

            auto ws3 = wbk.worksheet("Sheet3");
            REQUIRE(ws3.cell("A1").value().get<std::string>() == "Initial Sheet3");

            doc.close();
        }
    }

    SECTION("Unhandled entries (e.g. mock macro or media) are preserved across incremental saveAs")
    {
        const std::string xlsmFile = "incremental_macro.xlsm";
        const std::string xlsmSaveAs = "incremental_macro_saveas.xlsm";
        std::filesystem::remove(xlsmFile);
        std::filesystem::remove(xlsmSaveAs);

        // 1. Create a document with injected unhandled entry and media
        {
            XLDocument doc;
            doc.create(xlsmFile, XLForceOverwrite);
            doc.archive().addEntry("xl/media/image1.png", "fake PNG binary content here 123456");
            doc.archive().addEntry("xl/vbaProject.bin", "fake VBA project stream binary data");
            doc.contentTypes().addOverride("/xl/vbaProject.bin", XLContentType::VBAProject);
            doc.save();
            doc.close();
        }

        // 2. Open and saveAs to a different file
        {
            XLDocument doc;
            doc.open(xlsmFile);
            auto ws1 = doc.workbook().worksheet("Sheet1");
            ws1.cell("B2").value() = "Adding content";
            doc.saveAs(xlsmSaveAs, XLForceOverwrite);
            doc.close();
        }

        // 3. Verify saveAs target archive preserved both unhandled entries perfectly
        {
            XLZipArchive archive;
            archive.open(xlsmSaveAs);
            REQUIRE(archive.hasEntry("xl/media/image1.png"));
            REQUIRE(archive.getEntry("xl/media/image1.png") == "fake PNG binary content here 123456");

            REQUIRE(archive.hasEntry("xl/vbaProject.bin"));
            REQUIRE(archive.getEntry("xl/vbaProject.bin") == "fake VBA project stream binary data");
            archive.close();
        }

        std::filesystem::remove(xlsmFile);
        std::filesystem::remove(xlsmSaveAs);
    }

    SECTION("Compression level can be configured and applied without corrupting untouched entries")
    {
        XLDocument doc;
        doc.create(testFile, XLForceOverwrite);
        doc.setCompressionLevel(9); // Best compression
        REQUIRE(doc.compressionLevel() == 9);

        auto ws1 = doc.workbook().worksheet("Sheet1");
        ws1.cell("A1").value() = "High compression content";
        doc.save();
        doc.close();

        // Check file can be cleanly opened
        XLDocument doc2;
        doc2.open(testFile);
        REQUIRE(doc2.workbook().worksheet("Sheet1").cell("A1").value().get<std::string>() == "High compression content");
        doc2.close();
    }

    // Cleanup
    std::filesystem::remove(testFile);
    std::filesystem::remove(saveAsFile);
}

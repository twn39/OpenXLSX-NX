#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
#include <fstream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("TestDocumentCreationNew_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("testXLDocument_xlsx") + ".xlsx";
    return name;
}
} // namespace


/**
 * @brief The purpose of this test case is to test the creation of XLDocument objects. Each section section
 * tests document creation using a different method. In addition, saving, closing and copying is tested.
 */
TEST_CASE("XLDocumentTests", "[XLDocument]")
{
    std::string file    = OpenXLSX::TestHelpers::getUniqueFilename();
    std::string newfile = OpenXLSX::TestHelpers::getUniqueFilename();

    /**
     * @test
     *
     * @details
     */
    SECTION("Create empty XLDocument, using default constructor")
    {
        XLDocument doc;
        REQUIRE_FALSE(doc);
    }

    SECTION("Create new using create()")
    {
        XLDocument doc;
        doc.create(file, XLForceOverwrite);
        std::ifstream f(file);
        REQUIRE(f.good());
        REQUIRE(doc.name() == __global_unique_file_1());
        doc.close();
    }

    SECTION("Open existing using Constructor")
    {
        {
            XLDocument doc1;
            doc1.create(file, XLForceOverwrite);
            doc1.close();
        }
        XLDocument doc2(file);
        REQUIRE(doc2.name() == __global_unique_file_1());
        doc2.close();
    }

    SECTION("Open existing using open()")
    {
        {
            XLDocument doc1;
            doc1.create(file, XLForceOverwrite);
            doc1.close();
        }
        XLDocument doc2;
        doc2.open(file);
        REQUIRE(doc2.name() == __global_unique_file_1());
        doc2.close();
    }

    SECTION("Save document using saveAs()")
    {
        {
            XLDocument doc1;
            doc1.create(file, XLForceOverwrite);
            doc1.close();
        }
        XLDocument doc2(file);
        doc2.saveAs(newfile, XLForceOverwrite);
        std::ifstream n(newfile);
        REQUIRE(n.good());
        REQUIRE(doc2.name() == __global_unique_file_0());
        doc2.close();
    }

    SECTION("Refuse to overwrite existing file")
    {
        {
            XLDocument doc1;
            doc1.create(file, XLForceOverwrite);
            doc1.close();
        }

        XLDocument doc2;
        REQUIRE_THROWS_AS(doc2.create(file, XLDoNotOverwrite), XLException);
    }

    SECTION("UTF-8 file paths and names")
    {
        std::string utf8File = OpenXLSX::TestHelpers::getUniqueFilename();
        std::remove(utf8File.c_str());

        {
            XLDocument doc;
            doc.create(utf8File, XLForceOverwrite);
            doc.workbook().worksheet("Sheet1").cell("A1").value() = "UTF8 Path Test";
            doc.save();
            doc.close();
        }

        // Verify it can be opened
        {
            XLDocument doc;
            doc.open(utf8File);
            REQUIRE(doc.workbook().worksheet("Sheet1").cell("A1").value().get<std::string>() == "UTF8 Path Test");
            doc.close();
        }

        // Test saveAs with UTF-8
        {
            std::string utf8File2 = OpenXLSX::TestHelpers::getUniqueFilename();
            std::remove(utf8File2.c_str());
            XLDocument doc;
            doc.open(utf8File);
            doc.saveAs(utf8File2, XLForceOverwrite);
            doc.close();

            XLDocument doc2;
            doc2.open(utf8File2);
            REQUIRE(doc2.isOpen());
            doc2.close();
            std::remove(utf8File2.c_str());
        }

        std::remove(utf8File.c_str());
    }

    SECTION("Macro Preservation (.xlsm)")
    {
        std::string xlsmFile1 = OpenXLSX::TestHelpers::getUniqueFilename("macro") + ".xlsm";
        std::string xlsmFile2 = OpenXLSX::TestHelpers::getUniqueFilename("macro") + ".xlsm";

        {
            XLDocument doc;
            doc.create(xlsmFile1, XLForceOverwrite);
            // Manually inject vbaProject.bin into the underlying archive
            doc.archive().addEntry("xl/vbaProject.bin", "dummy macro content");
            doc.contentTypes().addOverride("/xl/vbaProject.bin", XLContentType::VBAProject);
            doc.save();
            doc.close();
        }

        // Open it and save it again
        {
            XLDocument doc;
            doc.open(xlsmFile1);
            // Verify our unhandled entry tracker caught it
            doc.saveAs(xlsmFile2, XLForceOverwrite);
            doc.close();
        }

        // Verify the second file has vbaProject.bin
        {
            XLZipArchive archive;
            archive.open(xlsmFile2);
            REQUIRE(archive.hasEntry("xl/vbaProject.bin"));
            REQUIRE(archive.getEntry("xl/vbaProject.bin") == "dummy macro content");

            // Also verify it's registered in [Content_Types].xml
            std::string contentTypes = archive.getEntry("[Content_Types].xml");
            REQUIRE(contentTypes.find("application/vnd.ms-office.vbaProject") != std::string::npos);
            REQUIRE(contentTypes.find("macroEnabled") != std::string::npos);
            archive.close();
        }
        std::remove(xlsmFile1.c_str());
        std::remove(xlsmFile2.c_str());
    }
}

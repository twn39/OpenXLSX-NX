#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
#include <fstream>
#include <vector>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("non_existent_random_file_123_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("NegativeTestsCorruptFilesandExceptions", "[CorruptFiles]")
{
    SECTION("File Not Found")
    {
        XLDocument doc;
        REQUIRE_THROWS_AS(doc.open(__global_unique_file_0()), XLException);
    }

    SECTION("Not a ZIP Archive")
    {
        const std::string dummy_file = OpenXLSX::TestHelpers::getUniqueFilename();
        std::ofstream     out(dummy_file);
        out << "This is just a text file, not a valid ZIP or XLSX file.";
        out.close();

        XLDocument doc;
        REQUIRE_THROWS_AS(doc.open(dummy_file), XLException);

        std::remove(dummy_file.c_str());
    }

    SECTION("Valid ZIP but missing Content_Types.xml")
    {
        const std::string testFile = OpenXLSX::TestHelpers::getUniqueFilename();
        {
            XLZipArchive archive;
            archive.open(testFile);
            archive.addEntry("dummy.txt", "hello");
            archive.save();
            archive.close();
        }
        
        XLDocument doc;
        REQUIRE_THROWS_AS(doc.open(testFile), XLException);
        std::remove(testFile.c_str());
    }
SECTION("Valid ZIP but malformed XML inside")
{
    // 1. Create a valid document first
    const std::string malformed_zip_file = OpenXLSX::TestHelpers::getUniqueFilename();
    {
        XLDocument doc;
        doc.create(malformed_zip_file, XLForceOverwrite);
        doc.save();
        doc.close();
    }

    // 2. We inject bad xml into [Content_Types].xml natively
    {
        XLZipArchive archive;
        archive.open(malformed_zip_file);
        if (archive.hasEntry("[Content_Types].xml")) {
            archive.deleteEntry("[Content_Types].xml");
        }
        std::string badXml = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ";
        archive.addEntry("[Content_Types].xml", badXml);
        archive.save();
        archive.close();
    }

    // 3. Trying to open it should cause pugi to throw, caught by OpenXLSX
    XLDocument doc;
    REQUIRE_THROWS_AS(doc.open(malformed_zip_file), XLException);

    std::remove(malformed_zip_file.c_str());
}
}

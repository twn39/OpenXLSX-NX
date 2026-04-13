#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <fstream>
#include <filesystem>

using namespace OpenXLSX;

TEST_CASE("FileSystem Failure Simulation and Negative Tests", "[FileSystem][Negative]")
{
    SECTION("Save to restricted path or invalid paths")
    {
        XLDocument doc;
        doc.create("normal_file.xlsx", XLForceOverwrite);
        doc.workbook().worksheet("Sheet1").cell("A1").value() = "Test";

#if defined(__unix__) || defined(__APPLE__)
        // Attempt to save to a root/system protected path.
        // Even if running as root, /root/ or /System/ should fail if directories don't exist
        // To be absolutely sure it fails everywhere, we use an invalid deep path instead.
        REQUIRE_THROWS_AS(doc.saveAs("/dev/null/forbidden_save_test.xlsx", XLForceOverwrite), XLException);
#elif defined(_WIN32)
        // On Windows GitHub Actions, the runner often runs as Administrator, so writing to C:\Windows might actually succeed!
        // Instead, we try to write to a reserved DOS device name or an invalid drive to guarantee a filesystem error.
        REQUIRE_THROWS_AS(doc.saveAs("Q:\\ImpossibleDrive\\forbidden_save_test.xlsx", XLForceOverwrite), XLException);
        REQUIRE_THROWS_AS(doc.saveAs("CON:\\forbidden_save_test.xlsx", XLForceOverwrite), XLException);
#endif
        doc.close();
        std::remove("normal_file.xlsx");
    }

    SECTION("Save to non-existent drive or deep directory")
    {
        XLDocument doc;
        doc.create("normal_file2.xlsx", XLForceOverwrite);
        doc.workbook().worksheet("Sheet1").cell("A1").value() = "Test";

        // Directory structure does not exist
        REQUIRE_THROWS_AS(doc.saveAs("/path/that/absolutely/does/not/exist/file.xlsx", XLForceOverwrite), XLException);
        
        doc.close();
        std::remove("normal_file2.xlsx");
    }

    SECTION("Corrupt internal ZIP structure during open (Bad ZIP)")
    {
        // Create a fake invalid zip file manually using fstream
        const std::string fakeZip = "not_a_real_zip_file.xlsx";
        std::ofstream out(fakeZip, std::ios::binary);
        out << "PK\x03\x04" << "This is definitely not a valid zip payload";
        out.close();

        XLDocument doc;
        // libzip should fail to parse this and OpenXLSX should catch it and throw XLException
        REQUIRE_THROWS_AS(doc.open(fakeZip), XLException);

        std::remove(fakeZip.c_str());
    }

    SECTION("Invalid Document Creation (Empty Path)")
    {
        XLDocument doc;
        // Creating an empty path should throw
        REQUIRE_THROWS_AS(doc.create("", XLForceOverwrite), XLException);
    }
}

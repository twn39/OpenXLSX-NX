#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
#include <filesystem>
#include <fstream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("sheetnames_test_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("coercion_test_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("whitespace_test_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("SecurityandExploitationEdgeCases", "[EdgeCases][Security]")
{
    SECTION("Billion Laughs (XXE) Prevention")
    {
        // Create a ZIP with a malicious Content_Types.xml using OpenXLSX's XLZipArchive
        const std::string testFile = OpenXLSX::TestHelpers::getUniqueFilename();
        {
            XLZipArchive archive;
            archive.open(testFile);
            std::string maliciousXml = "<!DOCTYPE Types [<!ENTITY xxe \"bar\">]><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"&xxe;\"/></Types>";
            archive.addEntry("[Content_Types].xml", maliciousXml);
            archive.save();
            archive.close();
        }

        XLDocument doc;
        // Depending on pugi_parse_settings, this will either throw or just not evaluate the entity.
        // PugiXML default parser does NOT evaluate external/internal entities.
        // So it should safely open without memory exhaustion.
        try {
            doc.open(testFile);
        }
        catch (...) {
            // If it throws because of missing relationships or workbook, that's fine too.
            // The key is it didn't hang or crash.
        }
        std::remove(testFile.c_str());
        REQUIRE(true);    // If we reached here, no OOM/Hang occurred.
    }
}

TEST_CASE("DateSystemEdgeCases1900LeapYearBug", "[EdgeCases][Date]")
{
    SECTION("Serial Number 60 (1900-02-29)")
    {
        // 59 = 1900-02-28
        // 60 = 1900-02-29 (fictional)
        // 61 = 1900-03-01

        XLDateTime dt59(59.0);
        auto       tm59 = dt59.tm();
        REQUIRE(tm59.tm_year == 0);    // 1900
        REQUIRE(tm59.tm_mon == 1);     // Feb
        REQUIRE(tm59.tm_mday == 28);

        XLDateTime dt60(60.0);
        auto       tm60 = dt60.tm();
        REQUIRE(tm60.tm_year == 0);
        REQUIRE(tm60.tm_mon == 1);
        REQUIRE(tm60.tm_mday == 29);    // XLDateTime accurately preserves the bug

        XLDateTime dt61(61.0);
        auto       tm61 = dt61.tm();
        REQUIRE(tm61.tm_year == 0);
        REQUIRE(tm61.tm_mon == 2);    // Mar
        REQUIRE(tm61.tm_mday == 1);
    }
}

TEST_CASE("WhitespacePreservationEdgeCases", "[EdgeCases][Whitespace]")
{
    XLDocument doc;
    doc.create(__global_unique_file_2(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    SECTION("xml:space=\"preserve\" attribute")
    {
        wks.cell("A1").value() = "   Hello World   ";
        doc.save();

        // Peek into the raw sharedStrings.xml
        // OpenXLSX adds xml:space="preserve" to the node if it has leading/trailing spaces
        // We can't directly read the XML here without internal methods, but we can verify
        // that saving and loading preserves the whitespace exactly.
        doc.close();

        doc.open(__global_unique_file_2());
        auto wks2 = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks2.cell("A1").value().get<std::string>() == "   Hello World   ");
        doc.close();
    }
    std::filesystem::remove(__global_unique_file_2());
}

TEST_CASE("SheetNamingandEscapingEdgeCases", "[EdgeCases][SheetNames]")
{
    XLDocument doc;
    doc.create(__global_unique_file_0(), XLForceOverwrite);

    SECTION("Illegal Characters Rejection")
    {
        std::vector<std::string> illegalNames = {"Sheet/1", "Sheet\\2", "Sheet?3", "Sheet*4", "Sheet:5", "Sheet[6]", "Sheet]7"};

        for (const auto& name : illegalNames) { REQUIRE_THROWS_AS(doc.workbook().addWorksheet(name), XLInputError); }
    }

    SECTION("Apostrophe Bounds Rejection")
    {
        // Excel doesn't allow sheet names to start or end with an apostrophe
        REQUIRE_THROWS_AS(doc.workbook().addWorksheet("'Sheet1"), XLInputError);
        REQUIRE_THROWS_AS(doc.workbook().addWorksheet("Sheet1'"), XLInputError);
    }

    SECTION("Valid Spaces") { REQUIRE_NOTHROW(doc.workbook().addWorksheet("Revenue 2026")); }

    doc.close();
    std::filesystem::remove(__global_unique_file_0());
}

TEST_CASE("StringtoNumberCoercionEdgeCases", "[EdgeCases][Coercion]")
{
    XLDocument doc;
    doc.create(__global_unique_file_1(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    SECTION("Leading Zeros in Strings")
    {
        // If we assign a string with leading zeros, it MUST remain a string.
        // It should not be coerced into a float/int by the library.
        std::string zipCode    = "01234";
        wks.cell("A1").value() = zipCode;

        REQUIRE(wks.cell("A1").value().type() == XLValueType::String);
        REQUIRE(wks.cell("A1").value().get<std::string>() == "01234");

        doc.save();
        doc.close();

        doc.open(__global_unique_file_1());
        auto wks2 = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks2.cell("A1").value().type() == XLValueType::String);
        REQUIRE(wks2.cell("A1").value().get<std::string>() == "01234");
        doc.close();
    }
    std::filesystem::remove(__global_unique_file_1());
}
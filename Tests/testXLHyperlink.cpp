#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
#include <pugixml.hpp>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testHyperlinkCRUD_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testInternalLink_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testHyperlink_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("XLWorksheetHyperlinkSupport", "[XLWorksheet]")
{
    SECTION("Add External Hyperlink")
    {
        XLDocument doc;
        doc.create(__global_unique_file_2(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value() = "Google";
        wks.addHyperlink("A1", "https://www.google.com", "Click to visit Google");

        doc.save();
        doc.close();

        // Verification
        XLDocument doc2;
        doc2.open(__global_unique_file_2());
        auto wks2 = doc2.workbook().worksheet("Sheet1");

        pugi::xml_document sheetDoc;
        sheetDoc.load_string(wks2.xmlData().c_str());
        auto hyperlinks = sheetDoc.child("worksheet").child("hyperlinks");
        REQUIRE_FALSE(hyperlinks.empty());

        auto link = hyperlinks.child("hyperlink");
        REQUIRE(std::string(link.attribute("ref").value()) == "A1");
        REQUIRE_FALSE(link.attribute("r:id").empty());
        REQUIRE(std::string(link.attribute("tooltip").value()) == "Click to visit Google");

        doc2.close();
    }

    SECTION("Add Internal Hyperlink")
    {
        XLDocument doc;
        doc.create(__global_unique_file_1(), XLForceOverwrite);
        doc.workbook().addWorksheet("TargetSheet");
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("B2").value() = "Go to TargetSheet";
        wks.addInternalHyperlink("B2", "TargetSheet!A1", "Jump to other sheet");

        doc.save();
        doc.close();

        // Verification
        XLDocument doc2;
        doc2.open(__global_unique_file_1());
        auto wks2 = doc2.workbook().worksheet("Sheet1");

        pugi::xml_document sheetDoc;
        sheetDoc.load_string(wks2.xmlData().c_str());
        auto link = sheetDoc.child("worksheet").child("hyperlinks").child("hyperlink");
        REQUIRE(std::string(link.attribute("ref").value()) == "B2");
        REQUIRE(std::string(link.attribute("location").value()) == "TargetSheet!A1");

        doc2.close();
    }

    SECTION("Hyperlink CRUD Operations")
    {
        XLDocument doc;
        doc.create(__global_unique_file_0(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // 1. Initial state
        REQUIRE_FALSE(wks.hasHyperlink("A1"));
        REQUIRE(wks.getHyperlink("A1") == "");

        // 2. Add and check
        wks.addHyperlink("A1", "https://www.google.com");
        REQUIRE(wks.hasHyperlink("A1"));
        REQUIRE(wks.getHyperlink("A1") == "https://www.google.com");

        // 3. Update and check
        wks.addHyperlink("A1", "https://www.github.com");
        REQUIRE(wks.hasHyperlink("A1"));
        REQUIRE(wks.getHyperlink("A1") == "https://www.github.com");

        // 4. Internal link
        wks.addInternalHyperlink("B2", "Sheet1!A10");
        REQUIRE(wks.hasHyperlink("B2"));
        REQUIRE(wks.getHyperlink("B2") == "Sheet1!A10");

        // 5. Remove
        wks.removeHyperlink("A1");
        REQUIRE_FALSE(wks.hasHyperlink("A1"));
        REQUIRE(wks.getHyperlink("A1") == "");

        doc.save();
        doc.close();
    }
}

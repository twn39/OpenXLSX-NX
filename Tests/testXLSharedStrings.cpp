#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>

using namespace OpenXLSX;

TEST_CASE("XLSharedStrings Tests", "[XLSharedStrings]")
{
    SECTION("Basic String Operations")
    {
        XLDocument doc;
        doc.create("./testXLSharedStrings.xlsx", XLForceOverwrite);

        // Shared strings are managed via the document but can be accessed
        auto& ss = doc.sharedStrings();

        int32_t initialCount = ss.stringCount();
        int32_t idx1         = ss.getOrCreateStringIndex("Hello");
        int32_t idx2         = ss.getOrCreateStringIndex("World");
        int32_t idx3         = ss.getOrCreateStringIndex("Hello");    // Should be same as idx1

        REQUIRE(idx1 == initialCount);
        REQUIRE(idx2 == initialCount + 1);
        REQUIRE(idx3 == idx1);
        REQUIRE(ss.stringCount() == initialCount + 2);

        REQUIRE(std::string(ss.getString(idx1)) == "Hello");
        REQUIRE(std::string(ss.getString(idx2)) == "World");

        REQUIRE(ss.stringExists("Hello") == true);
        REQUIRE(ss.stringExists("NonExistent") == false);

        ss.clearString(idx1);
        REQUIRE(std::string(ss.getString(idx1)) == "");
        // Clearing doesn't remove the entry to preserve indices
        REQUIRE(ss.stringCount() == initialCount + 2);

        doc.close();
    }

    SECTION("Large number of strings")
    {
        XLDocument doc;
        doc.create("./testXLSharedStringsLarge.xlsx", XLForceOverwrite);
        auto&   ss           = doc.sharedStrings();
        int32_t initialCount = ss.stringCount();

        for (int i = 0; i < 100; ++i) { ss.getOrCreateStringIndex("String " + std::to_string(i)); }

        REQUIRE(ss.stringCount() == initialCount + 100);
        REQUIRE(std::string(ss.getString(initialCount + 99)) == "String 99");

        doc.close();
    }
}

TEST_CASE("Shared Strings Index Cache Coherency", "[SharedStrings][Bugfix]") {
    XLDocument doc;
    doc.create("SharedStrings_Cache_Test.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    wks.cell("A1").value() = "Hello World";
    auto ss = doc.sharedStrings();
    int32_t idx = ss.getStringIndex("Hello World");
    REQUIRE(idx >= 0);
    REQUIRE(ss.stringExists("Hello World") == true);
    ss.clearString(idx);
    REQUIRE(ss.stringExists("Hello World") == false);
    REQUIRE(ss.getStringIndex("Hello World") == -1);
    doc.save();
    doc.close();
}

TEST_CASE("SharedStrings Lazy DOM and Reservation", "[XLSharedStrings]")
{
    SECTION("reserveStrings pre-allocates capacity")
    {
        XLDocument doc;
        doc.create("./testXLSharedStrings_reserve.xlsx", XLForceOverwrite);
        auto& ss = doc.sharedStrings();

        constexpr int kCount = 1000;
        ss.reserveStrings(kCount);

        for (int i = 0; i < kCount; ++i) {
            ss.getOrCreateStringIndex("Key " + std::to_string(i));
        }

        REQUIRE(ss.stringCount() == kCount);
        REQUIRE(std::string(ss.getString(0)) == "Key 0");
        REQUIRE(std::string(ss.getString(kCount - 1)) == "Key " + std::to_string(kCount - 1));

        doc.close();
    }

    SECTION("memoryUsageBytes returns non-zero value after strings are added")
    {
        XLDocument doc;
        doc.create("./testXLSharedStrings_mem.xlsx", XLForceOverwrite);
        auto& ss = doc.sharedStrings();

        ss.getOrCreateStringIndex("alpha");
        ss.getOrCreateStringIndex("beta");

        REQUIRE(ss.memoryUsageBytes() > 0);
        doc.close();
    }

    SECTION("Lazy DOM: strings survive save/reopen cycle")
    {
        // The lazy DOM path does NOT write to the pugi DOM on each appendString.
        // rewriteXmlFromCache() is called at save time and must produce a correct
        // shared strings XML file.
        {
            XLDocument doc;
            doc.create("./testXLSharedStrings_lazy.xlsx", XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // Write strings via the cell API (goes through appendString/lazy DOM)
            for (int i = 0; i < 50; ++i) {
                wks.cell(1, i + 1).value() = "Cell " + std::to_string(i);
            }

            doc.save();
            doc.close();
        }

        // Reopen and verify every string is correctly preserved
        XLDocument doc;
        doc.open("./testXLSharedStrings_lazy.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        for (int i = 0; i < 50; ++i) {
            const std::string expected = "Cell " + std::to_string(i);
            REQUIRE(wks.cell(1, i + 1).value().get<std::string>() == expected);
        }
        doc.close();
    }
}


#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"

using namespace OpenXLSX;

TEST_CASE("CryptoDecryptionTest") {
    SECTION("Standard Decryption Excelize File") {
        XLDocument doc;
        REQUIRE_NOTHROW(doc.open("Tests/Fixtures/Encrypted_Agile.xlsx", "OpenXLSX2026"));
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell("A1").value().getString() == "This is an Agile Encrypted Document.");
        REQUIRE(wks.cell("A2").value().getString() == "Password is: OpenXLSX2026");
    }
}

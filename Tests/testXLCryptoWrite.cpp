
#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"

using namespace OpenXLSX;

TEST_CASE("Crypto Encryption Test") {
    SECTION("Standard Encryption Export") {
        XLDocument doc;
        doc.create("Tests/Fixtures/OpenXLSX_Encrypted_Test.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        wks.cell("A1").value() = "This file was encrypted by OpenXLSX C++ Engine";
        wks.cell("B1").value() = 42;
        REQUIRE_NOTHROW(doc.saveAs("Tests/Fixtures/OpenXLSX_Encrypted_Test.xlsx", "SecretPass123"));
        
        XLDocument doc2;
        REQUIRE_NOTHROW(doc2.open("Tests/Fixtures/OpenXLSX_Encrypted_Test.xlsx", "SecretPass123"));
        REQUIRE(doc2.workbook().worksheet("Sheet1").cell("A1").value().getString() == "This file was encrypted by OpenXLSX C++ Engine");
    }
}

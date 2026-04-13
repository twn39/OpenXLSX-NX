#include "XLCrypto.hpp"
#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <fstream>
#include <vector>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__OpenXLSX_Encrypted_Cycle_Test_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__OpenXLSX_Wrong_Pass_Test_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__OpenXLSX_Long_Pass_Test_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_3() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__OpenXLSX_Empty_Pass_Test_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("CryptoEncryptionandDecryptionCycle") {
    SECTION("Standard Encryption Export and Readback") {
        XLDocument doc;
        doc.create(__global_unique_file_0(), XLForceOverwrite);
        
        auto wks = doc.workbook().worksheet("Sheet1");
        wks.cell("A1").value() = "This file was encrypted by OpenXLSX C++ Engine";
        wks.cell("B1").value() = 42;
        wks.cell("C1").formula() = "B1*2";
        
        // Save the document with a password
        REQUIRE_NOTHROW(doc.saveAs(__global_unique_file_0(), std::string("CyclePass123!"), XLForceOverwrite));
        
        // Verify it can be opened back
        XLDocument doc2;
        REQUIRE_NOTHROW(doc2.open(__global_unique_file_0(), "CyclePass123!"));
        
        // Verify the contents
        auto wks2 = doc2.workbook().worksheet("Sheet1");
        REQUIRE(wks2.cell("A1").value().getString() == "This file was encrypted by OpenXLSX C++ Engine");
        REQUIRE(wks2.cell("B1").value().get<int>() == 42);
        REQUIRE(wks2.cell("C1").formula().get() == "B1*2");
    }

    SECTION("Invalid Password Handling") {
        XLDocument doc;
        doc.create(__global_unique_file_1(), XLForceOverwrite);
        doc.workbook().worksheet("Sheet1").cell("A1").value() = "Secret Data";
        doc.saveAs(__global_unique_file_1(), std::string("CorrectPass"), XLForceOverwrite);

        XLDocument doc2;
        // Opening with the wrong password should throw an XLInternalError (due to HMAC/Verifier mismatch or ZIP signature mismatch)
        REQUIRE_THROWS_AS(doc2.open(__global_unique_file_1(), "WrongPass"), XLInternalError);
    }

    SECTION("Empty Password Handling") {
        XLDocument doc;
        doc.create(__global_unique_file_3(), XLForceOverwrite);
        doc.workbook().worksheet("Sheet1").cell("A1").value() = "Empty Password Data";
        
        // Some systems allow empty string passwords for encryption.
        // It should either throw an XLInternalError if blocked, or succeed.
        REQUIRE_NOTHROW(doc.saveAs(__global_unique_file_3(), std::string(""), XLForceOverwrite));

        XLDocument doc2;
        REQUIRE_NOTHROW(doc2.open(__global_unique_file_3(), ""));
        REQUIRE(doc2.workbook().worksheet("Sheet1").cell("A1").value().getString() == "Empty Password Data");
    }

    SECTION("Extremely Long Password Handling") {
        XLDocument doc;
        doc.create(__global_unique_file_2(), XLForceOverwrite);
        doc.workbook().worksheet("Sheet1").cell("A1").value() = "Long Password Data";
        
        std::string longPass(300, 'A');
        
        // If the implementation doesn't strictly block > 255 chars, it should at least not crash,
        // and ideally round-trip successfully.
        REQUIRE_NOTHROW(doc.saveAs(__global_unique_file_2(), longPass, XLForceOverwrite));

        XLDocument doc2;
        REQUIRE_NOTHROW(doc2.open(__global_unique_file_2(), longPass));
        REQUIRE(doc2.workbook().worksheet("Sheet1").cell("A1").value().getString() == "Long Password Data");
    }
}


#include "OpenXLSX.hpp"
#include "XLCrypto.hpp"
#include <catch2/catch_test_macros.hpp>
#include <cstdint>
#include <string>
#include <vector>

using namespace OpenXLSX;

TEST_CASE("Crypto OLE Header Detection", "[XLCrypto]")
{
    SECTION("Valid OLE Header")
    {
        std::vector<uint8_t> validOle = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00, 0x00};
        validOle.resize(512, 0);    // Need to be at least 512 bytes for OLE parsing usually
        REQUIRE(isEncryptedDocument(validOle) == true);
    }

    SECTION("Valid ZIP Header")
    {
        std::vector<uint8_t> validZip = {0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00};
        REQUIRE(isEncryptedDocument(validZip) == false);
    }

    SECTION("Empty and short boundaries")
    {
        std::vector<uint8_t> empty;
        REQUIRE(isEncryptedDocument(empty) == false);

        std::vector<uint8_t> shortData = {0xD0, 0xCF, 0x11};
        REQUIRE(isEncryptedDocument(shortData) == false);
    }
}

TEST_CASE("Crypto Hash Algorithms", "[XLCrypto]")
{
    SECTION("SHA-512 Hash Generation")
    {
        std::string          input = "OpenXLSX";
        std::vector<uint8_t> data(input.begin(), input.end());
        auto                 hash = Crypto::sha512Hash(data);

        REQUIRE(hash.size() == 64);

        // Expected SHA-512 for "OpenXLSX"
        std::vector<uint8_t> expected = {0xb8, 0x13, 0x79, 0x6d, 0x53, 0x8b, 0x52, 0xca, 0x83, 0x5b, 0x35, 0x09, 0xbd, 0xdb, 0x41, 0x05,
                                         0x7b, 0xaa, 0x5c, 0xbb, 0x91, 0xf4, 0x05, 0x66, 0x1f, 0xff, 0x96, 0xca, 0xec, 0xa3, 0xbd, 0x62,
                                         0xf4, 0xf1, 0x3a, 0xec, 0xae, 0x6f, 0x79, 0xcc, 0xd6, 0xb4, 0xb9, 0x12, 0xd5, 0x06, 0xe3, 0x96,
                                         0x87, 0x62, 0xe5, 0x81, 0x15, 0x25, 0xcd, 0xd2, 0xe9, 0x1c, 0x52, 0xea, 0x57, 0x5c, 0xf1, 0x7f};
        REQUIRE(hash == expected);
    }

    SECTION("Agile Key Generation (PBKDF2 variant)")
    {
        std::string          password = "password";
        std::vector<uint8_t> salt(16, 0xAA);
        int                  spinCount = 10;

        auto derivedKey = Crypto::generateAgileHash(password, salt, spinCount);
        REQUIRE(derivedKey.size() == 64);
        // It's deterministic. A non-empty output implies the pipeline didn't crash.
        REQUIRE(derivedKey[0] != 0x00);
    }
}

TEST_CASE("Crypto AES Decryption", "[XLCrypto]")
{
    SECTION("AES-256-CBC Decryption with PKCS#7 padding")
    {
        // Plaintext: "Test" (4 bytes) + 12 bytes of padding (0x0C)
        // Key: 32 bytes of 0x00, IV: 16 bytes of 0x00
        std::vector<uint8_t> zeroKey(32, 0x00);
        std::vector<uint8_t> zeroIv(16, 0x00);
        std::vector<uint8_t> cipherText = {0x0c, 0x3e, 0xc9, 0x66, 0x66, 0xd5, 0x56, 0x38, 0x62, 0xef, 0xe5, 0xf9, 0xb7, 0xeb, 0xa7, 0x4f};

        auto plaintext = Crypto::aes256CbcDecrypt(cipherText, zeroKey, zeroIv);
        REQUIRE(plaintext.size() == 4);
        REQUIRE(plaintext[0] == 'T');
        REQUIRE(plaintext[1] == 'e');
        REQUIRE(plaintext[2] == 's');
        REQUIRE(plaintext[3] == 't');
    }

    SECTION("AES-256-CBC Decryption Invalid Padding")
    {
        std::vector<uint8_t> zeroKey(32, 0x00);
        std::vector<uint8_t> zeroIv(16, 0x00);
        std::vector<uint8_t> badCiphertext = {
            0x00,
            0x01,
            0x02    // not a multiple of 16
        };
        auto plaintext = Crypto::aes256CbcDecrypt(badCiphertext, zeroKey, zeroIv);
        REQUIRE(plaintext.empty());    // Should fail gracefully
    }
}

TEST_CASE("Crypto Document Security", "[XLCrypto]")
{
    SECTION("Malformed OLE CFB Container")
    {
        // Valid OLE header but truncated FAT
        std::vector<uint8_t> badOle = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1,    // Magic
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x3E, 0x00, 0x03, 0x00, 0xFE, 0xFF, 0x09, 0x00, 0x06, 0x00,    // Sector shift 9, mini sector shift 6
                                       0x00, 0x00, 0x00, 0x00};
        REQUIRE_THROWS_AS(decryptDocument(badOle, "password"), XLInternalError);
    }
}

TEST_CASE("CryptoDecryptionTest")
{
    SECTION("Standard Decryption Excelize File")
    {
        XLDocument doc;
        REQUIRE_NOTHROW(doc.open("Tests/Fixtures/Encrypted_Agile.xlsx", "OpenXLSX2026"));
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell("A1").value().getString() == "This is an Agile Encrypted Document.");
        REQUIRE(wks.cell("A2").value().getString() == "Password is: OpenXLSX2026");
    }
}

#ifndef OPENXLSX_XLCRYPTO_HPP
#define OPENXLSX_XLCRYPTO_HPP

#pragma once

#include <vector>
#include <string>
#include <cstdint>
#include "OpenXLSX-Exports.hpp"
#include "XLException.hpp"
#include <gsl/span>

namespace OpenXLSX {

    /**
     * @brief Checks if a given byte stream represents an OLE CFB document.
     * @param data The file data.
     * @return true if it is an encrypted/OLE document.
     */
    OPENXLSX_EXPORT bool isEncryptedDocument(gsl::span<const uint8_t> data);

    /**
     * @brief High-level API to decrypt an entire document from a buffer.
     * @param data The CFB file data.
     * @param password The user password.
     * @return The decrypted ZIP package.
     */
    OPENXLSX_EXPORT std::vector<uint8_t> decryptDocument(gsl::span<const uint8_t> data, const std::string& password);

    /**
     * @brief High-level API to encrypt a ZIP package into a CFB container.
     * @param zipData The raw ZIP data.
     * @param password The user password.
     * @return The encrypted CFB OLE container.
     */
    OPENXLSX_EXPORT std::vector<uint8_t> encryptDocument(gsl::span<const uint8_t> zipData, const std::string& password);

    // Advanced Crypto routines for internal use
    namespace Crypto {
        std::vector<uint8_t> readCfbStream(gsl::span<const uint8_t> data, const std::string& streamName);
        std::vector<uint8_t> decryptAgilePackage(gsl::span<const uint8_t> encryptionInfo,
                                                 gsl::span<const uint8_t> encryptedPackage,
                                                 const std::string& password);
        std::vector<uint8_t> encryptStandardPackage(gsl::span<const uint8_t> zipData, const std::string& password);
        std::vector<uint8_t> decryptStandardPackage(gsl::span<const uint8_t> encryptionInfo,
                                                    gsl::span<const uint8_t> encryptedPackage,
                                                    const std::string& password);
        
        std::vector<uint8_t> aes256CbcDecrypt(gsl::span<const uint8_t> data, gsl::span<const uint8_t> key, gsl::span<const uint8_t> iv);
        std::vector<uint8_t> sha512Hash(gsl::span<const uint8_t> data);
        std::vector<uint8_t> generateAgileHash(const std::string& password, gsl::span<const uint8_t> salt, int spinCount);
    }
} // namespace OpenXLSX

#endif // OPENXLSX_XLCRYPTO_HPP

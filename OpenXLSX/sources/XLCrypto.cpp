#include <map>
#include "XLCrypto.hpp"
#include <mbedtls/aes.h>
#include <mbedtls/sha1.h>
#include <mbedtls/sha512.h>
#include <mbedtls/ctr_drbg.h>
#include <mbedtls/entropy.h>
#include "XLException.hpp"
#include <cstring>
#include <memory>
#include <array>
#include <pugixml.hpp>
#include <algorithm>
#include <iomanip>

using namespace OpenXLSX;

bool OpenXLSX::isEncryptedDocument(gsl::span<const uint8_t> data) {
    if (data.size() < 512) return false;
    const uint8_t magic[] = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
    return std::memcmp(data.data(), magic, 8) == 0;
}

std::vector<uint8_t> OpenXLSX::encryptDocument(gsl::span<const uint8_t> zipData, const std::string& password) {
    return Crypto::encryptStandardPackage(zipData, password);
}

std::vector<uint8_t> OpenXLSX::decryptDocument(gsl::span<const uint8_t> data, const std::string& password) {
    auto encryptionInfo = Crypto::readCfbStream(data, "EncryptionInfo");
    auto encryptedPackage = Crypto::readCfbStream(data, "EncryptedPackage");
    
    if (encryptionInfo.size() >= 4) {
        uint32_t version = encryptionInfo[0] | (encryptionInfo[1]<<8) | (encryptionInfo[2]<<16) | (encryptionInfo[3]<<24);
        if (version == 0x00020003) {
            return Crypto::decryptStandardPackage(encryptionInfo, encryptedPackage, password);
        } else if (version == 0x00040004) {
            return Crypto::decryptAgilePackage(encryptionInfo, encryptedPackage, password);
        }
    }
    
    throw XLInternalError("Unknown encryption version");
}

namespace OpenXLSX::Crypto {

std::vector<uint8_t> readCfbStream(gsl::span<const uint8_t> data, const std::string& streamName) {
    if (data.size() < 512) throw XLInternalError("Invalid CFB size");
    const uint8_t magic[] = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
    if (std::memcmp(data.data(), magic, 8) != 0) throw XLInternalError("Not an OLE/CFB file");

    auto getU32 = [](const uint8_t* p) { return p[0] | (p[1]<<8) | (p[2]<<16) | (p[3]<<24); };
    auto getU16 = [](const uint8_t* p) { return p[0] | (p[1]<<8); };
    const uint32_t ENDOFCHAIN = 0xFFFFFFFE;
    const uint32_t FREESECT = 0xFFFFFFFF;

    uint16_t sectorShift = getU16(data.data() + 0x1E);
    uint32_t sectorSize = 1 << sectorShift;
    
    std::vector<uint32_t> difat;
    for(int i=0; i<109; ++i) {
        uint32_t sec = getU32(data.data() + 0x4C + i*4);
        if (sec != FREESECT) difat.push_back(sec);
    }
    
    std::vector<uint32_t> fat;
    for(uint32_t sec : difat) {
        if ((sec + 1) * sectorSize + sectorSize > data.size()) break;
        const uint8_t* p = data.data() + (sec + 1) * sectorSize;
        for(uint32_t i=0; i<sectorSize/4; ++i) fat.push_back(getU32(p + i*4));
    }
    
    auto getChain = [&](uint32_t start) {
        std::vector<uint32_t> chain;
        uint32_t current = start;
        while (current != ENDOFCHAIN && current < fat.size()) {
            chain.push_back(current);
            current = fat[current];
        }
        return chain;
    };

    uint32_t firstMiniFatSector = getU32(data.data() + 0x3C);
    std::vector<uint32_t> minifat;
    if (firstMiniFatSector != ENDOFCHAIN) {
        auto chain = getChain(firstMiniFatSector);
        for(uint32_t sec : chain) {
            if ((sec + 1) * sectorSize + sectorSize > data.size()) break;
            const uint8_t* p = data.data() + (sec + 1) * sectorSize;
            for(uint32_t i=0; i<sectorSize/4; ++i) minifat.push_back(getU32(p + i*4));
        }
    }
    
    uint32_t firstDirSector = getU32(data.data() + 0x30);
    auto dirChain = getChain(firstDirSector);
    uint32_t rootStartSector = 0;
    
    struct Entry { uint32_t start, size; };
    std::map<std::string, Entry> entries;
    
    for(uint32_t sec : dirChain) {
        if ((sec + 1) * sectorSize + sectorSize > data.size()) break;
        const uint8_t* p = data.data() + (sec + 1) * sectorSize;
        for(uint32_t i=0; i<sectorSize/128; ++i) {
            const uint8_t* dir = p + i*128;
            uint16_t nameLen = getU16(dir + 0x40);
            if (nameLen == 0) continue;
            std::string name;
            for(int j=0; j<nameLen-2; j+=2) name += (char)dir[j];
            uint32_t start = getU32(dir + 0x74);
            uint32_t size = getU32(dir + 0x78);
            entries[name] = {start, size};
            if (dir[0x42] == 5) rootStartSector = start;
        }
    }
    
    if (entries.find(streamName) == entries.end()) return std::vector<uint8_t>();
    
    uint32_t miniStreamCutoffSize = getU32(data.data() + 0x38);
    auto rootChain = getChain(rootStartSector);
    auto d = entries[streamName];
    std::vector<uint8_t> out;
    
    if (d.size < miniStreamCutoffSize) {
        auto getMiniChain = [&](uint32_t start) {
            std::vector<uint32_t> chain;
            uint32_t current = start;
            while (current != ENDOFCHAIN && current < minifat.size()) {
                chain.push_back(current);
                current = minifat[current];
            }
            return chain;
        };
        auto chain = getMiniChain(d.start);
        for(uint32_t sec : chain) {
            uint32_t offset = sec * 64;
            if (offset / sectorSize >= rootChain.size()) break;
            uint32_t physSec = rootChain[offset / sectorSize];
            uint32_t physOff = offset % sectorSize;
            if ((physSec + 1) * sectorSize + physOff + 64 > data.size()) break;
            const uint8_t* ptr = data.data() + (physSec + 1) * sectorSize + physOff;
            out.insert(out.end(), ptr, ptr + 64);
        }
    } else {
        auto chain = getChain(d.start);
        for(uint32_t sec : chain) {
            if ((sec + 1) * sectorSize + sectorSize > data.size()) break;
            const uint8_t* ptr = data.data() + (sec + 1) * sectorSize;
            out.insert(out.end(), ptr, ptr + sectorSize);
        }
    }
    
    if (out.size() > d.size) out.resize(d.size);
    return out;
}


#include <mbedtls/ctr_drbg.h>
#include <mbedtls/entropy.h>

std::vector<uint8_t> buildCFB(const std::vector<uint8_t>& info, const std::vector<uint8_t>& pkg) {
    auto putU32 = [](std::vector<uint8_t>& buf, uint32_t offset, uint32_t val) {
        buf[offset] = val & 0xFF; buf[offset+1] = (val >> 8) & 0xFF; buf[offset+2] = (val >> 16) & 0xFF; buf[offset+3] = (val >> 24) & 0xFF;
    };
    auto putU16 = [](std::vector<uint8_t>& buf, uint32_t offset, uint16_t val) {
        buf[offset] = val & 0xFF; buf[offset+1] = (val >> 8) & 0xFF;
    };

    uint32_t pkgSectors = (pkg.size() + 511) / 512;
    uint32_t miniSectors = (info.size() + 63) / 64;
    uint32_t miniStreamSize = miniSectors * 64;
    uint32_t miniStreamSectors = (miniStreamSize + 511) / 512;
    uint32_t miniFatSectors = (miniSectors * 4 + 511) / 512;
    uint32_t dirSectors = 1;
    
    uint32_t dataSectors = dirSectors + miniFatSectors + miniStreamSectors + pkgSectors;
    uint32_t fatSectors = 1;
    while ((dataSectors + fatSectors) * 4 > fatSectors * 512) fatSectors++;
    
    std::vector<uint8_t> cfb(512 + (fatSectors + dataSectors) * 512, 0);
    const uint8_t magic[] = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
    std::memcpy(cfb.data(), magic, 8);
    putU16(cfb, 0x18, 0x003E); putU16(cfb, 0x1A, 0x0003); putU16(cfb, 0x1C, 0xFFFE); putU16(cfb, 0x1E, 0x0009); putU16(cfb, 0x20, 0x0006);
    putU32(cfb, 0x2C, fatSectors);
    
    uint32_t currentSec = 0;
    std::vector<uint32_t> fatLocs;
    for (uint32_t i=0; i<fatSectors; i++) fatLocs.push_back(currentSec++);
    uint32_t miniFatLoc = currentSec; currentSec += miniFatSectors;
    uint32_t dirLoc = currentSec; currentSec += dirSectors;
    uint32_t miniStreamLoc = currentSec; currentSec += miniStreamSectors;
    uint32_t pkgLoc = currentSec; currentSec += pkgSectors;
    
    putU32(cfb, 0x30, dirLoc); putU32(cfb, 0x38, 0x1000); 
    putU32(cfb, 0x3C, miniFatSectors > 0 ? miniFatLoc : 0xFFFFFFFE); putU32(cfb, 0x40, miniFatSectors); putU32(cfb, 0x44, 0xFFFFFFFE);
    
    for (int i=0; i<109; i++) putU32(cfb, 0x4C + i*4, i < fatSectors ? fatLocs[i] : 0xFFFFFFFF);
    

    auto writeChain = [&](uint32_t startSec, uint32_t numSecs) {
        for (uint32_t i=0; i<numSecs; i++) {
            uint32_t sec = startSec + i;
            putU32(cfb, 512 + fatLocs[sec / 128] * 512 + (sec % 128) * 4, i == numSecs - 1 ? 0xFFFFFFFE : sec + 1);
        }
    };
    
    // Initialize FAT with FREESECT (0xFFFFFFFF)
    for (uint32_t sec : fatLocs) {
        for (int i=0; i<128; i++) putU32(cfb, 512 + sec*512 + i*4, 0xFFFFFFFF);
    }
    
    for (uint32_t sec : fatLocs) putU32(cfb, 512 + fatLocs[sec/128]*512 + (sec%128)*4, 0xFFFFFFFD);
    writeChain(miniFatLoc, miniFatSectors); writeChain(dirLoc, dirSectors); writeChain(miniStreamLoc, miniStreamSectors); writeChain(pkgLoc, pkgSectors);
    
    // Initialize MiniFAT with FREESECT (0xFFFFFFFF)
    for (uint32_t i=0; i<miniFatSectors*128; i++) {
        putU32(cfb, 512 + miniFatLoc*512 + i*4, 0xFFFFFFFF);
    }
    for (uint32_t i=0; i<miniSectors; i++) putU32(cfb, 512 + miniFatLoc*512 + i*4, i == miniSectors - 1 ? 0xFFFFFFFE : i + 1);

    
    std::memcpy(cfb.data() + 512 + miniStreamLoc*512, info.data(), info.size());
    std::memcpy(cfb.data() + 512 + pkgLoc*512, pkg.data(), pkg.size());
    
    auto writeDir = [&](int idx, const char* name, int type, int left, int right, int child, uint32_t start, uint32_t size) {
        uint32_t offset = 512 + dirLoc*512 + idx*128;
        int nameLen = std::strlen(name);
        for (int i=0; i<nameLen; i++) cfb[offset + i*2] = name[i];
        putU16(cfb, offset + 64, (nameLen+1)*2);
        cfb[offset + 66] = type; cfb[offset + 67] = 1; // Always Black (1)
        putU32(cfb, offset + 68, left); putU32(cfb, offset + 72, right); putU32(cfb, offset + 76, child);
        putU32(cfb, offset + 116, start); putU32(cfb, offset + 120, size);
    };
    
    // CFB Directory RB-Tree Rules: Compare by Name Length (including null byte), then String.
    // EncryptionInfo (15) < EncryptedPackage (17).
    // Therefore, EncryptionInfo MUST be the left child of EncryptedPackage.
    writeDir(0, "Root Entry", 5, 0xFFFFFFFF, 0xFFFFFFFF, 1, miniStreamLoc, miniStreamSize);
    writeDir(1, "EncryptedPackage", 2, 2, 0xFFFFFFFF, 0xFFFFFFFF, pkgLoc, pkg.size());
    writeDir(2, "EncryptionInfo", 2, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF, 0, info.size());
    
    // Unused directory entries should have pointers = 0xFFFFFFFF
    for (uint32_t i=3; i<dirSectors*4; i++) {
        uint32_t offset = 512 + dirLoc*512 + i*128;
        putU32(cfb, offset + 68, 0xFFFFFFFF); // left
        putU32(cfb, offset + 72, 0xFFFFFFFF); // right
        putU32(cfb, offset + 76, 0xFFFFFFFF); // child
    }
    return cfb;
}

std::vector<uint8_t> encryptStandardPackage(gsl::span<const uint8_t> zipData, const std::string& password) {
    auto putU32 = [](std::vector<uint8_t>& buf, uint32_t offset, uint32_t val) {
        buf[offset] = val & 0xFF; buf[offset+1] = (val >> 8) & 0xFF; buf[offset+2] = (val >> 16) & 0xFF; buf[offset+3] = (val >> 24) & 0xFF;
    };
    
    std::vector<uint8_t> salt(16);
    std::vector<uint8_t> verifier(16);
    mbedtls_entropy_context entropy;
    mbedtls_ctr_drbg_context ctr_drbg;
    mbedtls_entropy_init(&entropy);
    mbedtls_ctr_drbg_init(&ctr_drbg);
    const char* pers = "openxlsx_rand";
    mbedtls_ctr_drbg_seed(&ctr_drbg, mbedtls_entropy_func, &entropy, (const unsigned char*)pers, std::strlen(pers));
    mbedtls_ctr_drbg_random(&ctr_drbg, salt.data(), 16);
    mbedtls_ctr_drbg_random(&ctr_drbg, verifier.data(), 16);
    mbedtls_ctr_drbg_free(&ctr_drbg);
    mbedtls_entropy_free(&entropy);
    
    std::vector<uint8_t> utf16pw;
    for(char c : password) { utf16pw.push_back(static_cast<uint8_t>(c)); utf16pw.push_back(0); }
    
    std::vector<uint8_t> initialData = salt;
    initialData.insert(initialData.end(), utf16pw.begin(), utf16pw.end());
    
    auto hashSHA1 = [](const std::vector<uint8_t>& d) {
        std::vector<uint8_t> h(20); // SHA-1 size
        mbedtls_sha1(d.data(), d.size(), h.data());
        return h;
    };
    
    auto H = hashSHA1(initialData);
    for(int i = 0; i < 50000; ++i) {
        std::vector<uint8_t> iterData = {static_cast<uint8_t>(i&0xFF), static_cast<uint8_t>((i>>8)&0xFF), static_cast<uint8_t>((i>>16)&0xFF), static_cast<uint8_t>((i>>24)&0xFF)};
        iterData.insert(iterData.end(), H.begin(), H.end());
        H = hashSHA1(iterData);
    }
    
    std::vector<uint8_t> finalData = H;
    finalData.insert(finalData.end(), {0, 0, 0, 0});
    auto H_final = hashSHA1(finalData);
    
    std::vector<uint8_t> buf1(64, 0x36), buf2(64, 0x5C);
    for(int i=0; i<20; ++i) { buf1[i] ^= H_final[i]; buf2[i] ^= H_final[i]; }
    
    auto x1 = hashSHA1(buf1);
    auto x2 = hashSHA1(buf2);
    
    std::vector<uint8_t> keyDerived = x1;
    keyDerived.insert(keyDerived.end(), x2.begin(), x2.end());
    keyDerived.resize(16); // 128-bit AES
    
    mbedtls_platform_zeroize(utf16pw.data(), utf16pw.size());
    mbedtls_platform_zeroize(initialData.data(), initialData.size());

    // Encrypt verifier and its hash
    auto aesEcbEncrypt = [&](const std::vector<uint8_t>& data) {
        std::vector<uint8_t> out(data.size());
        mbedtls_aes_context ctx;
        mbedtls_aes_init(&ctx);
        mbedtls_aes_setkey_enc(&ctx, keyDerived.data(), 128);
        for(size_t i=0; i<data.size(); i+=16) {
            mbedtls_aes_crypt_ecb(&ctx, MBEDTLS_AES_ENCRYPT, data.data()+i, out.data()+i);
        }
        mbedtls_aes_free(&ctx);
        return out;
    };
    
    auto encVerifier = aesEcbEncrypt(verifier);
    auto verifierHash = hashSHA1(verifier);
    verifierHash.resize(32, 0);
    auto encVerifierHash = aesEcbEncrypt(verifierHash);
    
    // Construct EncryptionInfo
    std::vector<uint8_t> info(248, 0);
    putU32(info, 0, 0x00020003); // Version
    putU32(info, 4, 0x00000024); // Flags
    putU32(info, 8, 164); // HeaderSize
    putU32(info, 12, 0x00000024); // Flags
    putU32(info, 16, 0x00000000); // SizeExtra
    putU32(info, 20, 0x0000660E); // AlgID (AES-128)
    putU32(info, 24, 0x00008004); // AlgHash (SHA1)
    putU32(info, 28, 128); // KeySize
    putU32(info, 32, 0x00000018); // ProviderType
    
    std::string csp = "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)";
    for(size_t i=0; i<csp.size(); i++) { info[44 + i*2] = csp[i]; }
    
    uint32_t headerEnd = 12 + 164;
    putU32(info, headerEnd, 16); // SaltSize
    std::memcpy(info.data() + headerEnd + 4, salt.data(), 16);
    std::memcpy(info.data() + headerEnd + 20, encVerifier.data(), 16);
    putU32(info, headerEnd + 36, 20); // VerifierHashSize
    std::memcpy(info.data() + headerEnd + 40, encVerifierHash.data(), 32);
    
    // Construct EncryptedPackage
    std::vector<uint8_t> payload(zipData.begin(), zipData.end());
    uint32_t rem = payload.size() % 16;
    if (rem != 0) payload.insert(payload.end(), 16 - rem, 0); // Pad to 16
    
    std::vector<uint8_t> encData(payload.size());
    mbedtls_aes_context ctx;
    mbedtls_aes_init(&ctx);
    mbedtls_aes_setkey_enc(&ctx, keyDerived.data(), 128);
    for(size_t i=0; i<payload.size(); i+=16) {
        mbedtls_aes_crypt_ecb(&ctx, MBEDTLS_AES_ENCRYPT, payload.data()+i, encData.data()+i);
    }
    mbedtls_aes_free(&ctx);
    
    std::vector<uint8_t> encPackage(8 + encData.size());
    uint64_t originalSize = zipData.size();
    encPackage[0] = originalSize & 0xFF; encPackage[1] = (originalSize >> 8) & 0xFF;
    encPackage[2] = (originalSize >> 16) & 0xFF; encPackage[3] = (originalSize >> 24) & 0xFF;
    encPackage[4] = (originalSize >> 32) & 0xFF; encPackage[5] = (originalSize >> 40) & 0xFF;
    encPackage[6] = (originalSize >> 48) & 0xFF; encPackage[7] = (originalSize >> 56) & 0xFF;
    std::memcpy(encPackage.data() + 8, encData.data(), encData.size());

    return buildCFB(info, encPackage);
}

std::vector<uint8_t> aes256CbcDecrypt(gsl::span<const uint8_t> data, gsl::span<const uint8_t> key, gsl::span<const uint8_t> iv) { return {}; }
std::vector<uint8_t> sha512Hash(gsl::span<const uint8_t> data) { return {}; }

std::vector<uint8_t> generateAgileHash(const std::string& password, gsl::span<const uint8_t> salt, int spinCount) {
    std::vector<uint8_t> utf16pw;
    for (char c : password) {
        utf16pw.push_back(static_cast<uint8_t>(c));
        utf16pw.push_back(0x00);
    }
    
    std::vector<uint8_t> initialData(salt.begin(), salt.end());
    initialData.insert(initialData.end(), utf16pw.begin(), utf16pw.end());
    
    auto H = sha512Hash(initialData);

    for (uint32_t i = 0; i < static_cast<uint32_t>(spinCount); ++i) {
        std::vector<uint8_t> loopData(4);
        loopData[0] = (i & 0xFF);
        loopData[1] = ((i >> 8) & 0xFF);
        loopData[2] = ((i >> 16) & 0xFF);
        loopData[3] = ((i >> 24) & 0xFF);
        loopData.insert(loopData.end(), H.begin(), H.end());
        H = sha512Hash(loopData);
    }
    
    mbedtls_platform_zeroize(utf16pw.data(), utf16pw.size());
    mbedtls_platform_zeroize(initialData.data(), initialData.size());
    
    return H;
}

std::vector<uint8_t> decryptAgilePackage(gsl::span<const uint8_t> encryptionInfo,
                                         gsl::span<const uint8_t> encryptedPackage,
                                         const std::string& password) {
    if (encryptionInfo.empty() || encryptedPackage.empty()) return std::vector<uint8_t>();
    
    pugi::xml_document doc;
    pugi::xml_parse_result result = doc.load_buffer(encryptionInfo.data(), encryptionInfo.size());
    if (!result) throw XLInternalError("Failed to parse EncryptionInfo XML");

    // Placeholder for actual agile logic:
    return std::vector<uint8_t>();
}

std::vector<uint8_t> decryptStandardPackage(gsl::span<const uint8_t> encryptionInfo,
                                            gsl::span<const uint8_t> encryptedPackage,
                                            const std::string& password) {
    if (encryptionInfo.size() < 24 || encryptedPackage.size() < 8) return std::vector<uint8_t>();

    auto getU32 = [](const uint8_t* p) { return p[0] | (p[1]<<8) | (p[2]<<16) | (p[3]<<24); };

    uint32_t headerSize = getU32(encryptionInfo.data() + 8);
    uint32_t keySize = getU32(encryptionInfo.data() + 28);
    uint32_t saltSize = getU32(encryptionInfo.data() + 12 + headerSize);
    
    if (encryptionInfo.size() < 12 + headerSize + 4 + saltSize) throw XLInternalError("Invalid Standard Encryption Header");
    std::vector<uint8_t> salt(encryptionInfo.data() + 12 + headerSize + 4, encryptionInfo.data() + 12 + headerSize + 4 + saltSize);
    
    std::vector<uint8_t> utf16pw;
    for(char c : password) { utf16pw.push_back(static_cast<uint8_t>(c)); utf16pw.push_back(0); }
    
    std::vector<uint8_t> initialData = salt;
    initialData.insert(initialData.end(), utf16pw.begin(), utf16pw.end());
    
    auto hashSHA1 = [](const std::vector<uint8_t>& d) {
        std::vector<uint8_t> h(20); // SHA-1 size
        mbedtls_sha1(d.data(), d.size(), h.data());
        return h;
    };
    
    auto H = hashSHA1(initialData);
    
    for(int i = 0; i < 50000; ++i) {
        std::vector<uint8_t> iterData = {static_cast<uint8_t>(i&0xFF), static_cast<uint8_t>((i>>8)&0xFF), static_cast<uint8_t>((i>>16)&0xFF), static_cast<uint8_t>((i>>24)&0xFF)};
        iterData.insert(iterData.end(), H.begin(), H.end());
        H = hashSHA1(iterData);
    }
    
    std::vector<uint8_t> finalData = H;
    finalData.insert(finalData.end(), {0, 0, 0, 0});
    auto H_final = hashSHA1(finalData);
    
    std::vector<uint8_t> buf1(64, 0x36), buf2(64, 0x5C);
    for(int i=0; i<20; ++i) { buf1[i] ^= H_final[i]; buf2[i] ^= H_final[i]; }
    
    auto x1 = hashSHA1(buf1);
    auto x2 = hashSHA1(buf2);
    
    std::vector<uint8_t> keyDerived = x1;
    keyDerived.insert(keyDerived.end(), x2.begin(), x2.end());
    keyDerived.resize(keySize / 8);
    
    mbedtls_platform_zeroize(utf16pw.data(), utf16pw.size());
    mbedtls_platform_zeroize(initialData.data(), initialData.size());

    // ECB Decryption
    std::vector<uint8_t> payload(encryptedPackage.begin() + 8, encryptedPackage.end());
    std::vector<uint8_t> decrypted(payload.size());
    
    mbedtls_aes_context ctx;
    mbedtls_aes_init(&ctx);
    auto cleanup = gsl::finally([&] { mbedtls_aes_free(&ctx); });

    if (mbedtls_aes_setkey_dec(&ctx, keyDerived.data(), keySize) != 0) {
        throw XLInternalError("AES key setup failed");
    }
    
    for(size_t i=0; i<payload.size(); i+=16) {
        mbedtls_aes_crypt_ecb(&ctx, MBEDTLS_AES_DECRYPT, payload.data()+i, decrypted.data()+i);
    }
    int len1 = payload.size(), len2 = 0;
    
    uint64_t originalSize = getU32(encryptedPackage.data()) | ((uint64_t)getU32(encryptedPackage.data() + 4) << 32);
    if (originalSize > 0 && originalSize <= static_cast<uint64_t>(len1 + len2)) {
        decrypted.resize(originalSize);
    } else {
        decrypted.resize(len1 + len2);
    }
    return decrypted;
}

} // namespace OpenXLSX::Crypto

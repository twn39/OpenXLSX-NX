/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2026, Curry Tang

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

// ===== External Includes ===== //
#include <zip.h>
#include <cstring>
#include <stdexcept>
#include <vector>
#include <filesystem>

// ===== OpenXLSX Includes ===== //
#include "XLZipArchive.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

struct XLZipArchive::LibZipApp {
    zip_t* archive = nullptr;
    std::string currentPath;

    LibZipApp() = default;
    ~LibZipApp() {
        if (archive) {
            zip_discard(archive);
            archive = nullptr;
        }
    }

    LibZipApp(const LibZipApp&) = delete;
    LibZipApp& operator=(const LibZipApp&) = delete;
};

XLZipArchive::XLZipArchive() : m_archive(nullptr) {}

XLZipArchive::~XLZipArchive() = default;

XLZipArchive::operator bool() const { return isValid() && isOpen(); }

bool XLZipArchive::isValid() const { 
    return m_archive != nullptr; 
}

bool XLZipArchive::isOpen() const { 
    return m_archive && m_archive->archive != nullptr; 
}

void XLZipArchive::open(const std::string& fileName) {
    if (isOpen()) close();

    if (!m_archive) m_archive = std::make_shared<LibZipApp>();

    int err = 0;
#if defined(_WIN32)
    zip_error_t zerr;
    zip_error_init(&zerr);
    // Use wide string for Windows to support UTF-8 paths
    std::wstring wPath  = std::filesystem::u8path(fileName).wstring();
    zip_source_t* src = zip_source_win32w_create(wPath.c_str(), 0, -1, &zerr);
    if (!src) {
        std::string msg = zip_error_strerror(&zerr);
        zip_error_fini(&zerr);
        throw XLInternalError("Failed to create zip source: " + msg);
    }
    m_archive->archive = zip_open_from_source(src, ZIP_CREATE, &zerr);
    if (!m_archive->archive) {
        std::string msg = zip_error_strerror(&zerr);
        zip_source_free(src);
        zip_error_fini(&zerr);
        throw XLInternalError("Failed to open zip archive from source: " + msg);
    }
    zip_error_fini(&zerr);
#else
    m_archive->archive = zip_open(fileName.c_str(), ZIP_CREATE, &err);
    if (!m_archive->archive) {
        zip_error_t zerr;
        zip_error_init_with_code(&zerr, err);
        std::string msg = zip_error_strerror(&zerr);
        zip_error_fini(&zerr);
        throw XLInternalError("Failed to open zip archive: " + msg);
    }
#endif
    m_archive->currentPath = fileName;
}

void XLZipArchive::close() {
    if (isOpen()) {
        zip_t* ptr = m_archive->archive;
        m_archive->archive = nullptr;
        if (zip_close(ptr) < 0) {
            std::string msg = zip_strerror(ptr);
            zip_discard(ptr);
            throw XLInternalError("Failed to close zip archive: " + msg);
        }
    }
}

void XLZipArchive::save(const std::string& path) {
    if (!isOpen()) return;

    if (path.empty() || path == m_archive->currentPath) {
        std::string current = m_archive->currentPath;
        close();
        open(current);
    } else {
        std::string current = m_archive->currentPath;
        close();
        if (std::filesystem::exists(std::filesystem::u8path(current))) {
            std::filesystem::copy_file(std::filesystem::u8path(current), std::filesystem::u8path(path), std::filesystem::copy_options::overwrite_existing);
        }
        open(current);
    }
}

void XLZipArchive::addEntry(const std::string& name, const std::string& data) {
    if (!isOpen()) throw XLInternalError("Archive not open");

    void* rawBuffer = malloc(data.size());
    if (!rawBuffer) throw std::bad_alloc();
    
    // Use RAII to guarantee safety in case of exceptions before zip_source_buffer
    std::unique_ptr<void, decltype(&std::free)> safeBuffer(rawBuffer, std::free);
    memcpy(safeBuffer.get(), data.c_str(), data.size());

    zip_source_t* s = zip_source_buffer(m_archive->archive, safeBuffer.get(), static_cast<zip_uint64_t>(data.size()), 1);
    if (!s) {
        throw XLInternalError("Failed to create zip source");
    }
    // Release ownership to libzip since zip_source_buffer(..., 1) will free it
    safeBuffer.release();

    zip_int64_t idx = zip_name_locate(m_archive->archive, name.c_str(), 0);
    if (idx >= 0) {
        if (zip_file_replace(m_archive->archive, idx, s, ZIP_FL_ENC_UTF_8) < 0) {
            zip_source_free(s);
            throw XLInternalError(std::string("Failed to replace entry: ") + zip_strerror(m_archive->archive));
        }
    } else {
        if (zip_file_add(m_archive->archive, name.c_str(), s, ZIP_FL_ENC_UTF_8) < 0) {
            zip_source_free(s);
            throw XLInternalError(std::string("Failed to add entry: ") + zip_strerror(m_archive->archive));
        }
    }
}

void XLZipArchive::deleteEntry(const std::string& entryName) {
    if (!isOpen()) return;

    zip_int64_t idx = zip_name_locate(m_archive->archive, entryName.c_str(), 0);
    if (idx >= 0) {
        if (zip_delete(m_archive->archive, idx) < 0) {
            throw XLInternalError(std::string("Failed to delete entry: ") + zip_strerror(m_archive->archive));
        }
    }
}

std::string XLZipArchive::getEntry(const std::string& name) const {
    if (!isOpen()) throw XLInternalError("Archive not open");

    zip_stat_t st;
    zip_stat_init(&st);
    if (zip_stat(m_archive->archive, name.c_str(), 0, &st) < 0) {
        throw XLInternalError("Entry not found: " + name);
    }

    zip_file_t* f = zip_fopen(m_archive->archive, name.c_str(), 0);
    if (!f) throw XLInternalError("Failed to open entry: " + name);

    std::string result;
    result.resize(st.size);
    zip_int64_t bytesRead = zip_fread(f, result.data(), st.size);
    if (bytesRead < 0) {
        zip_fclose(f);
        throw XLInternalError("Failed to read entry: " + name);
    }
    result.resize(bytesRead);
    zip_fclose(f);

    return result;
}

bool XLZipArchive::hasEntry(const std::string& entryName) const {
    if (!isOpen()) return false;
    return zip_name_locate(m_archive->archive, entryName.c_str(), 0) >= 0;
}

// ===== External Includes ===== //
#include <cstring>
#include <filesystem>
#include <stdexcept>
#include <vector>
#include <zip.h>

// ===== OpenXLSX Includes ===== //
#include "XLException.hpp"
#include "XLZipArchive.hpp"

using namespace OpenXLSX;

struct XLZipArchive::LibZipApp
{
    zip_t*      archive = nullptr;
    std::string currentPath;
    bool        isModified = false;    // Track if we explicitly want to save

    LibZipApp() = default;
    ~LibZipApp()
    {
        if (archive) {
            zip_discard(archive);    // Discard any unsaved changes safely
            archive = nullptr;
        }
    }

    LibZipApp(const LibZipApp&)            = delete;
    LibZipApp& operator=(const LibZipApp&) = delete;
};

XLZipArchive::XLZipArchive() : m_archive(nullptr) {}

XLZipArchive::~XLZipArchive() = default;

XLZipArchive::operator bool() const { return isValid() && isOpen(); }

bool XLZipArchive::isValid() const { return m_archive != nullptr; }

bool XLZipArchive::isOpen() const { return m_archive && m_archive->archive != nullptr; }

void XLZipArchive::open(std::string_view fileName)
{
    if (isOpen()) close();

    if (!m_archive) m_archive = std::make_shared<LibZipApp>();
    m_archive->isModified = false;    // Reset modification flag on open

    int err = 0;
#if defined(_WIN32)
    zip_error_t zerr;
    zip_error_init(&zerr);
    // Use wide string for Windows to support UTF-8 paths
    std::wstring  wPath = std::filesystem::u8path(fileName).wstring();
    zip_source_t* src   = zip_source_win32w_create(wPath.c_str(), 0, -1, &zerr);
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
    m_archive->archive = zip_open(std::string(fileName).c_str(), ZIP_CREATE, &err);
    if (!m_archive->archive) {
        zip_error_t zerr;
        zip_error_init_with_code(&zerr, err);
        std::string msg = zip_error_strerror(&zerr);
        zip_error_fini(&zerr);
        throw XLInternalError("Failed to open zip archive: " + msg);
    }
#endif
    m_archive->currentPath = std::string(fileName);
}

void XLZipArchive::close()
{
    if (isOpen()) {
        zip_t* ptr         = m_archive->archive;
        m_archive->archive = nullptr;
        // If save() was called, we should commit changes.
        // Otherwise, we discard to prevent silent corruption on read-only operations.
        if (m_archive->isModified) {
            if (zip_close(ptr) < 0) {
                std::string msg = zip_strerror(ptr);
                zip_discard(ptr);
                throw XLInternalError("Failed to close zip archive: " + msg);
            }
        }
        else {
            zip_discard(ptr);
        }
    }
}

void XLZipArchive::save(std::string_view path)
{
    if (!isOpen()) return;

    if (path.empty() || path == m_archive->currentPath) {
        std::string current   = m_archive->currentPath;
        m_archive->isModified = true;    // Mark as modified before closing to trigger commit
        close();
        open(current);
    }
    else {
        std::string current   = m_archive->currentPath;
        m_archive->isModified = true;    // Mark as modified
        close();                         // Commits changes to the original file
        if (std::filesystem::exists(std::filesystem::u8path(current))) {
            std::filesystem::copy_file(std::filesystem::u8path(current),
                                       std::filesystem::u8path(path),
                                       std::filesystem::copy_options::overwrite_existing);
        }
        open(current);
    }
}

void XLZipArchive::addEntry(std::string_view name, std::string_view data)
{
    if (!isOpen()) throw XLInternalError("Archive not open");

    m_archive->isModified = true;    // Mark as modified

    // Create a robust copy using string and release it to zip_source_buffer
    // We use new char[] instead of malloc to follow C++ guidelines slightly better,
    // though malloc is also fine here since libzip's zip_source_buffer(..., 1) will internally call free().
    // Actually, libzip *requires* the buffer to be allocated with malloc() when using free_data=1
    void* rawBuffer = std::malloc(data.size());
    if (!rawBuffer) throw std::bad_alloc();

    std::memcpy(rawBuffer, data.data(), data.size());

    zip_source_t* s = zip_source_buffer(m_archive->archive, rawBuffer, static_cast<zip_uint64_t>(data.size()), 1);
    if (!s) { 
        std::free(rawBuffer); // Free immediately if source creation fails
        throw XLInternalError("Failed to create zip source"); 
    }

    zip_int64_t idx = zip_name_locate(m_archive->archive, std::string(name).c_str(), 0);
    if (idx >= 0) {
        if (zip_file_replace(m_archive->archive, idx, s, ZIP_FL_ENC_UTF_8) < 0) {
            zip_source_free(s);
            throw XLInternalError(std::string("Failed to replace entry: ") + zip_strerror(m_archive->archive));
        }
    }
    else {
        if (zip_file_add(m_archive->archive, std::string(name).c_str(), s, ZIP_FL_ENC_UTF_8) < 0) {
            zip_source_free(s);
            throw XLInternalError(std::string("Failed to add entry: ") + zip_strerror(m_archive->archive));
        }
    }
}

void XLZipArchive::deleteEntry(std::string_view entryName)
{
    if (!isOpen()) return;

    m_archive->isModified = true;    // Mark as modified
    zip_int64_t idx       = zip_name_locate(m_archive->archive, std::string(entryName).c_str(), 0);
    if (idx >= 0) {
        if (zip_delete(m_archive->archive, idx) < 0) {
            throw XLInternalError(std::string("Failed to delete entry: ") + zip_strerror(m_archive->archive));
        }
    }
}

std::string XLZipArchive::getEntry(std::string_view name) const
{
    if (!isOpen()) throw XLInternalError("Archive not open");

    zip_stat_t st;
    zip_stat_init(&st);
    if (zip_stat(m_archive->archive, std::string(name).c_str(), 0, &st) < 0) { throw XLInternalError("Entry not found: " + std::string(name)); }

    zip_file_t* f = zip_fopen(m_archive->archive, std::string(name).c_str(), 0);
    if (!f) throw XLInternalError("Failed to open entry: " + std::string(name));

    std::string result;
    result.resize(st.size);
    zip_int64_t bytesRead = zip_fread(f, result.data(), st.size);
    if (bytesRead < 0) {
        zip_fclose(f);
        throw XLInternalError("Failed to read entry: " + std::string(name));
    }
    result.resize(bytesRead);
    zip_fclose(f);

    return result;
}

bool XLZipArchive::hasEntry(std::string_view entryName) const
{
    if (!isOpen()) return false;
    return zip_name_locate(m_archive->archive, std::string(entryName).c_str(), 0) >= 0;
}

std::vector<std::string> XLZipArchive::entryNames() const
{
    if (!isOpen()) return {};

    zip_int64_t numEntries = zip_get_num_entries(m_archive->archive, 0);
    if (numEntries <= 0) return {};

    std::vector<std::string> result;
    result.reserve(static_cast<size_t>(numEntries));
    for (zip_int64_t i = 0; i < numEntries; ++i) { 
        result.emplace_back(zip_get_name(m_archive->archive, i, 0)); 
    }

    return result;
}

void XLZipArchive::addEntryFromFile(std::string_view name, std::string_view filePath)
{
    if (!isOpen()) throw XLInternalError("Archive not open");

    m_archive->isModified = true;

    zip_source_t* s = zip_source_file(m_archive->archive, std::string(filePath).c_str(), 0, 0);
    if (!s) { 
        throw XLInternalError("Failed to create zip source from file: " + std::string(filePath)); 
    }

    zip_int64_t index = zip_file_add(m_archive->archive, std::string(name).c_str(), s, ZIP_FL_OVERWRITE);
    if (index < 0) {
        zip_source_free(s);
        throw XLInternalError("Failed to add file entry to archive");
    }
}

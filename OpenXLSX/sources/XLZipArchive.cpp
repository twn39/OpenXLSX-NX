// ===== External Includes ===== //
#include <cstring>
#include <deque>
#include <filesystem>
#include <stdexcept>
#include <vector>
#include <zip.h>

// ===== OpenXLSX Includes ===== //
#include "XLException.hpp"
#include "XLZipArchive.hpp"
#include <memory>

using namespace OpenXLSX;

struct ZipArchiveDeleter {
    void operator()(zip_t* ptr) const {
        if (ptr) zip_discard(ptr);
    }
};
using ZipArchivePtr = std::unique_ptr<zip_t, ZipArchiveDeleter>;

struct ZipSourceDeleter {
    void operator()(zip_source_t* src) const {
        if (src) zip_source_free(src);
    }
};
using ZipSourcePtr = std::unique_ptr<zip_source_t, ZipSourceDeleter>;

struct XLZipArchive::LibZipApp
{
    ZipArchivePtr archive{nullptr};
    std::string   currentPath;
    bool          isModified = false;    // Track if we explicitly want to save
    
    // Stable deque string cache for Zero-Copy zip_source_buffer_create
    std::deque<std::string> stringCache; 

    LibZipApp() = default;
    LibZipApp(const LibZipApp&)            = delete;
    LibZipApp& operator=(const LibZipApp&) = delete;
};

XLZipArchive::XLZipArchive() : m_archive(nullptr) {}

XLZipArchive::~XLZipArchive() = default;

XLZipArchive::operator bool() const { return isValid() && isOpen(); }

bool XLZipArchive::isValid() const { return m_archive != nullptr; }

bool XLZipArchive::isOpen() const { return m_archive && m_archive->archive != nullptr; }

void XLZipArchive::setCompressionLevel(int level)
{
    m_compressionLevel = level;
}

int XLZipArchive::compressionLevel() const
{
    return m_compressionLevel;
}

void XLZipArchive::open(std::string_view fileName)
{
    if (isOpen()) close();

    if (!m_archive) m_archive = std::make_shared<LibZipApp>();
    m_archive->isModified = false;    // Reset modification flag on open

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
    m_archive->archive.reset(zip_open_from_source(src, ZIP_CREATE, &zerr));
    if (!m_archive->archive.get()) {
        std::string msg = zip_error_strerror(&zerr);
        zip_source_free(src);
        zip_error_fini(&zerr);
        throw XLInternalError("Failed to open zip archive from source: " + msg);
    }
    zip_error_fini(&zerr);
    #else
    int err = 0;
    m_archive->archive.reset(zip_open(std::string(fileName).c_str(), ZIP_CREATE, &err));
    if (!m_archive->archive.get()) {
        zip_error_t error;
        zip_error_init_with_code(&error, err);
        std::string msg = zip_error_strerror(&error);
        zip_error_fini(&error);
        throw XLInternalError("Failed to open zip archive: " + msg);
    }
    #endif
    m_archive->currentPath = std::string(fileName);
}

void XLZipArchive::close()
{
    if (isOpen()) {
        ZipArchivePtr ptr = std::move(m_archive->archive);

        // If save() was called, we should commit changes.
        // Otherwise, we discard to prevent silent corruption on read-only operations.
        if (m_archive->isModified) {
            zip_int64_t numEntries = zip_get_num_entries(ptr.get(), 0);
            for (zip_int64_t i = 0; i < numEntries; ++i) {
                if (zip_get_name(ptr.get(), i, 0) != nullptr) {
                    zip_int32_t method = (m_compressionLevel == 0) ? ZIP_CM_STORE : ZIP_CM_DEFLATE;
                    zip_set_file_compression(ptr.get(), i, method, m_compressionLevel);
                }
            }

            if (zip_close(ptr.get()) < 0) {
                throw XLInternalError("Failed to close zip archive: " + std::string(zip_strerror(ptr.get())));
            }
            // libzip has taken ownership and freed the pointer on successful close
            ptr.release();
        }
        // If m_archive->isModified is false, ptr leaves scope and ZipArchiveDeleter
        // safely calls zip_discard.
        
        // Clear string cache AFTER libzip finishes reading pointers
        m_archive->stringCache.clear();
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
            try {
                std::filesystem::copy_file(std::filesystem::u8path(current),
                                           std::filesystem::u8path(path),
                                           std::filesystem::copy_options::overwrite_existing);
            } catch (const std::filesystem::filesystem_error& e) {
                throw XLInternalError(std::string("Failed to save to new path: ") + e.what());
            }
        }
        open(current);
    }
}

void XLZipArchive::addEntry(std::string_view name, std::string data)
{
    if (!isOpen()) throw XLInternalError("Archive not open");

    m_archive->isModified = true;    // Mark as modified

    // 1. Move temporary string to stable deque arena
    m_archive->stringCache.push_back(std::move(data));
    const std::string& cachedStr = m_archive->stringCache.back();

    // 2. Use zip_source_buffer_create for 0-copy. Libzip reads this during zip_close().
    zip_error_t zerr;
    zip_error_init(&zerr);
    ZipSourcePtr s(zip_source_buffer_create(cachedStr.data() ? cachedStr.data() : "",
                                            static_cast<zip_uint64_t>(cachedStr.size()),
                                            0, &zerr));
    if (!s) {
        throw XLInternalError(std::string("Failed to create zip source: ") + zip_error_strerror(&zerr));
    }

    zip_int64_t idx = zip_name_locate(m_archive->archive.get(), std::string(name).c_str(), 0);
    if (idx >= 0) {
        if (zip_file_replace(m_archive->archive.get(), idx, s.get(), ZIP_FL_ENC_UTF_8) < 0) {
            throw XLInternalError(std::string("Failed to replace entry: ") + zip_strerror(m_archive->archive.get()));
        }
    }
    else {
        if (zip_file_add(m_archive->archive.get(), std::string(name).c_str(), s.get(), ZIP_FL_ENC_UTF_8) < 0) {
            throw XLInternalError(std::string("Failed to add entry: ") + zip_strerror(m_archive->archive.get()));
        }
    }
    // API successfully took ownership
    s.release();
}

void XLZipArchive::addEntryAllocated(std::string_view name, void* data, size_t size)
{
    if (!isOpen()) {
        if (data) std::free(data); // ensure cleanup
        throw XLInternalError("Archive not open");
    }

    m_archive->isModified = true;    // Mark as modified

    // Create a robust copy using string and release it to zip_source_buffer
    // We use new char[] instead of malloc to follow C++ guidelines slightly better,
    // though malloc is also fine here since libzip's zip_source_buffer(..., 1) will internally call free().
    // Actually, libzip *requires* the buffer to be allocated with malloc() when using free_data=1
    ZipSourcePtr s(zip_source_buffer(m_archive->archive.get(), data ? data : "", static_cast<zip_uint64_t>(size), data ? 1 : 0));
    if (!s) {
        if (data) std::free(data);    // Free immediately if source creation fails
        throw XLInternalError("Failed to create zip source");
    }

    zip_int64_t idx = zip_name_locate(m_archive->archive.get(), std::string(name).c_str(), 0);
    if (idx >= 0) {
        if (zip_file_replace(m_archive->archive.get(), idx, s.get(), ZIP_FL_ENC_UTF_8) < 0) {
            throw XLInternalError(std::string("Failed to replace entry: ") + zip_strerror(m_archive->archive.get()));
        }
    }
    else {
        if (zip_file_add(m_archive->archive.get(), std::string(name).c_str(), s.get(), ZIP_FL_ENC_UTF_8) < 0) {
            throw XLInternalError(std::string("Failed to add entry: ") + zip_strerror(m_archive->archive.get()));
        }
    }
    // API successfully took ownership
    s.release();
}

void XLZipArchive::deleteEntry(std::string_view entryName)
{
    if (!isOpen()) return;

    m_archive->isModified = true;    // Mark as modified
    zip_int64_t idx       = zip_name_locate(m_archive->archive.get(), std::string(entryName).c_str(), 0);
    if (idx >= 0) {
        if (zip_delete(m_archive->archive.get(), idx) < 0) {
            throw XLInternalError(std::string("Failed to delete entry: ") + zip_strerror(m_archive->archive.get()));
        }
    }
}

std::string XLZipArchive::getEntry(std::string_view name) const
{
    if (!isOpen()) throw XLInternalError("Archive not open");

    zip_stat_t st;
    zip_stat_init(&st);
    if (zip_stat(m_archive->archive.get(), std::string(name).c_str(), 0, &st) < 0) {
        throw XLInternalError("Entry not found: " + std::string(name));
    }

    zip_file_t* f = zip_fopen(m_archive->archive.get(), std::string(name).c_str(), 0);
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

void* XLZipArchive::openEntryStream(std::string_view name) const
{
    if (!isOpen()) throw XLInternalError("Archive not open");
    zip_file_t* f = zip_fopen(m_archive->archive.get(), std::string(name).c_str(), 0);
    if (!f) throw XLInternalError("Failed to open entry stream: " + std::string(name));
    return f;
}

int64_t XLZipArchive::readEntryStream(void* stream, char* buffer, uint64_t size) const
{
    if (!stream) return -1;
    return zip_fread(static_cast<zip_file_t*>(stream), buffer, size);
}

void XLZipArchive::closeEntryStream(void* stream) const
{
    if (stream) zip_fclose(static_cast<zip_file_t*>(stream));
}

bool XLZipArchive::hasEntry(std::string_view entryName) const
{
    if (!isOpen()) return false;
    return zip_name_locate(m_archive->archive.get(), std::string(entryName).c_str(), 0) >= 0;
}

std::vector<std::string> XLZipArchive::entryNames() const
{
    if (!isOpen()) return {};

    zip_int64_t numEntries = zip_get_num_entries(m_archive->archive.get(), 0);
    if (numEntries <= 0) return {};

    std::vector<std::string> result;
    result.reserve(static_cast<size_t>(numEntries));
    for (zip_int64_t i = 0; i < numEntries; ++i) {
        // zip_get_name() returns nullptr for entries marked for deletion (pending delete).
        // Constructing std::string from nullptr is UB and causes SIGSEGV on some platforms.
        const char* name = zip_get_name(m_archive->archive.get(), i, 0);
        if (name != nullptr) { result.emplace_back(name); }
    }

    return result;
}

void XLZipArchive::addEntryFromFile(std::string_view name, std::string_view filePath)
{
    if (!isOpen()) throw XLInternalError("Archive not open");

    m_archive->isModified = true;

    ZipSourcePtr s(zip_source_file(m_archive->archive.get(), std::string(filePath).c_str(), 0, 0));
    if (!s) { throw XLInternalError("Failed to create zip source from file: " + std::string(filePath)); }

    zip_int64_t index = zip_file_add(m_archive->archive.get(), std::string(name).c_str(), s.get(), ZIP_FL_OVERWRITE);
    if (index < 0) {
        throw XLInternalError("Failed to add file entry to archive");
    }
    s.release();
}

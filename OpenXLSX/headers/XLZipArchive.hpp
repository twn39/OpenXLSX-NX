#ifndef OPENXLSX_XLZIPARCHIVE_HPP
#define OPENXLSX_XLZIPARCHIVE_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"

#include <memory>
#include <string>

namespace OpenXLSX
{
    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLZipArchive
    {
    public:
        /**
         * @brief
         */
        XLZipArchive();

        /**
         * @brief
         * @param other
         */
        XLZipArchive(const XLZipArchive& other) = default;

        /**
         * @brief
         * @param other
         */
        XLZipArchive(XLZipArchive&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLZipArchive();

        /**
         * @brief
         * @param other
         * @return
         */
        XLZipArchive& operator=(const XLZipArchive& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLZipArchive& operator=(XLZipArchive&& other) noexcept = default;

        /**
         * @brief
         * @return
         */
        explicit operator bool() const;

        [[nodiscard]] bool isValid() const;

        /**
         * @brief
         * @return
         */
        [[nodiscard]] bool isOpen() const;

        /**
         * @brief
         * @param fileName
         */
        void open(std::string_view fileName);

        /**
         * @brief
         */
        void close();

        /**
         * @brief
         * @param path
         */
        void save(std::string_view path = "");

        /**
         * @brief
         * @param name
         * @param data
         */
        void addEntry(std::string_view name, std::string_view data);

        /**
         * @brief Add a file to the ZIP archive directly from the file system (useful for streaming large files).
         * @param name The name of the entry within the ZIP archive.
         * @param filePath The path to the file on disk.
         */
        void addEntryFromFile(std::string_view name, std::string_view filePath);


        /**
         * @brief
         * @param entryName
         */
        void deleteEntry(std::string_view entryName);

        /**
         * @brief
         * @param name
         * @return
         */
        [[nodiscard]] std::string getEntry(std::string_view name) const;

        /**
         * @brief Opens a stream for reading a specific entry.
         * @param name The name of the entry to open.
         * @return An opaque pointer to the stream.
         */
        void* openEntryStream(std::string_view name) const;

        /**
         * @brief Reads data from an opened stream.
         * @param stream The opaque stream pointer.
         * @param buffer The buffer to read into.
         * @param size The maximum number of bytes to read.
         * @return The number of bytes read, or -1 on error.
         */
        int64_t readEntryStream(void* stream, char* buffer, uint64_t size) const;

        /**
         * @brief Closes an opened stream.
         * @param stream The opaque stream pointer.
         */
        void closeEntryStream(void* stream) const;

        /**
         * @brief
         * @param entryName
         * @return
         */
        [[nodiscard]] bool hasEntry(std::string_view entryName) const;

        /**
         * @brief Get a list of all entry names in the archive.
         * @return A vector of strings containing the names of all entries.
         */
        [[nodiscard]] std::vector<std::string> entryNames() const;

    private:
        struct LibZipApp;
        std::shared_ptr<LibZipApp> m_archive; /**< */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLZIPARCHIVE_HPP

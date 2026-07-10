#ifndef OPENXLSX_TEST_MOCKPACKAGEPORTS_HPP
#define OPENXLSX_TEST_MOCKPACKAGEPORTS_HPP

/**
 * @file MockPackagePorts.hpp
 * @brief In-memory package / archive mocks for unit tests (no real .xlsx required).
 */

#include "IZipArchive.hpp"
#include "XLContentTypes.hpp"
#include "XLException.hpp"
#include "XLPackageServices.hpp"
#include "XLRelationships.hpp"
#include "XLSharedStringTable.hpp"
#include "XLXmlData.hpp"

#include <map>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

namespace OpenXLSX::test
{

    /**
     * @brief Minimal in-memory ZIP backend compatible with IZipArchive type erasure.
     * @details Entry storage is shared via shared_ptr so IZipArchive's value-copy of the
     *          backend still sees the same entries (required by type-erasure clone).
     */
    class InMemoryZipBackend
    {
    public:
        InMemoryZipBackend() : m_entries(std::make_shared<std::map<std::string, std::string>>()) {}

        bool isValid() const { return true; }
        bool isOpen() const { return m_open; }
        void open(const std::string&) { m_open = true; }
        void close() const {}
        void save(const std::string&) const {}

        void addEntry(const std::string& name, std::string data) { (*m_entries)[name] = std::move(data); }

        void addEntryAllocated(std::string_view name, void* data, size_t size)
        {
            (*m_entries)[std::string(name)] = std::string(static_cast<const char*>(data), size);
        }

        void addEntryFromFile(std::string_view, std::string_view) {}

        void deleteEntry(const std::string& entryName) { m_entries->erase(entryName); }

        std::string getEntry(const std::string& name) const
        {
            auto it = m_entries->find(name);
            return it != m_entries->end() ? it->second : std::string{};
        }

        void* openEntryStream(std::string_view) const { return nullptr; }
        int64_t readEntryStream(void*, char*, uint64_t) const { return 0; }
        void closeEntryStream(void*) const {}

        bool hasEntry(const std::string& entryName) const { return m_entries->find(entryName) != m_entries->end(); }

        std::vector<std::string> entryNames() const
        {
            std::vector<std::string> names;
            names.reserve(m_entries->size());
            for (const auto& kv : *m_entries) names.push_back(kv.first);
            return names;
        }

        void setCompressionLevel(int level) { m_level = level; }
        int  compressionLevel() const { return m_level; }

        std::map<std::string, std::string>& entries() { return *m_entries; }

    private:
        bool                                              m_open{true};
        int                                               m_level{1};
        std::shared_ptr<std::map<std::string, std::string>> m_entries;
    };

    /**
     * @brief XLPackageServices mock: real IZipArchive over InMemoryZipBackend; other methods stubbed.
     * @details Sufficient for DrawingItem::imageBinary and similar archive-only consumers.
     */
    class MockPackageServices final : public XLPackageServices
    {
    public:
        MockPackageServices() : m_archive(m_backend) {}

        InMemoryZipBackend&       backend() { return m_backend; }
        const InMemoryZipBackend& backend() const { return m_backend; }

        IZipArchive&       archive() override { return m_archive; }
        const IZipArchive& archive() const override { return m_archive; }

        XLContentTypes& contentTypes() override
        {
            throw XLException("MockPackageServices::contentTypes not configured");
        }
        const XLContentTypes& contentTypes() const override
        {
            throw XLException("MockPackageServices::contentTypes not configured");
        }

        XLXmlData* findXmlPart(std::string_view, bool doNotThrow = false) override
        {
            if (doNotThrow) return nullptr;
            throw XLException("MockPackageServices::findXmlPart not configured");
        }
        const XLXmlData* findXmlPart(std::string_view, bool doNotThrow = false) const override
        {
            if (doNotThrow) return nullptr;
            throw XLException("MockPackageServices::findXmlPart not configured");
        }

        XLXmlData* emplaceXmlPart(const std::string&, const std::string&, XLContentType) override
        {
            throw XLException("MockPackageServices::emplaceXmlPart not configured");
        }

        XLRelationships& workbookRelationships() override
        {
            throw XLException("MockPackageServices::workbookRelationships not configured");
        }

        XLRelationships xmlRelationships(std::string_view) override
        {
            throw XLException("MockPackageServices::xmlRelationships not configured");
        }

        XLRelationships drawingRelationships(std::string_view) override
        {
            throw XLException("MockPackageServices::drawingRelationships not configured");
        }

    private:
        InMemoryZipBackend m_backend;
        IZipArchive        m_archive;
    };

}    // namespace OpenXLSX::test

#endif    // OPENXLSX_TEST_MOCKPACKAGEPORTS_HPP

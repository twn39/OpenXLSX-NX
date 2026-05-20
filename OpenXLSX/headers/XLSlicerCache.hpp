#ifndef OPENXLSX_XLSLICERCACHE_HPP
#define OPENXLSX_XLSLICERCACHE_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include <string>
#include <string_view>
#include <vector>

namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLSlicerCache final : public XLXmlFile
    {
    public:
        /**
         * @brief Default constructor.
         */
        XLSlicerCache() : XLXmlFile(nullptr) {}

        /**
         * @brief Constructor that wraps an XLXmlData pointer.
         */
        explicit XLSlicerCache(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

        ~XLSlicerCache() = default;
        XLSlicerCache(const XLSlicerCache& other) = default;
        XLSlicerCache(XLSlicerCache&& other) noexcept = default;
        XLSlicerCache& operator=(const XLSlicerCache& other) = default;
        XLSlicerCache& operator=(XLSlicerCache&& other) noexcept = default;

        /**
         * @brief Get the name of the slicer cache.
         */
        std::string name() const;

        /**
         * @brief Set the name of the slicer cache.
         */
        void setName(std::string_view name);

        /**
         * @brief Get the source name (e.g. column name or pivot field name).
         */
        std::string sourceName() const;

        /**
         * @brief Set the source name.
         */
        void setSourceName(std::string_view sourceName);

        /**
         * @brief Synchronize selected filtering items with the pivot cache definition.
         * @param pivotCache The pivot table's cache definition containing all fields and shared items.
         * @param selectedItems List of items that should be selected/filtered. If empty, all items are selected.
         */
        void syncWithPivotCache(const class XLPivotCacheDefinition& pivotCache, const std::vector<std::string>& selectedItems);
    };
} // namespace OpenXLSX

#endif // OPENXLSX_XLSLICERCACHE_HPP

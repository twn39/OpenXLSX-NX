#ifndef OPENXLSX_XLCONTENTTYPES_HPP
#define OPENXLSX_XLCONTENTTYPES_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <cstdint>    // uint8_t
#include <gsl/pointers>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @brief
     */
    enum class XLContentType : uint8_t {
        Workbook,
        Relationships,
        WorkbookMacroEnabled,
        Worksheet,
        Chartsheet,
        ExternalLink,
        Theme,
        Styles,
        SharedStrings,
        Drawing,
        Chart,
        ChartStyle,
        ChartColorStyle,
        ControlProperties,
        CalculationChain,
        VBAProject,
        CoreProperties,
        ExtendedProperties,
        CustomProperties,
        Comments,
        Table,
        VMLDrawing,
        Hyperlink,
        Unknown,
        PivotTable,
        PivotCacheDefinition,
        PivotCacheRecords
    };

    /**
     * @brief utility function: determine the name of an XLContentType value
     * @param type the XLContentType to get a name for
     * @return a string with the name of type
     */
    std::string XLContentTypeToString(XLContentType type);

    class OPENXLSX_EXPORT XLContentItem
    {
        friend class XLContentTypes;

    public:    // ---------- Public Member Functions ---------- //
        XLContentItem();

        explicit XLContentItem(const XMLNode& node);

        ~XLContentItem() = default;

        XLContentItem(const XLContentItem& other);

        XLContentItem(XLContentItem&& other) noexcept;

        XLContentItem& operator=(const XLContentItem& other);

        XLContentItem& operator=(XLContentItem&& other) noexcept;

        XLContentType type() const;

        std::string path() const;

    private:
        std::unique_ptr<XMLNode> m_contentNode; /**< */
    };

    // XLContentTypes Class

    /**
     * @brief The purpose of this class is to load, store add and save item in the [Content_Types].xml file.
     */
    class OPENXLSX_EXPORT XLContentTypes : public XLXmlFile
    {
    public:    // ---------- Public Member Functions ---------- //
        XLContentTypes();

        explicit XLContentTypes(gsl::not_null<XLXmlData*> xmlData);

        ~XLContentTypes() = default;

        XLContentTypes(const XLContentTypes& other);

        XLContentTypes(XLContentTypes&& other) noexcept;

        XLContentTypes& operator=(const XLContentTypes& other);

        XLContentTypes& operator=(XLContentTypes&& other) noexcept;

        /**
         * @brief Check if a default extension exists.
         * @param extension The extension
         * @return true if it exists
         */
        bool hasDefault(std::string_view extension) const;

        /**
         * @brief Check if an override exists.
         * @param path The path
         * @return true if it exists
         */
        bool hasOverride(std::string_view path) const;

        /**
         * @brief Add a new default extension key/value pair to the data store.
         * @param extension The extension
         * @param contentType The content type
         */
        void addDefault(std::string_view extension, std::string_view contentType);

        /**
         * @brief Add a new override key/getValue pair to the data store.
         * @param path The key
         * @param type The getValue
         */
        void addOverride(std::string_view path, XLContentType type);

        /**
         * @brief Update an existing override key/getValue pair or add if it does not exist.
         * @param path The key
         * @param type The getValue
         */
        void updateOverride(std::string_view path, XLContentType type);

        void deleteOverride(std::string_view path);

        void deleteOverride(const XLContentItem& item);

        XLContentItem contentItem(std::string_view path);

        std::vector<XLContentItem> getContentItems();

    private:    // ---------- Private Member Variables ---------- //
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLCONTENTTYPES_HPP

#ifndef OPENXLSX_XLCONTENTTYPES_HPP
#define OPENXLSX_XLCONTENTTYPES_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <cstdint>    // uint8_t
#include <memory>
#include <string>
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
        Unknown
    };

    /**
     * @brief utility function: determine the name of an XLContentType value
     * @param type the XLContentType to get a name for
     * @return a string with the name of type
     */
    std::string XLContentTypeToString(XLContentType type);

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLContentItem
    {
        friend class XLContentTypes;

    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLContentItem();

        /**
         * @brief
         * @param node
         */
        explicit XLContentItem(const XMLNode& node);

        /**
         * @brief
         */
        ~XLContentItem() = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLContentItem(const XLContentItem& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLContentItem(XLContentItem&& other) noexcept;

        /**
         * @brief
         * @param other
         * @return
         */
        XLContentItem& operator=(const XLContentItem& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLContentItem& operator=(XLContentItem&& other) noexcept;

        /**
         * @brief
         * @return
         */
        XLContentType type() const;

        /**
         * @brief
         * @return
         */
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
        /**
         * @brief
         */
        XLContentTypes();

        /**
         * @brief
         * @param xmlData
         */
        explicit XLContentTypes(XLXmlData* xmlData);

        /**
         * @brief Destructor
         */
        ~XLContentTypes() = default;

        /**
         * @brief
         * @param other
         */
        XLContentTypes(const XLContentTypes& other);

        /**
         * @brief
         * @param other
         */
        XLContentTypes(XLContentTypes&& other) noexcept;

        /**
         * @brief
         * @param other
         * @return
         */
        XLContentTypes& operator=(const XLContentTypes& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLContentTypes& operator=(XLContentTypes&& other) noexcept;

        /**
         * @brief Check if a default extension exists.
         * @param extension The extension
         * @return true if it exists
         */
        bool hasDefault(const std::string& extension) const;

        /**
         * @brief Check if an override exists.
         * @param path The path
         * @return true if it exists
         */
        bool hasOverride(const std::string& path) const;

        /**
         * @brief Add a new default extension key/value pair to the data store.
         * @param extension The extension
         * @param contentType The content type
         */
        void addDefault(const std::string& extension, const std::string& contentType);

        /**
         * @brief Add a new override key/getValue pair to the data store.
         * @param path The key
         * @param type The getValue
         */
        void addOverride(const std::string& path, XLContentType type);

        /**
         * @brief
         * @param path
         */
        void deleteOverride(const std::string& path);

        /**
         * @brief
         * @param item
         */
        void deleteOverride(const XLContentItem& item);

        /**
         * @brief
         * @param path
         * @return
         */
        XLContentItem contentItem(const std::string& path);

        /**
         * @brief
         * @return
         */
        std::vector<XLContentItem> getContentItems();

    private:    // ---------- Private Member Variables ---------- //
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLCONTENTTYPES_HPP

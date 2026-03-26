#ifndef OPENXLSX_XLRELATIONSHIPS_HPP
#define OPENXLSX_XLRELATIONSHIPS_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <gsl/pointers>
#include <random>    // std::mt19937
#include <string>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @brief Enable use of random (relationship) IDs
     */
    void UseRandomIDs();

    /**
     * @brief Disable use of random (relationship) IDs (default behavior)
     */
    void UseSequentialIDs();

    /**
     * @brief Return a 32 bit random value
     * @return A 32 bit random value
     */
    extern thread_local std::mt19937 Rand32;

    /**
     * @brief Return a 64 bit random value (by invoking Rand32 twice)
     * @return A 64 bit random value
     */
    uint64_t Rand64();

    /**
     * @brief Initialize XLRand32 data source
     * @param pseudoRandom If true, sequence will be reproducible with a constant seed
     */
    void InitRandom(bool pseudoRandom = false);

    class XLRelationships;

    class XLRelationshipItem;

    /**
     * @brief An enum of the possible relationship (or XML document) types used in relationship (.rels) XML files.
     */
    enum class XLRelationshipType {
        CoreProperties,
        ExtendedProperties,
        CustomProperties,
        Workbook,
        Worksheet,
        Chartsheet,
        Dialogsheet,
        Macrosheet,
        CalculationChain,
        ExternalLink,
        ExternalLinkPath,
        Theme,
        Styles,
        Chart,
        ChartStyle,
        ChartColorStyle,
        Image,
        Drawing,
        VMLDrawing,
        SharedStrings,
        PrinterSettings,
        VBAProject,
        ControlProperties,
        Comments,
        Table,
        Hyperlink,
        Unknown,
        PivotTable,
        Slicer,
        SlicerCache,
        PivotCacheDefinition,
        PivotCacheRecords
    };
}    //     namespace OpenXLSX

namespace OpenXLSX_XLRelationships
{    // special namespace to avoid naming conflict with another GetStringFromType function
    using namespace OpenXLSX;
    /**
     * @brief helper function, used only within module and from XLProperties.cpp / XLAppProperties::createFromTemplate
     * @param type the XLRelationshipType for which to return the correct XML string
     */
    std::string GetStringFromType(XLRelationshipType type);
}    //    namespace OpenXLSX_XLRelationships

namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLRelationshipItem
    {
    public:
        XLRelationshipItem() = default;

        /**
         * @param node The XML element defining the relationship.
         */
        explicit XLRelationshipItem(const XMLNode& node);

        ~XLRelationshipItem() = default;
        XLRelationshipItem(const XLRelationshipItem& other) = default;
        XLRelationshipItem(XLRelationshipItem&& other) noexcept = default;
        XLRelationshipItem& operator=(const XLRelationshipItem& other) = default;
        XLRelationshipItem& operator=(XLRelationshipItem&& other) noexcept = default;

        [[nodiscard]] XLRelationshipType type() const;

        /**
         * @brief Access the target URI or internal package path.
         */
        [[nodiscard]] std::string target() const;

        [[nodiscard]] std::string id() const;

        [[nodiscard]] bool empty() const;

    private:
        XMLNode m_relationshipNode; /**< Underlying XML element. Held by value. */
    };

    class OPENXLSX_EXPORT XLRelationships : public XLXmlFile
    {
    public:
        XLRelationships() = default;

        /**
         * @param xmlData Managed XML data source.
         * @param pathTo Path to the .rels file, used to resolve relative targets to absolute paths.
         */
        explicit XLRelationships(gsl::not_null<XLXmlData*> xmlData, std::string pathTo);

        ~XLRelationships();

        XLRelationships(const XLRelationships& other) = default;
        XLRelationships(XLRelationships&& other) noexcept = default;
        XLRelationships& operator=(const XLRelationships& other) = default;
        XLRelationships& operator=(XLRelationships&& other) noexcept = default;

        [[nodiscard]] XLRelationshipItem relationshipById(std::string_view id) const;

        /**
         * @param target The Target path to look up.
         * @param throwIfNotFound If false, returns an empty XLRelationshipItem on failure.
         */
        [[nodiscard]] XLRelationshipItem relationshipByTarget(std::string_view target, bool throwIfNotFound = true) const;

        [[nodiscard]] std::vector<XLRelationshipItem> relationships() const;

        void deleteRelationship(std::string_view relID);
        void deleteRelationship(const XLRelationshipItem& item);

        /**
         * @brief Create and append a new relationship.
         * @param isExternal Marking as 'External' prevents path normalization issues for URIs.
         */
        XLRelationshipItem addRelationship(XLRelationshipType type, std::string_view target, bool isExternal = false);

        [[nodiscard]] bool targetExists(std::string_view target) const;
        [[nodiscard]] bool idExists(std::string_view id) const;

        /**
         * @brief Serialize the underlying XML to a stream.
         */
        void print(std::basic_ostream<char>& ostr) const;

    private:
        std::string m_path; /**< Internal package path to the relationships file. */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLRELATIONSHIPS_HPP

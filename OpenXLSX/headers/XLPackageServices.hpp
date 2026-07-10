#ifndef OPENXLSX_XLPACKAGESERVICES_HPP
#define OPENXLSX_XLPACKAGESERVICES_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

/**
 * @file XLPackageServices.hpp
 * @brief Narrow package-level ports for XML parts that must touch the OPC package
 *        without depending on the full XLDocument façade (open/save/command bus, etc.).
 *
 * @details Typical consumers are Drawing / Pivot / Slicer helpers that only need:
 *          - ZIP archive I/O
 *          - managed XML part lookup / creation
 *          - relationship part access
 *          - [Content_Types].xml registry
 *
 *          Prefer `XLXmlFile::package()` over `parentDoc()` when only these services are required.
 *          Prefer `XLXmlFile::partFactory()` for createChart / sheetDrawing / addImage / slicers.
 *          Keep `parentDoc()` for workbook model, styles, persons, and command/query APIs.
 */

#include "IZipArchive.hpp"
#include "OpenXLSX-Exports.hpp"
#include "XLContentTypes.hpp"
#include "XLRelationships.hpp"

#include <cstdint>
#include <string>
#include <string_view>

namespace OpenXLSX
{
    class XLXmlData;

    /**
     * @brief Abstract package services port (archive + XML parts + relationships + content types).
     *
     * @note Implementations must outlive any object that stores a pointer/reference to them.
     *       `XLDocument` is the primary implementation.
     */
    class OPENXLSX_EXPORT XLPackageServices
    {
    public:
        virtual ~XLPackageServices() = default;

        // ----- Archive (ZIP / OPC binary store) -----

        [[nodiscard]] virtual IZipArchive&       archive()       = 0;
        [[nodiscard]] virtual const IZipArchive& archive() const = 0;

        // ----- Content types registry ([Content_Types].xml) -----

        [[nodiscard]] virtual XLContentTypes&       contentTypes()       = 0;
        [[nodiscard]] virtual const XLContentTypes& contentTypes() const = 0;

        // ----- Managed XML parts -----

        /**
         * @brief Locate a managed XML part by package path (e.g. "xl/pivotCache/pivotCacheDefinition1.xml").
         * @param doNotThrow When true, return nullptr if missing instead of throwing.
         */
        [[nodiscard]] virtual XLXmlData*       findXmlPart(std::string_view path, bool doNotThrow = false)       = 0;
        [[nodiscard]] virtual const XLXmlData* findXmlPart(std::string_view path, bool doNotThrow = false) const = 0;

        /**
         * @brief Create and register a new managed XML part in the document's part list.
         */
        virtual XLXmlData* emplaceXmlPart(const std::string& path, const std::string& id, XLContentType type) = 0;

        // ----- Relationships -----

        [[nodiscard]] virtual XLRelationships& workbookRelationships() = 0;

        /**
         * @brief Relationships for an arbitrary package XML part (…/_rels/name.xml.rels).
         */
        [[nodiscard]] virtual XLRelationships xmlRelationships(std::string_view xmlPath) = 0;

        /**
         * @brief Relationships for a drawing part.
         */
        [[nodiscard]] virtual XLRelationships drawingRelationships(std::string_view drawingPath) = 0;
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLPACKAGESERVICES_HPP

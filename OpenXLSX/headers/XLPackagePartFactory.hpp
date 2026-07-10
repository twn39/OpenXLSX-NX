#ifndef OPENXLSX_XLPACKAGEPARTFACTORY_HPP
#define OPENXLSX_XLPACKAGEPARTFACTORY_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

/**
 * @file XLPackagePartFactory.hpp
 * @brief Narrow factory port for creating / resolving OPC package parts used by worksheets.
 *
 * @details Worksheet drawing / chart / comment / slicer helpers often only need part factories
 *          (createChart, sheetDrawing, addImage, createSlicer, …) rather than the full
 *          XLDocument open/save/command façade. Prefer `XLXmlFile::partFactory()` for those paths.
 *
 *          Use `XLPackageServices` for archive / managed XML / relationships / content types.
 *          Keep `parentDoc()` for workbook model, styles, persons, and command/query APIs.
 */

#include "OpenXLSX-Exports.hpp"
#include "XLChart.hpp"
#include "XLComments.hpp"
#include "XLDrawing.hpp"
#include "XLRelationships.hpp"
#include "XLTables.hpp"
#include "XLThreadedComments.hpp"

#include <cstdint>
#include <gsl/span>
#include <string>
#include <string_view>

namespace OpenXLSX
{
    /**
     * @brief Abstract factory for worksheet-facing package parts.
     * @note Implemented by XLDocument. Lifetime must outlive consumers.
     * @note XLVmlDrawing is defined in XLDrawing.hpp (included above).
     */
    class OPENXLSX_EXPORT XLPackagePartFactory
    {
    public:
        virtual ~XLPackagePartFactory() = default;

        // ----- Sheet accessory presence -----

        [[nodiscard]] virtual bool hasSheetRelationships(uint16_t sheetXmlNo, bool isChartsheet = false) const = 0;
        [[nodiscard]] virtual bool hasSheetVmlDrawing(uint16_t sheetXmlNo) const                               = 0;
        [[nodiscard]] virtual bool hasSheetComments(uint16_t sheetXmlNo) const                                 = 0;
        [[nodiscard]] virtual bool hasSheetThreadedComments(uint16_t sheetXmlNo) const                         = 0;
        [[nodiscard]] virtual bool hasSheetDrawing(uint16_t sheetXmlNo) const                                  = 0;
        [[nodiscard]] virtual bool hasSheetTables(uint16_t sheetXmlNo) const                                   = 0;

        // ----- Sheet accessory access / creation -----

        [[nodiscard]] virtual XLRelationships sheetRelationships(uint16_t sheetXmlNo, bool isChartsheet = false) = 0;
        [[nodiscard]] virtual XLDrawing       sheetDrawing(uint16_t sheetXmlNo)                                   = 0;
        [[nodiscard]] virtual XLDrawing       createDrawing()                                                     = 0;
        [[nodiscard]] virtual XLDrawing       drawing(std::string_view path)                                      = 0;
        [[nodiscard]] virtual XLVmlDrawing    sheetVmlDrawing(uint16_t sheetXmlNo)                                = 0;
        [[nodiscard]] virtual XLComments      sheetComments(uint16_t sheetXmlNo)                                  = 0;
        [[nodiscard]] virtual XLThreadedComments sheetThreadedComments(uint16_t sheetXmlNo)                      = 0;
        [[nodiscard]] virtual XLTables        sheetTables(uint16_t sheetXmlNo)                                    = 0;

        // ----- Charts & media -----

        [[nodiscard]] virtual XLChart     createChart(XLChartType type = XLChartType::Bar) = 0;
        [[nodiscard]] virtual std::string addImage(std::string_view name, std::string_view data) = 0;
        [[nodiscard]] virtual std::string addImage(std::string_view name, gsl::span<const uint8_t> data) = 0;
        [[nodiscard]] virtual std::string getImage(std::string_view path) const = 0;

        // ----- Tables / slicers -----

        [[nodiscard]] virtual uint32_t nextTableId() const = 0;

        [[nodiscard]] virtual std::string createTableSlicerCache(uint32_t         tableId,
                                                                 uint32_t         tableColumnId,
                                                                 std::string_view name,
                                                                 std::string_view sourceName) = 0;
        [[nodiscard]] virtual std::string findOrCreateTableSlicerCache(uint32_t         tableId,
                                                                       uint32_t         tableColumnId,
                                                                       std::string_view name,
                                                                       std::string_view sourceName) = 0;

        [[nodiscard]] virtual std::string createPivotSlicerCache(uint32_t         pivotCacheId,
                                                                 uint32_t         sheetId,
                                                                 std::string_view pivotTableName,
                                                                 std::string_view name,
                                                                 std::string_view sourceName) = 0;
        [[nodiscard]] virtual std::string findOrCreatePivotSlicerCache(uint32_t         pivotCacheId,
                                                                       uint32_t         sheetId,
                                                                       std::string_view pivotTableName,
                                                                       std::string_view name,
                                                                       std::string_view sourceName) = 0;

        [[nodiscard]] virtual std::string createSlicer(std::string_view name,
                                                       std::string_view cacheName,
                                                       std::string_view caption,
                                                       std::string_view existingSlicerFile = {}) = 0;

        virtual void deleteSlicerFileAndOrphanCache(const std::string& name) = 0;
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLPACKAGEPARTFACTORY_HPP

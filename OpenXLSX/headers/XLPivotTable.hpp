#ifndef OPENXLSX_XLPIVOTTABLE_HPP
#define OPENXLSX_XLPIVOTTABLE_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include <string>
#include <string_view>
#include <vector>

namespace OpenXLSX
{
    enum class XLPivotSubtotal { Sum, Average, Count, Max, Min, Product };

    struct XLPivotField {
        std::string name;             
        XLPivotSubtotal subtotal = XLPivotSubtotal::Sum;     
        std::string customName;       
    };

    struct XLPivotTableOptions {
        std::string name;             // e.g., "PivotTable1"
        std::string sourceRange;      // e.g., "Sheet1!$A$1:$D$100"
        std::string targetCell;       // e.g., "A1" (will be placed in the calling worksheet)
        
        std::vector<XLPivotField> rows;
        std::vector<XLPivotField> columns;
        std::vector<XLPivotField> data;
        std::vector<XLPivotField> filters;
    };

    class OPENXLSX_EXPORT XLPivotTable final : public XLXmlFile
    {
    public:
        class XLRelationships relationships();

        XLPivotTable() : XLXmlFile(nullptr) {}
        explicit XLPivotTable(XLXmlData* xmlData);
        ~XLPivotTable() = default;
        XLPivotTable(const XLPivotTable& other) = default;
        XLPivotTable(XLPivotTable&& other) noexcept = default;
        XLPivotTable& operator=(const XLPivotTable& other) = default;
        XLPivotTable& operator=(XLPivotTable&& other) noexcept = default;
        
        void setName(std::string_view name);
    };

    class OPENXLSX_EXPORT XLPivotCacheDefinition final : public XLXmlFile
    {
    public:
        class XLRelationships relationships();

        XLPivotCacheDefinition() : XLXmlFile(nullptr) {}
        explicit XLPivotCacheDefinition(XLXmlData* xmlData);
        ~XLPivotCacheDefinition() = default;
        XLPivotCacheDefinition(const XLPivotCacheDefinition& other) = default;
        XLPivotCacheDefinition(XLPivotCacheDefinition&& other) noexcept = default;
        XLPivotCacheDefinition& operator=(const XLPivotCacheDefinition& other) = default;
        XLPivotCacheDefinition& operator=(XLPivotCacheDefinition&& other) noexcept = default;
    };
    
    class OPENXLSX_EXPORT XLPivotCacheRecords final : public XLXmlFile
    {
    public:
        XLPivotCacheRecords() : XLXmlFile(nullptr) {}
        explicit XLPivotCacheRecords(XLXmlData* xmlData);
        ~XLPivotCacheRecords() = default;
        XLPivotCacheRecords(const XLPivotCacheRecords& other) = default;
        XLPivotCacheRecords(XLPivotCacheRecords&& other) noexcept = default;
        XLPivotCacheRecords& operator=(const XLPivotCacheRecords& other) = default;
        XLPivotCacheRecords& operator=(XLPivotCacheRecords&& other) noexcept = default;
    };
}

#endif // OPENXLSX_XLPIVOTTABLE_HPP

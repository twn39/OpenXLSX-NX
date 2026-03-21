// ===== External Includes ===== //
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLPivotTable.hpp"
#include "XLXmlData.hpp"
#include "XLContentTypes.hpp"
#include "XLDocument.hpp"
#include "XLRelationships.hpp"
#include "XLException.hpp"

namespace OpenXLSX
{
    XLPivotTable::XLPivotTable(XLXmlData* xmlData) : XLXmlFile(xmlData)
    {
        if (xmlData && xmlData->getXmlType() != XLContentType::PivotTable) {
            throw XLInternalError("XLPivotTable constructor: Invalid XML data.");
        }
    }
    
    void XLPivotTable::setName(std::string_view name) {
        xmlDocument().document_element().attribute("name").set_value(std::string(name).c_str());
    }

    XLPivotCacheDefinition::XLPivotCacheDefinition(XLXmlData* xmlData) : XLXmlFile(xmlData)
    {
        if (xmlData && xmlData->getXmlType() != XLContentType::PivotCacheDefinition) {
            throw XLInternalError("XLPivotCacheDefinition constructor: Invalid XML data.");
        }
    }

    XLPivotCacheRecords::XLPivotCacheRecords(XLXmlData* xmlData) : XLXmlFile(xmlData)
    {
        if (xmlData && xmlData->getXmlType() != XLContentType::PivotCacheRecords) {
            throw XLInternalError("XLPivotCacheRecords constructor: Invalid XML data.");
        }
    }
}


namespace OpenXLSX {
    XLRelationships XLPivotTable::relationships() {
        return parentDoc().xmlRelationships(getXmlPath());
    }
    
    XLRelationships XLPivotCacheDefinition::relationships() {
        return parentDoc().xmlRelationships(getXmlPath());
    }
}

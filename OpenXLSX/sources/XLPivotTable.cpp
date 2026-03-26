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
    

    std::string XLPivotTable::name() const {
        return xmlDocument().document_element().attribute("name").value();
    }

    void XLPivotTable::setRefreshOnLoad(bool refresh) {
        // We will just let the user set it. But without parent pointer, it's hard to get the cache.
        // Actually we DO have parent document pointer!
        // XLPivotTable inherits XLXmlFile which has parentDoc().
        
        XMLNode root = xmlDocument().document_element();
        std::string cacheIdStr = root.attribute("cacheId").value();
        
        // Find in workbook pivot caches
        XMLNode wbRoot = parentDoc().workbook().xmlDocument().document_element();
        XMLNode pivotCaches = wbRoot.child("pivotCaches");
        std::string rId = "";
        for (auto pc : pivotCaches.children("pivotCache")) {
            if (std::string(pc.attribute("cacheId").value()) == cacheIdStr) {
                rId = pc.attribute("r:id").value();
                break;
            }
        }
        
        if (!rId.empty()) {
            std::string cacheTargetPath = parentDoc().workbookRelationships().relationshipById(rId).target();
            // Resolve relative path to absolute
            if (cacheTargetPath[0] != '/') cacheTargetPath = "/xl/" + cacheTargetPath;
            
            // Get the xml document for cache
            XLPivotCacheDefinition cacheDef(parentDoc().getXmlData(cacheTargetPath.substr(1)));
            XMLDocument& cacheDoc = cacheDef.xmlDocument();
            XMLNode cacheRoot = cacheDoc.document_element();
            if (refresh) {
                if (!cacheRoot.attribute("refreshOnLoad")) cacheRoot.append_attribute("refreshOnLoad").set_value("1");
                else cacheRoot.attribute("refreshOnLoad").set_value("1");
            } else {
                if (cacheRoot.attribute("refreshOnLoad")) cacheRoot.remove_attribute("refreshOnLoad");
            }
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

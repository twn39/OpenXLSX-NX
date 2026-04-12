// ===== External Includes ===== //
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLContentTypes.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLPivotTable.hpp"
#include "XLRelationships.hpp"
#include "XLXmlData.hpp"

namespace OpenXLSX
{
    XLPivotTable::XLPivotTable(XLXmlData* xmlData) : XLXmlFile(xmlData)
    {
        if (xmlData && xmlData->getXmlType() != XLContentType::PivotTable) {
            throw XLInternalError("XLPivotTable constructor: Invalid XML data.");
        }
    }

    std::string XLPivotTable::name() const { return xmlDocument().document_element().attribute("name").value(); }

    std::string XLPivotTable::targetCell() const
    {
        XMLNode locNode = xmlDocument().document_element().child("location");
        if (locNode) return locNode.attribute("ref").value();
        return "";
    }

    XLPivotCacheDefinition XLPivotTable::cacheDefinition() const
    {
        XMLNode     root       = xmlDocument().document_element();
        std::string cacheIdStr = root.attribute("cacheId").value();

        // Find in workbook pivot caches
        XMLNode     wbRoot      = parentDoc().workbook().xmlDocument().document_element();
        XMLNode     pivotCaches = wbRoot.child("pivotCaches");
        std::string rId         = "";
        for (auto pc : pivotCaches.children("pivotCache")) {
            if (std::string(pc.attribute("cacheId").value()) == cacheIdStr) {
                rId = pc.attribute("r:id").value();
                break;
            }
        }

        if (!rId.empty()) {
            std::string cacheTargetPath = const_cast<XLDocument&>(parentDoc()).workbookRelationships().relationshipById(rId).target();
            if (!cacheTargetPath.empty() && cacheTargetPath[0] != '/') cacheTargetPath = "/xl/" + cacheTargetPath;
            std::string targetPath = !cacheTargetPath.empty() && cacheTargetPath[0] == '/' ? cacheTargetPath.substr(1) : cacheTargetPath;
            
            XLXmlData* xmlData = const_cast<XLDocument&>(parentDoc()).getXmlData(targetPath, true);
            if (xmlData == nullptr) {
                xmlData = &const_cast<XLDocument&>(parentDoc()).m_data.emplace_back(&const_cast<XLDocument&>(parentDoc()), targetPath, rId, XLContentType::PivotCacheDefinition);
            }
            return XLPivotCacheDefinition(xmlData);
        }
        throw XLInternalError("Could not locate pivot cache definition for pivot table.");
    }

    std::string XLPivotTable::sourceRange() const
    {
        return cacheDefinition().sourceRange();
    }

    void XLPivotTable::changeSourceRange(std::string_view newRange)
    {
        cacheDefinition().changeSourceRange(newRange);
    }

    void XLPivotTable::setRefreshOnLoad(bool refresh)
    {
        // We will just let the user set it. But without parent pointer, it's hard to get the cache.
        // Actually we DO have parent document pointer!
        // XLPivotTable inherits XLXmlFile which has parentDoc().

        XMLNode     root       = xmlDocument().document_element();
        std::string cacheIdStr = root.attribute("cacheId").value();

        // Find in workbook pivot caches
        XMLNode     wbRoot      = parentDoc().workbook().xmlDocument().document_element();
        XMLNode     pivotCaches = wbRoot.child("pivotCaches");
        std::string rId         = "";
        for (auto pc : pivotCaches.children("pivotCache")) {
            if (std::string(pc.attribute("cacheId").value()) == cacheIdStr) {
                rId = pc.attribute("r:id").value();
                break;
            }
        }

        if (!rId.empty()) {
            std::string cacheTargetPath = parentDoc().workbookRelationships().relationshipById(rId).target();
            // Resolve relative path to absolute
            if (!cacheTargetPath.empty() && cacheTargetPath[0] != '/') cacheTargetPath = "/xl/" + cacheTargetPath;
            std::string targetPath = !cacheTargetPath.empty() && cacheTargetPath[0] == '/' ? cacheTargetPath.substr(1) : cacheTargetPath;

            // Get the xml document for cache
            XLPivotCacheDefinition cacheDef(parentDoc().getXmlData(targetPath));
            XMLDocument&           cacheDoc  = cacheDef.xmlDocument();
            XMLNode                cacheRoot = cacheDoc.document_element();
            if (refresh) {
                if (!cacheRoot.attribute("refreshOnLoad"))
                    cacheRoot.append_attribute("refreshOnLoad").set_value("1");
                else
                    cacheRoot.attribute("refreshOnLoad").set_value("1");
            }
            else {
                if (cacheRoot.attribute("refreshOnLoad")) cacheRoot.remove_attribute("refreshOnLoad");
            }
        }
    }

    void XLPivotTable::setName(std::string_view name)
    { xmlDocument().document_element().attribute("name").set_value(std::string(name).c_str()); }

    XLPivotCacheDefinition::XLPivotCacheDefinition(XLXmlData* xmlData) : XLXmlFile(xmlData)
    {
        if (xmlData && xmlData->getXmlType() != XLContentType::PivotCacheDefinition) {
            throw XLInternalError("XLPivotCacheDefinition constructor: Invalid XML data.");
        }
    }

    std::string XLPivotCacheDefinition::sourceRange() const
    {
        XMLNode sourceNode = xmlDocument().document_element().child("cacheSource").child("worksheetSource");
        if (sourceNode) {
            std::string sheet = sourceNode.attribute("sheet").value();
            std::string ref   = sourceNode.attribute("ref").value();
            return sheet + "!" + ref;
        }
        return "";
    }

    void XLPivotCacheDefinition::changeSourceRange(std::string_view newRange)
    {
        XMLNode sourceNode = xmlDocument().document_element().child("cacheSource").child("worksheetSource");
        if (!sourceNode) throw XLInternalError("Could not locate worksheetSource in pivot cache.");

        std::string rangeStr(newRange);
        auto        bangPos = rangeStr.find('!');
        if (bangPos == std::string::npos) throw XLInputError("Invalid range format. Expected 'Sheet!A1:Z100'");

        std::string sheet = rangeStr.substr(0, bangPos);
        std::string ref   = rangeStr.substr(bangPos + 1);

        if (sheet.length() >= 2 && sheet.front() == '\'' && sheet.back() == '\'') {
            sheet = sheet.substr(1, sheet.length() - 2);
        }

        sourceNode.attribute("sheet").set_value(sheet.c_str());
        sourceNode.attribute("ref").set_value(ref.c_str());
    }

    XLPivotCacheRecords::XLPivotCacheRecords(XLXmlData* xmlData) : XLXmlFile(xmlData)
    {
        if (xmlData && xmlData->getXmlType() != XLContentType::PivotCacheRecords) {
            throw XLInternalError("XLPivotCacheRecords constructor: Invalid XML data.");
        }
    }
}    // namespace OpenXLSX

namespace OpenXLSX
{
    XLRelationships XLPivotTable::relationships() { return parentDoc().xmlRelationships(getXmlPath()); }

    XLRelationships XLPivotCacheDefinition::relationships() { return parentDoc().xmlRelationships(getXmlPath()); }
}    // namespace OpenXLSX

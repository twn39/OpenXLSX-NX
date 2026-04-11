#include "XLDocument.hpp"
#include "XLUtilities.hpp"
#include "XLWorkbook.hpp"
#include "XLWorksheet.hpp"
#include <unordered_set>

using namespace OpenXLSX;

std::vector<XLPivotTable> XLWorksheet::pivotTables()
{
    std::vector<XLPivotTable> result;
    XMLNode ptNode = xmlDocument().document_element().child("pivotTables");
    if (!ptNode.empty()) {
        for (auto pt : ptNode.children("pivotTable")) {
            std::string rId = pt.attribute("r:id").value();
            if (!rId.empty()) {
                std::string targetPath = relationships().relationshipById(rId).target();
                // Resolve to absolute path
                if (targetPath[0] != '/') {
                    std::string sourcePath = getXmlPath();
                    auto slashPos = sourcePath.find_last_of('/');
                    if (slashPos != std::string::npos) {
                        targetPath = sourcePath.substr(0, slashPos + 1) + targetPath;
                    } else {
                        targetPath = "/" + targetPath;
                    }
                }
                
                // Normalizing path
                while (targetPath.find("/../") != std::string::npos) {
                    auto pos = targetPath.find("/../");
                    auto prevSlash = targetPath.find_last_of('/', pos - 1);
                    if (prevSlash != std::string::npos) {
                        targetPath.erase(prevSlash + 1, pos - prevSlash + 3);
                    } else {
                        targetPath.erase(0, pos + 4);
                    }
                }
                
                if (targetPath[0] == '/') targetPath = targetPath.substr(1);

                XLXmlData* xmlData = const_cast<XLDocument&>(parentDoc()).getXmlData(targetPath, true);
                if (xmlData == nullptr) {
                    xmlData = &const_cast<XLDocument&>(parentDoc()).m_data.emplace_back(&const_cast<XLDocument&>(parentDoc()), targetPath, rId, XLContentType::PivotTable);
                }
                result.emplace_back(xmlData);
            }
        }
    }
    return result;
}

bool XLWorksheet::deletePivotTable(std::string_view name)
{
    XMLNode ptNode = xmlDocument().document_element().child("pivotTables");
    if (ptNode.empty()) return false;

    std::string targetRId;
    XMLNode targetNode;

    for (auto pt : ptNode.children("pivotTable")) {
        std::string rId = pt.attribute("r:id").value();
        if (!rId.empty()) {
            std::string targetPath = relationships().relationshipById(rId).target();
            if (targetPath[0] != '/') {
                std::string sourcePath = getXmlPath();
                auto slashPos = sourcePath.find_last_of('/');
                if (slashPos != std::string::npos) targetPath = sourcePath.substr(0, slashPos + 1) + targetPath;
            }
            
            while (targetPath.find("/../") != std::string::npos) {
                auto pos = targetPath.find("/../");
                auto prevSlash = targetPath.find_last_of('/', pos - 1);
                if (prevSlash != std::string::npos) targetPath.erase(prevSlash + 1, pos - prevSlash + 3);
                else targetPath.erase(0, pos + 4);
            }
            if (targetPath[0] == '/') targetPath = targetPath.substr(1);

            XLXmlData* xmlData = const_cast<XLDocument&>(parentDoc()).getXmlData(targetPath, true);
            if (xmlData == nullptr) {
                xmlData = &const_cast<XLDocument&>(parentDoc()).m_data.emplace_back(&const_cast<XLDocument&>(parentDoc()), targetPath, rId, XLContentType::PivotTable);
            }
            XLPivotTable ptObj(xmlData);
            if (ptObj.name() == name) {
                targetRId = rId;
                targetNode = pt;
                break;
            }
        }
    }

    if (!targetNode.empty()) {
        ptNode.remove_child(targetNode);
        if (ptNode.first_child().empty()) {
            xmlDocument().document_element().remove_child(ptNode);
        }
        relationships().deleteRelationship(targetRId);
        return true;
    }

    return false;
}

XLPivotTable XLWorksheet::addPivotTable(const XLPivotTableOptions& options)
{
    auto& doc = parentDoc();
    auto  wb  = parentDoc().workbook();

    XMLNode wbNode          = wb.xmlDocument().document_element();
    XMLNode pivotCachesNode = wbNode.child("pivotCaches");

    std::string targetSourceRange = options.sourceRange();
    auto bang = targetSourceRange.find('!');
    if (bang != std::string::npos) {
        std::string sheet = targetSourceRange.substr(0, bang);
        if (sheet.length() >= 2 && sheet.front() == '\'' && sheet.back() == '\'') {
            targetSourceRange = sheet.substr(1, sheet.length() - 2) + "!" + targetSourceRange.substr(bang + 1);
        }
    }

    uint32_t newCacheId = 0;
    XLPivotCacheDefinition cacheDef(nullptr);

    if (!pivotCachesNode.empty()) {
        for (auto cache : pivotCachesNode.children("pivotCache")) {
            std::string rId = cache.attribute("r:id").value();
            if (!rId.empty()) {
                std::string targetPath = doc.workbookRelationships().relationshipById(rId).target();
                if (targetPath[0] != '/') targetPath = "/xl/" + targetPath;

                while (targetPath.find("/../") != std::string::npos) {
                    auto pos = targetPath.find("/../");
                    auto prevSlash = targetPath.find_last_of('/', pos - 1);
                    if (prevSlash != std::string::npos) targetPath.erase(prevSlash + 1, pos - prevSlash + 3);
                    else targetPath.erase(0, pos + 4);
                }
                if (targetPath[0] == '/') targetPath = targetPath.substr(1);

                XLXmlData* xmlData = doc.getXmlData(targetPath, true);
                if (xmlData != nullptr) {
                    XLPivotCacheDefinition existingCache(xmlData);
                    std::string existingRange = existingCache.sourceRange();
                    auto exBang = existingRange.find('!');
                    if (exBang != std::string::npos) {
                        std::string sheet = existingRange.substr(0, exBang);
                        if (sheet.length() >= 2 && sheet.front() == '\'' && sheet.back() == '\'') {
                            existingRange = sheet.substr(1, sheet.length() - 2) + "!" + existingRange.substr(exBang + 1);
                        }
                    }
                    if (existingRange == targetSourceRange) {
                        newCacheId = cache.attribute("cacheId").as_uint();
                        cacheDef = existingCache;
                        break;
                    }
                }
            }
        }
    }

    bool isNewCache = (newCacheId == 0);

    if (isNewCache) {
        cacheDef = doc.createPivotCacheDefinition();

        std::string cacheDefRelPath = getPathARelativeToPathB(cacheDef.getXmlPath(), wb.getXmlPath());
        auto        cacheRel        = doc.workbookRelationships().addRelationship(XLRelationshipType::PivotCacheDefinition, cacheDefRelPath);

        if (pivotCachesNode.empty()) {
            pivotCachesNode = wbNode.insert_child_before("pivotCaches", wbNode.child("extLst"));
            if (pivotCachesNode.empty()) pivotCachesNode = wbNode.append_child("pivotCaches");
        }

        newCacheId = 1;
        for (auto cache : pivotCachesNode.children("pivotCache")) {
            uint32_t id = cache.attribute("cacheId").as_uint();
            if (id >= newCacheId) newCacheId = id + 1;
        }

        XMLNode pcNode = pivotCachesNode.append_child("pivotCache");
        pcNode.append_attribute("cacheId").set_value(newCacheId);
        pcNode.append_attribute("r:id").set_value(cacheRel.id().c_str());
    }

    auto pivotTable = doc.createPivotTable();

    std::string cacheDefRelPathFromPt = getPathARelativeToPathB(cacheDef.getXmlPath(), pivotTable.getXmlPath());
    pivotTable.relationships().addRelationship(XLRelationshipType::PivotCacheDefinition, cacheDefRelPathFromPt);

    std::string ptRelPath = getPathARelativeToPathB(pivotTable.getXmlPath(), getXmlPath());

    XMLNode ptRelPathNode = xmlDocument().document_element().child("pivotTables");
    if (ptRelPathNode.empty()) {
        XMLNode docElement = xmlDocument().document_element();
        ptRelPathNode      = appendAndGetNode(docElement, "pivotTables", m_nodeOrder);
    }

    XLRelationshipItem ptRel;
    if (!relationships().targetExists(ptRelPath)) { ptRel = relationships().addRelationship(XLRelationshipType::PivotTable, ptRelPath); }
    else {
        ptRel = relationships().relationshipByTarget(ptRelPath);
    }

    XMLNode ptEntry = ptRelPathNode.append_child("pivotTable");
    ptEntry.append_attribute("r:id").set_value(ptRel.id().c_str());
    ptEntry.append_attribute("cacheId").set_value(newCacheId);

    XMLNode cacheDefRoot = cacheDef.xmlDocument().document_element();

    std::string sourceSheet = "";
    std::string sourceRef   = targetSourceRange;
    size_t      exclaPos    = targetSourceRange.find('!');
    if (exclaPos != std::string::npos) {
        sourceSheet = targetSourceRange.substr(0, exclaPos);
        sourceRef   = targetSourceRange.substr(exclaPos + 1);
    }

    if (isNewCache) {
        XMLNode sourceNode   = cacheDefRoot.child("cacheSource").child("worksheetSource");
        sourceNode.attribute("ref").set_value(sourceRef.c_str());
        if (!sourceSheet.empty()) { sourceNode.attribute("sheet").set_value(sourceSheet.c_str()); }
    }

    XMLNode ptRoot = pivotTable.xmlDocument().document_element();

    ptRoot.attribute("name").set_value(options.name().empty() ? "PivotTable1" : options.name().c_str());
    ptRoot.attribute("cacheId").set_value(newCacheId);

    auto setBoolAttr = [&](XMLNode node, const char* attr, bool val) {
        if (!node.attribute(attr)) node.append_attribute(attr) = val ? "1" : "0";
        else node.attribute(attr).set_value(val ? "1" : "0");
    };
    
    setBoolAttr(ptRoot, "rowGrandTotals", options.rowGrandTotals());
    setBoolAttr(ptRoot, "colGrandTotals", options.colGrandTotals());
    setBoolAttr(ptRoot, "showDrill", options.showDrill());
    setBoolAttr(ptRoot, "useAutoFormatting", options.useAutoFormatting());
    setBoolAttr(ptRoot, "pageOverThenDown", options.pageOverThenDown());
    setBoolAttr(ptRoot, "mergeItem", options.mergeItem());
    setBoolAttr(ptRoot, "compactData", options.compactData());
    setBoolAttr(ptRoot, "showError", options.showError());
    
    XMLNode styleInfoNode = ptRoot.child("pivotTableStyleInfo");
    if (!styleInfoNode) styleInfoNode = ptRoot.append_child("pivotTableStyleInfo");
    
    if (!styleInfoNode.attribute("name")) styleInfoNode.append_attribute("name") = options.pivotTableStyleName().c_str();
    else styleInfoNode.attribute("name").set_value(options.pivotTableStyleName().c_str());
    
    setBoolAttr(styleInfoNode, "showRowHeaders", options.showRowHeaders());
    setBoolAttr(styleInfoNode, "showColHeaders", options.showColHeaders());
    setBoolAttr(styleInfoNode, "showRowStripes", options.showRowStripes());
    setBoolAttr(styleInfoNode, "showColStripes", options.showColStripes());
    setBoolAttr(styleInfoNode, "showLastColumn", options.showLastColumn());

    XMLNode locNode = ptRoot.child("location");
    locNode.attribute("ref").set_value(options.targetCell().c_str());

    XMLNode pivotFieldsNode = ptRoot.child("pivotFields");
    if (pivotFieldsNode.empty()) pivotFieldsNode = ptRoot.insert_child_after("pivotFields", ptRoot.child("location"));
    pivotFieldsNode.remove_children();

    std::vector<std::string> headers;
    
    if (isNewCache) {
        try {
            XLWorksheet     srcWks = sourceSheet.empty() ? wb.worksheet(name()) : wb.worksheet(sourceSheet);
            XLCellReference startRef(sourceRef.substr(0, sourceRef.find(':')));
            XLCellReference endRef(sourceRef.substr(sourceRef.find(':') + 1));

            for (uint16_t col = startRef.column(); col <= endRef.column(); ++col) {
                XLCellValue val        = srcWks.cell(startRef.row(), col).value();
                std::string headerName = val.type() == XLValueType::String ? val.get<std::string>() : "Field" + std::to_string(col);
                headers.push_back(headerName);
            }
        }
        catch (...) {
            // Fallback if parsing fails
            if (headers.empty()) headers.push_back("Field1");
        }

        XMLNode cacheFieldsNode = cacheDefRoot.child("cacheFields");
        if (cacheFieldsNode.empty()) cacheFieldsNode = cacheDefRoot.append_child("cacheFields");
        cacheFieldsNode.remove_children();    // clean up template
        if (!cacheFieldsNode.attribute("count")) cacheFieldsNode.append_attribute("count").set_value(headers.size());
        else cacheFieldsNode.attribute("count").set_value(headers.size());

        for (const auto& h : headers) {
            XMLNode fieldNode = cacheFieldsNode.append_child("cacheField");
            fieldNode.append_attribute("name").set_value(h.c_str());
            fieldNode.append_attribute("numFmtId").set_value("0");
            
            XMLNode sharedItemsNode = fieldNode.append_child("sharedItems");
            sharedItemsNode.append_attribute("containsBlank").set_value("1");
            sharedItemsNode.append_attribute("count").set_value("0");
            sharedItemsNode.append_child("m");
        }
    } else {
        XMLNode cacheFieldsNode = cacheDefRoot.child("cacheFields");
        for (auto cf : cacheFieldsNode.children("cacheField")) {
            headers.push_back(cf.attribute("name").value());
        }
    }

    // remove all old template children
    pivotFieldsNode.remove_children();
    XMLNode rowFieldsNode = ptRoot.child("rowFields");
    if (!rowFieldsNode.empty()) ptRoot.remove_child(rowFieldsNode);
    XMLNode rowItemsNode = ptRoot.child("rowItems");
    if (!rowItemsNode.empty()) ptRoot.remove_child(rowItemsNode);
    XMLNode colItemsNode = ptRoot.child("colItems");
    if (!colItemsNode.empty()) ptRoot.remove_child(colItemsNode);
    XMLNode pageFieldsNodeOld = ptRoot.child("pageFields");
    if (!pageFieldsNodeOld.empty()) ptRoot.remove_child(pageFieldsNodeOld);
    XMLNode dataFieldsNode = ptRoot.child("dataFields");
    if (!dataFieldsNode.empty()) ptRoot.remove_child(dataFieldsNode);

    // Rebuild proper

    // Rebuild proper pivot fields
    pivotFieldsNode.attribute("count").set_value(headers.size());

    std::vector<int> rowIndices;
    std::vector<int> colIndices;
    std::vector<int> dataIndices;
    std::vector<int> filterIndices;

    auto findFieldIndex = [&](const std::string& name) -> int {
        auto it = std::find(headers.begin(), headers.end(), name);
        if (it != headers.end()) return gsl::narrow_cast<int>(std::distance(headers.begin(), it));
        return -1;
    };

    for (const auto& rowFld : options.rows()) {
        int idx = findFieldIndex(rowFld.name);
        if (idx >= 0) rowIndices.push_back(idx);
    }
    for (const auto& colFld : options.columns()) {
        int idx = findFieldIndex(colFld.name);
        if (idx >= 0) colIndices.push_back(idx);
    }
    for (const auto& dataFld : options.data()) {
        int idx = findFieldIndex(dataFld.name);
        if (idx >= 0) dataIndices.push_back(idx);
    }
    for (const auto& filterFld : options.filters()) {
        int idx = findFieldIndex(filterFld.name);
        if (idx >= 0) filterIndices.push_back(idx);
    }

    int i = 0;
    for (const auto& h : headers) {
        (void)h;
        XMLNode ptFieldNode = pivotFieldsNode.append_child("pivotField");

        bool isRow  = std::find(rowIndices.begin(), rowIndices.end(), i) != rowIndices.end();
        bool isCol  = std::find(colIndices.begin(), colIndices.end(), i) != colIndices.end();
        bool isData = std::find(dataIndices.begin(), dataIndices.end(), i) != dataIndices.end();
        bool isFilter = std::find(filterIndices.begin(), filterIndices.end(), i) != filterIndices.end();

        if (isRow) {
            ptFieldNode.append_attribute("axis").set_value("axisRow");
            ptFieldNode.append_attribute("compact").set_value("0");
            ptFieldNode.append_attribute("outline").set_value("0");
            ptFieldNode.append_attribute("showAll").set_value("0");
            ptFieldNode.append_attribute("defaultSubtotal").set_value("1");
            XMLNode itemsNode = ptFieldNode.append_child("items");
            itemsNode.append_attribute("count").set_value("1");
            itemsNode.append_child("item").append_attribute("t").set_value("default");
        }
        else if (isCol) {
            ptFieldNode.append_attribute("axis").set_value("axisCol");
            ptFieldNode.append_attribute("compact").set_value("0");
            ptFieldNode.append_attribute("outline").set_value("0");
            ptFieldNode.append_attribute("showAll").set_value("0");
            ptFieldNode.append_attribute("defaultSubtotal").set_value("1");
            XMLNode itemsNode = ptFieldNode.append_child("items");
            itemsNode.append_attribute("count").set_value("1");
            itemsNode.append_child("item").append_attribute("t").set_value("default");
        }
        else if (isFilter) {
            ptFieldNode.append_attribute("axis").set_value("axisPage");
            ptFieldNode.append_attribute("compact").set_value("0");
            ptFieldNode.append_attribute("outline").set_value("0");
            ptFieldNode.append_attribute("showAll").set_value("0");
            ptFieldNode.append_attribute("defaultSubtotal").set_value("1");
            XMLNode itemsNode = ptFieldNode.append_child("items");
            itemsNode.append_attribute("count").set_value("1");
            itemsNode.append_child("item").append_attribute("t").set_value("default");
        }
        else if (isData) {
            ptFieldNode.append_attribute("dataField").set_value("1");
            ptFieldNode.append_attribute("showAll").set_value("0");
        }
        else {
            ptFieldNode.append_attribute("showAll").set_value("0");
        }
        i++;
    }

    bool hasVirtualDataField = (options.data().size() > 1);
    bool virtualDataOnRows = options.dataOnRows();

    if (!rowIndices.empty() || (hasVirtualDataField && virtualDataOnRows)) {
        rowFieldsNode = ptRoot.insert_child_after("rowFields", pivotFieldsNode);
        rowFieldsNode.append_attribute("count").set_value(rowIndices.size() + (hasVirtualDataField && virtualDataOnRows ? 1 : 0));
        for (int idx : rowIndices) { rowFieldsNode.append_child("field").append_attribute("x").set_value(idx); }
        if (hasVirtualDataField && virtualDataOnRows) {
            rowFieldsNode.append_child("field").append_attribute("x").set_value("-2"); // -2 represents the Values field in Excel
        }

        rowItemsNode = ptRoot.insert_child_after("rowItems", rowFieldsNode);
        rowItemsNode.append_attribute("count").set_value("1");
        rowItemsNode.append_child("i").append_child("x");
    }

    if (!colIndices.empty() || (hasVirtualDataField && !virtualDataOnRows)) {
        XMLNode insertAfter   = rowItemsNode.empty() ? (rowFieldsNode.empty() ? pivotFieldsNode : rowFieldsNode) : rowItemsNode;
        XMLNode colFieldsNode = ptRoot.insert_child_after("colFields", insertAfter);
        colFieldsNode.append_attribute("count").set_value(colIndices.size() + (hasVirtualDataField && !virtualDataOnRows ? 1 : 0));
        for (int idx : colIndices) { colFieldsNode.append_child("field").append_attribute("x").set_value(idx); }
        if (hasVirtualDataField && !virtualDataOnRows) {
            colFieldsNode.append_child("field").append_attribute("x").set_value("-2"); // -2 represents the Values field in Excel
        }

        colItemsNode = ptRoot.insert_child_after("colItems", colFieldsNode);
        colItemsNode.append_attribute("count").set_value("1");
        colItemsNode.append_child("i");
    }

    XMLNode pageFieldsNode;
    if (!filterIndices.empty()) {
        XMLNode insertAfter = colItemsNode.empty() ? (rowItemsNode.empty() ? (rowFieldsNode.empty() ? pivotFieldsNode : rowFieldsNode) : rowItemsNode) : colItemsNode;
        pageFieldsNode = ptRoot.insert_child_after("pageFields", insertAfter);
        pageFieldsNode.append_attribute("count").set_value(filterIndices.size());
        for (int idx : filterIndices) { 
            XMLNode pageField = pageFieldsNode.append_child("pageField");
            pageField.append_attribute("fld").set_value(idx);
        }
    }
    if (!dataIndices.empty()) {
        dataFieldsNode = ptRoot.insert_child_before("dataFields", ptRoot.child("pivotTableStyleInfo"));
        dataFieldsNode.append_attribute("count").set_value(options.data().size());

        for (const auto& dataFld : options.data()) {
            int idx = findFieldIndex(dataFld.name);
            if (idx < 0) continue;

            XMLNode dfield = dataFieldsNode.append_child("dataField");

            std::string dName = dataFld.customName.empty() ? ("Sum of " + dataFld.name) : dataFld.customName;
            dfield.append_attribute("name").set_value(dName.c_str());
            dfield.append_attribute("fld").set_value(idx);
            dfield.append_attribute("baseField").set_value("0"); // Required by Excel for valid schema mapping
            dfield.append_attribute("baseItem").set_value("0");

            if (dataFld.numFmtId != 0) {
                dfield.append_attribute("numFmtId").set_value(dataFld.numFmtId);
            }

            std::string subType = "sum";
            switch (dataFld.subtotal) {
                case XLPivotSubtotal::Average:
                    subType = "average";
                    break;
                case XLPivotSubtotal::Count:
                    subType = "count";
                    break;
                case XLPivotSubtotal::Max:
                    subType = "max";
                    break;
                case XLPivotSubtotal::Min:
                    subType = "min";
                    break;
                case XLPivotSubtotal::Product:
                    subType = "product";
                    break;
                default:
                    subType = "sum";
                    break;
            }
            dfield.append_attribute("subtotal").set_value(subType.c_str());
        }
    }

    locNode.attribute("ref").set_value(options.targetCell().c_str());
    return pivotTable;
}

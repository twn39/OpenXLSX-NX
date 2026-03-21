#include "XLWorksheet.hpp"
#include "XLDocument.hpp"
#include "XLWorkbook.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

XLPivotTable XLWorksheet::addPivotTable(const XLPivotTableOptions& options)
{
    auto& doc = parentDoc();
    auto wb = parentDoc().workbook();

    auto pivotTable = doc.createPivotTable();
    auto cacheDef = doc.createPivotCacheDefinition();
    auto cacheRec = doc.createPivotCacheRecords(cacheDef.getXmlPath());

    std::string recordsRelPath = getPathARelativeToPathB(cacheRec.getXmlPath(), cacheDef.getXmlPath());
    cacheDef.relationships().addRelationship(XLRelationshipType::PivotCacheRecords, recordsRelPath);

    std::string cacheDefRelPath = getPathARelativeToPathB(cacheDef.getXmlPath(), wb.getXmlPath());
    auto cacheRel = doc.workbookRelationships().addRelationship(XLRelationshipType::PivotCacheDefinition, cacheDefRelPath);

    XMLNode wbNode = wb.xmlDocument().document_element();
    XMLNode pivotCachesNode = wbNode.child("pivotCaches");
    if (pivotCachesNode.empty()) {
        pivotCachesNode = wbNode.insert_child_before("pivotCaches", wbNode.child("extLst"));
        if (pivotCachesNode.empty()) pivotCachesNode = wbNode.append_child("pivotCaches");
    }
    
    uint32_t newCacheId = 1;
    for (auto cache : pivotCachesNode.children("pivotCache")) {
        uint32_t id = cache.attribute("cacheId").as_uint();
        if (id >= newCacheId) newCacheId = id + 1;
    }
    
    XMLNode pcNode = pivotCachesNode.append_child("pivotCache");
    pcNode.append_attribute("cacheId").set_value(newCacheId);
    pcNode.append_attribute("r:id").set_value(cacheRel.id().c_str());

    std::string cacheDefRelPathFromPt = getPathARelativeToPathB(cacheDef.getXmlPath(), pivotTable.getXmlPath());
    pivotTable.relationships().addRelationship(XLRelationshipType::PivotCacheDefinition, cacheDefRelPathFromPt);

    std::string ptRelPath = getPathARelativeToPathB(pivotTable.getXmlPath(), getXmlPath());
    
    
    XMLNode ptRelPathNode = xmlDocument().document_element().child("pivotTables");
    if (ptRelPathNode.empty()) {
        XMLNode docElement = xmlDocument().document_element();
        ptRelPathNode = appendAndGetNode(docElement, "pivotTables", m_nodeOrder);
    }
    
        // Get or Create
    
    XLRelationshipItem ptRel;
    if (!relationships().targetExists(ptRelPath)) {
        ptRel = relationships().addRelationship(XLRelationshipType::PivotTable, ptRelPath);
    } else {
        ptRel = relationships().relationshipByTarget(ptRelPath);
    }


    
    XMLNode ptEntry = ptRelPathNode.append_child("pivotTableDefinition");
    ptEntry.append_attribute("r:id").set_value(ptRel.id().c_str());
    ptEntry.append_attribute("cacheId").set_value(newCacheId);




    
    XMLNode cacheDefRoot = cacheDef.xmlDocument().document_element();
    XMLNode sourceNode = cacheDefRoot.child("cacheSource").child("worksheetSource");
    
    std::string sourceSheet = "";
    std::string sourceRef = options.sourceRange;
    size_t exclaPos = options.sourceRange.find('!');
    if (exclaPos != std::string::npos) {
        sourceSheet = options.sourceRange.substr(0, exclaPos);
        sourceRef = options.sourceRange.substr(exclaPos + 1);
    }
    
    sourceNode.attribute("ref").set_value(sourceRef.c_str());
    if (!sourceSheet.empty()) {
        sourceNode.attribute("sheet").set_value(sourceSheet.c_str());
    }

    // Attempt to read headers from source range to populate cacheFields
    XMLNode cacheFieldsNode = cacheDefRoot.child("cacheFields");
    if (cacheFieldsNode.empty()) cacheFieldsNode = cacheDefRoot.append_child("cacheFields");
    cacheFieldsNode.remove_children(); // clean up template

    XMLNode ptRoot = pivotTable.xmlDocument().document_element();

    ptRoot.attribute("name").set_value(options.name.empty() ? "PivotTable1" : options.name.c_str());
    ptRoot.attribute("cacheId").set_value(newCacheId);
    
    XMLNode locNode = ptRoot.child("location");
    locNode.attribute("ref").set_value(options.targetCell.c_str());

    XMLNode pivotFieldsNode = ptRoot.child("pivotFields");
    if (pivotFieldsNode.empty()) pivotFieldsNode = ptRoot.insert_child_after("pivotFields", ptRoot.child("location"));
    pivotFieldsNode.remove_children();

    std::vector<std::string> headers;
    try {
        XLWorksheet srcWks = sourceSheet.empty() ? wb.worksheet(name()) : wb.worksheet(sourceSheet);
        XLCellReference startRef(sourceRef.substr(0, sourceRef.find(':')));
        XLCellReference endRef(sourceRef.substr(sourceRef.find(':') + 1));
        
        for (uint16_t col = startRef.column(); col <= endRef.column(); ++col) {
            XLCellValue val = srcWks.cell(startRef.row(), col).value();
            std::string headerName = val.type() == XLValueType::String ? val.get<std::string>() : "Field" + std::to_string(col);
            headers.push_back(headerName);
        }
    } catch(...) {
        // Fallback if parsing fails
        headers.push_back("Field1");
    }

    
    
    cacheFieldsNode.attribute("count").set_value(headers.size());

    // remove all old template children
    pivotFieldsNode.remove_children();
    XMLNode rowFieldsNode = ptRoot.child("rowFields");
    if (!rowFieldsNode.empty()) ptRoot.remove_child(rowFieldsNode);
    XMLNode rowItemsNode = ptRoot.child("rowItems");
    if (!rowItemsNode.empty()) ptRoot.remove_child(rowItemsNode);
    XMLNode colItemsNode = ptRoot.child("colItems");
    if (!colItemsNode.empty()) ptRoot.remove_child(colItemsNode);
    XMLNode dataFieldsNode = ptRoot.child("dataFields");
    if (!dataFieldsNode.empty()) ptRoot.remove_child(dataFieldsNode);

    // Rebuild proper
    
    // Rebuild proper pivot fields
    pivotFieldsNode.attribute("count").set_value(headers.size());

    std::vector<int> rowIndices;
    std::vector<int> colIndices;
    std::vector<int> dataIndices;
    
    auto findFieldIndex = [&](const std::string& name) -> int {
        auto it = std::find(headers.begin(), headers.end(), name);
        if (it != headers.end()) return std::distance(headers.begin(), it);
        return -1;
    };

    for (const auto& rowFld : options.rows) {
        int idx = findFieldIndex(rowFld.name);
        if (idx >= 0) rowIndices.push_back(idx);
    }
    for (const auto& colFld : options.columns) {
        int idx = findFieldIndex(colFld.name);
        if (idx >= 0) colIndices.push_back(idx);
    }
    for (const auto& dataFld : options.data) {
        int idx = findFieldIndex(dataFld.name);
        if (idx >= 0) dataIndices.push_back(idx);
    }

    int i = 0;
    for (const auto& h : headers) {
        XMLNode fieldNode = cacheFieldsNode.append_child("cacheField");
        fieldNode.append_attribute("name").set_value(h.c_str());
        fieldNode.append_attribute("sharedItems").set_value("1"); 
        fieldNode.append_child("sharedItems").append_attribute("count").set_value("0"); 

        XMLNode ptFieldNode = pivotFieldsNode.append_child("pivotField");
        
        bool isRow = std::find(rowIndices.begin(), rowIndices.end(), i) != rowIndices.end();
        bool isCol = std::find(colIndices.begin(), colIndices.end(), i) != colIndices.end();
        bool isData = std::find(dataIndices.begin(), dataIndices.end(), i) != dataIndices.end();

        if (isRow) {
            ptFieldNode.append_attribute("axis").set_value("axisRow");
            ptFieldNode.append_attribute("compact").set_value("0");
            ptFieldNode.append_attribute("outline").set_value("0");
            ptFieldNode.append_attribute("showAll").set_value("0");
            ptFieldNode.append_attribute("defaultSubtotal").set_value("1");
            XMLNode itemsNode = ptFieldNode.append_child("items");
            itemsNode.append_attribute("count").set_value("1");
            itemsNode.append_child("item").append_attribute("t").set_value("default");
        } else if (isCol) {
            ptFieldNode.append_attribute("axis").set_value("axisCol");
            ptFieldNode.append_attribute("compact").set_value("0");
            ptFieldNode.append_attribute("outline").set_value("0");
            ptFieldNode.append_attribute("showAll").set_value("0");
            ptFieldNode.append_attribute("defaultSubtotal").set_value("1");
            XMLNode itemsNode = ptFieldNode.append_child("items");
            itemsNode.append_attribute("count").set_value("1");
            itemsNode.append_child("item").append_attribute("t").set_value("default");
        } else if (isData) {
            ptFieldNode.append_attribute("dataField").set_value("1");
            ptFieldNode.append_attribute("showAll").set_value("0");
        } else {
            ptFieldNode.append_attribute("showAll").set_value("0");
        }
        i++;
    }

    if (!rowIndices.empty()) {
        rowFieldsNode = ptRoot.insert_child_after("rowFields", pivotFieldsNode);
        rowFieldsNode.append_attribute("count").set_value(rowIndices.size());
        for (int idx : rowIndices) {
            rowFieldsNode.append_child("field").append_attribute("x").set_value(idx);
        }

        rowItemsNode = ptRoot.insert_child_after("rowItems", rowFieldsNode);
        rowItemsNode.append_attribute("count").set_value("1");
        rowItemsNode.append_child("i").append_child("x");
    }

    if (!colIndices.empty()) {
        XMLNode insertAfter = rowItemsNode.empty() ? (rowFieldsNode.empty() ? pivotFieldsNode : rowFieldsNode) : rowItemsNode;
        XMLNode colFieldsNode = ptRoot.insert_child_after("colFields", insertAfter);
        colFieldsNode.append_attribute("count").set_value(colIndices.size());
        for (int idx : colIndices) {
            colFieldsNode.append_child("field").append_attribute("x").set_value(idx);
        }

        colItemsNode = ptRoot.insert_child_after("colItems", colFieldsNode);
        colItemsNode.append_attribute("count").set_value("1");
        colItemsNode.append_child("i").append_child("x");
    }

    if (!dataIndices.empty()) {
        dataFieldsNode = ptRoot.insert_child_before("dataFields", ptRoot.child("pivotTableStyleInfo"));
        dataFieldsNode.append_attribute("count").set_value(options.data.size());
        
        for (const auto& dataFld : options.data) {
            int idx = findFieldIndex(dataFld.name);
            if (idx < 0) continue;

            XMLNode dfield = dataFieldsNode.append_child("dataField");
            
            std::string dName = dataFld.customName.empty() ? ("Sum of " + dataFld.name) : dataFld.customName;
            dfield.append_attribute("name").set_value(dName.c_str());
            dfield.append_attribute("fld").set_value(idx);
            
            std::string subType = "sum";
            switch (dataFld.subtotal) {
                case XLPivotSubtotal::Average: subType = "average"; break;
                case XLPivotSubtotal::Count:   subType = "count"; break;
                case XLPivotSubtotal::Max:     subType = "max"; break;
                case XLPivotSubtotal::Min:     subType = "min"; break;
                case XLPivotSubtotal::Product: subType = "product"; break;
                default:                       subType = "sum"; break;
            }
            if (subType != "sum") { // Excel default is sum, only write if different, but let's just write
                dfield.append_attribute("subtotal").set_value(subType.c_str());
            }
        }
    }

    locNode.attribute("ref").set_value(options.targetCell.c_str());
    return pivotTable;
}

#include "XLDocument.hpp"
#include "XLUtilities.hpp"
#include "XLWorkbook.hpp"
#include "XLWorksheet.hpp"

#include "XLChart.hpp"
#include "XLComments.hpp"
#include "XLConditionalFormatting.hpp"
#include "XLDataValidation.hpp"
#include "XLDrawing.hpp"
#include "XLImageOptions.hpp"
#include "XLMergeCells.hpp"
#include "XLPageSetup.hpp"
#include "XLPivotTable.hpp"
#include "XLRelationships.hpp"
#include "XLSparkline.hpp"
#include "XLStreamReader.hpp"
#include "XLStreamWriter.hpp"
#include "XLTables.hpp"
#include "XLThreadedComments.hpp"
#include <unordered_set>
#include <algorithm>

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
                if (!targetPath.empty() && targetPath[0] != '/') {
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
                    auto prevSlash = pos > 0 ? targetPath.find_last_of('/', pos - 1) : std::string::npos;
                    if (prevSlash != std::string::npos) {
                        targetPath.erase(prevSlash + 1, pos - prevSlash + 3);
                    } else {
                        targetPath.erase(0, pos + 4);
                    }
                }
                
                if (!targetPath.empty() && targetPath[0] == '/') targetPath = targetPath.substr(1);

                XLXmlData* xmlData = const_cast<XLDocument&>(parentDoc()).getXmlData(XLInternalAccess{}, targetPath, true);
                if (xmlData == nullptr) {
                    xmlData = const_cast<XLDocument&>(parentDoc()).addXmlData(XLInternalAccess{}, targetPath, rId, XLContentType::PivotTable);
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
            if (!targetPath.empty() && targetPath[0] != '/') {
                std::string sourcePath = getXmlPath();
                auto slashPos = sourcePath.find_last_of('/');
                if (slashPos != std::string::npos) targetPath = sourcePath.substr(0, slashPos + 1) + targetPath;
            }
            
            while (targetPath.find("/../") != std::string::npos) {
                auto pos = targetPath.find("/../");
                auto prevSlash = pos > 0 ? targetPath.find_last_of('/', pos - 1) : std::string::npos;
                if (prevSlash != std::string::npos) targetPath.erase(prevSlash + 1, pos - prevSlash + 3);
                else targetPath.erase(0, pos + 4);
            }
            if (!targetPath.empty() && targetPath[0] == '/') targetPath = targetPath.substr(1);

            XLXmlData* xmlData = const_cast<XLDocument&>(parentDoc()).getXmlData(XLInternalAccess{}, targetPath, true);
            if (xmlData == nullptr) {
                xmlData = const_cast<XLDocument&>(parentDoc()).addXmlData(XLInternalAccess{}, targetPath, rId, XLContentType::PivotTable);
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
                if (!targetPath.empty() && targetPath[0] != '/') targetPath = "/xl/" + targetPath;

                while (targetPath.find("/../") != std::string::npos) {
                    auto pos = targetPath.find("/../");
                    auto prevSlash = pos > 0 ? targetPath.find_last_of('/', pos - 1) : std::string::npos;
                    if (prevSlash != std::string::npos) targetPath.erase(prevSlash + 1, pos - prevSlash + 3);
                    else targetPath.erase(0, pos + 4);
                }
                if (!targetPath.empty() && targetPath[0] == '/') targetPath = targetPath.substr(1);

                XLXmlData* xmlData = doc.getXmlData(XLInternalAccess{}, targetPath, true);
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
    setBoolAttr(ptRoot, "compactData", options.classicLayout() ? false : options.compactData());
    setBoolAttr(ptRoot, "showError", options.showError());
    setBoolAttr(ptRoot, "gridDropZones", options.classicLayout());
    setBoolAttr(ptRoot, "fieldPrintTitles", options.fieldPrintTitles());
    setBoolAttr(ptRoot, "itemPrintTitles", options.itemPrintTitles());
    
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
    std::vector<std::vector<std::string>> sharedItemStrings;
    // Track which columns are purely numeric (data fields stored inline in records)
    std::vector<bool> columnIsNumeric;

    // Fields used as row/col/filter axes must always use sharedItems (for filtering/grouping)
    std::unordered_set<std::string> axisFieldNames;
    for (const auto& f : options.rows())    axisFieldNames.insert(f.name);
    for (const auto& f : options.columns()) axisFieldNames.insert(f.name);
    for (const auto& f : options.filters()) axisFieldNames.insert(f.name);
    
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

        std::vector<std::vector<XLCellValue>> uniqueValuesPerColumn(headers.size());
        sharedItemStrings.resize(headers.size());
        columnIsNumeric.resize(headers.size(), false);

        try {
            XLWorksheet     srcWks = sourceSheet.empty() ? wb.worksheet(name()) : wb.worksheet(sourceSheet);
            XLCellReference startRef(sourceRef.substr(0, sourceRef.find(':')));
            XLCellReference endRef(sourceRef.substr(sourceRef.find(':') + 1));

            size_t colIdx = 0;
            for (uint16_t col = startRef.column(); col <= endRef.column(); ++col) {
                if (colIdx >= headers.size()) break;
                std::vector<XLCellValue> uniques;
                std::vector<std::string> uniqueStrs;
                std::unordered_set<std::string> seen;
                bool allNumeric = true;
                bool anyValue = false;
                for (uint32_t row = startRef.row() + 1; row <= endRef.row(); ++row) {
                    XLCellValue val = srcWks.cell(row, col).value();
                    std::string strVal;
                    if (val.type() == XLValueType::Empty) {
                        strVal = "";
                        allNumeric = false;
                    } else if (val.type() == XLValueType::Boolean) {
                        strVal = val.get<bool>() ? "1" : "0";
                        allNumeric = false;
                        anyValue = true;
                    } else if (val.type() == XLValueType::Integer || val.type() == XLValueType::Float) {
                        strVal = val.getString();
                        anyValue = true;
                    } else {
                        strVal = val.getString();
                        allNumeric = false;
                        anyValue = true;
                    }
                    if (seen.find(strVal) == seen.end()) {
                        seen.insert(strVal);
                        uniques.push_back(val);
                        uniqueStrs.push_back(strVal);
                    }
                }
                // A column is "numeric" (inline in records) only if ALL values are numbers
                // AND the field is NOT used as a row/col/filter axis (those need sharedItems for grouping)
                bool isAxisField = (axisFieldNames.count(headers[colIdx]) > 0);
                columnIsNumeric[colIdx] = allNumeric && anyValue && !isAxisField;
                uniqueValuesPerColumn[colIdx] = std::move(uniques);
                sharedItemStrings[colIdx] = std::move(uniqueStrs);
                colIdx++;
            }
        }
        catch (...) {
            // Fallback if parsing fails
        }

        XMLNode cacheFieldsNode = cacheDefRoot.child("cacheFields");
        if (cacheFieldsNode.empty()) cacheFieldsNode = cacheDefRoot.append_child("cacheFields");
        cacheFieldsNode.remove_children();    // clean up template
        if (!cacheFieldsNode.attribute("count")) cacheFieldsNode.append_attribute("count").set_value(headers.size());
        else cacheFieldsNode.attribute("count").set_value(headers.size());

        size_t colIdx = 0;
        for (const auto& h : headers) {
            XMLNode fieldNode = cacheFieldsNode.append_child("cacheField");
            fieldNode.append_attribute("name").set_value(h.c_str());
            fieldNode.append_attribute("numFmtId").set_value("0");

            XMLNode sharedItemsNode = fieldNode.append_child("sharedItems");
            const auto& uniques = uniqueValuesPerColumn[colIdx];

            if (uniques.empty()) {
                // No data: blank placeholder
                sharedItemsNode.append_attribute("containsBlank").set_value("1");
                sharedItemsNode.append_attribute("count").set_value("0");
                sharedItemsNode.append_child("m");
            } else if (columnIsNumeric[colIdx]) {
                // Pure numeric field (data aggregation field, not an axis/filter field).
                // Excel uses self-closing sharedItems with metadata only, no child elements.
                // This matches what Excel generates after repair.
                double minVal = std::numeric_limits<double>::max();
                double maxVal = std::numeric_limits<double>::lowest();
                bool   allInt = true;
                for (const auto& val : uniques) {
                    double d = (val.type() == XLValueType::Integer)
                                   ? static_cast<double>(val.get<int64_t>())
                                   : val.get<double>();
                    if (val.type() != XLValueType::Integer) allInt = false;
                    if (d < minVal) minVal = d;
                    if (d > maxVal) maxVal = d;
                }
                sharedItemsNode.append_attribute("containsSemiMixedTypes").set_value("0");
                sharedItemsNode.append_attribute("containsString").set_value("0");
                sharedItemsNode.append_attribute("containsNumber").set_value("1");
                if (allInt) sharedItemsNode.append_attribute("containsInteger").set_value("1");
                auto fmtNum = [](double v) -> std::string {
                    if (v == std::floor(v) && std::abs(v) < 1e15)
                        return std::to_string(static_cast<int64_t>(v));
                    return fmt::format("{}", v);
                };
                sharedItemsNode.append_attribute("minValue").set_value(fmtNum(minVal).c_str());
                sharedItemsNode.append_attribute("maxValue").set_value(fmtNum(maxVal).c_str());
                // No count attribute and no children for pure numeric fields
            } else {
                // Categorical / string / axis field — list actual distinct values
                // so pivotField item x-indices resolve correctly.
                // Excel: string fields have only count, no containsString attribute.
                bool     hasBlank  = false;
                uint32_t valCount  = 0;

                for (const auto& val : uniques) {
                    if (val.type() == XLValueType::Empty) {
                        hasBlank = true;
                        sharedItemsNode.append_child("m");
                    } else if (val.type() == XLValueType::Boolean) {
                        auto bNode = sharedItemsNode.append_child("b");
                        bNode.append_attribute("v").set_value(val.get<bool>() ? "1" : "0");
                        valCount++;
                    } else if (val.type() == XLValueType::Integer) {
                        // Numeric values in an axis/filter field (e.g. Year field)
                        auto nNode = sharedItemsNode.append_child("n");
                        nNode.append_attribute("v").set_value(std::to_string(val.get<int64_t>()).c_str());
                        valCount++;
                    } else if (val.type() == XLValueType::Float) {
                        auto nNode = sharedItemsNode.append_child("n");
                        nNode.append_attribute("v").set_value(fmt::format("{}", val.get<double>()).c_str());
                        valCount++;
                    } else {
                        auto sNode = sharedItemsNode.append_child("s");
                        sNode.append_attribute("v").set_value(val.get<std::string>().c_str());
                        valCount++;
                    }
                }
                if (hasBlank) sharedItemsNode.append_attribute("containsBlank").set_value("1");
                sharedItemsNode.append_attribute("count").set_value(valCount);
            }
            colIdx++;
        }

        // Do NOT create pivotCacheRecords — set refreshOnLoad=1 so Excel rebuilds
        // the cache when the file is opened. This matches what excelize and openpyxl do.
        // Ensure saveData="0" (template default) is preserved.
        if (cacheDefRoot.attribute("saveData")) {
            cacheDefRoot.attribute("saveData").set_value("0");
        }
    } else {
        XMLNode cacheFieldsNode = cacheDefRoot.child("cacheFields");
        for (auto cf : cacheFieldsNode.children("cacheField")) {
            headers.push_back(cf.attribute("name").value());
            
            std::vector<std::string> uniqueStrs;
            XMLNode sharedItemsNode = cf.child("sharedItems");
            if (!sharedItemsNode.empty()) {
                for (auto item : sharedItemsNode.children()) {
                    std::string tagName = item.name();
                    if (tagName == "m") {
                        uniqueStrs.push_back("");
                    } else if (tagName == "b") {
                        uniqueStrs.push_back(item.attribute("val").value());
                    } else if (tagName == "n") {
                        uniqueStrs.push_back(item.attribute("val").value());
                    } else if (tagName == "s") {
                        uniqueStrs.push_back(item.attribute("val").value());
                    }
                }
            }
            sharedItemStrings.push_back(std::move(uniqueStrs));
        }
    }

    // remove all old template children
    pivotFieldsNode.remove_children();
    XMLNode rowFieldsNode = ptRoot.child("rowFields");
    if (!rowFieldsNode.empty()) ptRoot.remove_child(rowFieldsNode);
    XMLNode rowItemsNode = ptRoot.child("rowItems");
    if (!rowItemsNode.empty()) ptRoot.remove_child(rowItemsNode);
    XMLNode colFieldsNodeOld = ptRoot.child("colFields");
    if (!colFieldsNodeOld.empty()) ptRoot.remove_child(colFieldsNodeOld);
    XMLNode colItemsNode = ptRoot.child("colItems");
    if (!colItemsNode.empty()) ptRoot.remove_child(colItemsNode);
    XMLNode pageFieldsNodeOld = ptRoot.child("pageFields");
    if (!pageFieldsNodeOld.empty()) ptRoot.remove_child(pageFieldsNodeOld);
    XMLNode dataFieldsNode = ptRoot.child("dataFields");
    if (!dataFieldsNode.empty()) ptRoot.remove_child(dataFieldsNode);

    // Remove any stray whitespace text nodes left by template cleanup
    {
        std::vector<XMLNode> toRemove;
        for (auto node = ptRoot.first_child(); node; node = node.next_sibling()) {
            if (node.type() == pugi::node_pcdata || node.type() == pugi::node_cdata) {
                std::string txt = node.value();
                bool allSpace = true;
                for (char c : txt) { if (!std::isspace(static_cast<unsigned char>(c))) { allSpace = false; break; } }
                if (allSpace) toRemove.push_back(node);
            }
        }
        for (auto& n : toRemove) ptRoot.remove_child(n);
    }

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

    auto findRowField = [&](const std::string& name) -> const XLPivotField* {
        for (const auto& rf : options.rows()) {
            if (rf.name == name) return &rf;
        }
        return nullptr;
    };
    auto findColField = [&](const std::string& name) -> const XLPivotField* {
        for (const auto& cf : options.columns()) {
            if (cf.name == name) return &cf;
        }
        return nullptr;
    };
    auto findFilterField = [&](const std::string& name) -> const XLPivotField* {
        for (const auto& ff : options.filters()) {
            if (ff.name == name) return &ff;
        }
        return nullptr;
    };

    auto buildItemsNode = [&](XMLNode ptFieldNode, int fieldIdx, const std::vector<std::string>& selectedItems) {
        XMLNode itemsNode = ptFieldNode.append_child("items");
        if (selectedItems.empty() || fieldIdx >= static_cast<int>(sharedItemStrings.size())) {
            itemsNode.append_attribute("count").set_value("1");
            itemsNode.append_child("item").append_attribute("t").set_value("default");
        } else {
            const auto& uniques = sharedItemStrings[fieldIdx];
            uint32_t itemCount = 0;
            for (size_t x = 0; x < uniques.size(); ++x) {
                bool isSelected = (std::find(selectedItems.begin(), selectedItems.end(), uniques[x]) != selectedItems.end());
                XMLNode itemNode = itemsNode.append_child("item");
                itemNode.append_attribute("x").set_value(x);
                if (!isSelected) {
                    itemNode.append_attribute("h").set_value("1");
                }
                itemCount++;
            }
            itemsNode.append_child("item").append_attribute("t").set_value("default");
            itemCount++;
            itemsNode.append_attribute("count").set_value(itemCount);
        }
    };

    int i = 0;
    for (const auto& h : headers) {
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
            
            const XLPivotField* rf = findRowField(h);
            buildItemsNode(ptFieldNode, i, rf ? rf->selectedItems : std::vector<std::string>{});
        }
        else if (isCol) {
            ptFieldNode.append_attribute("axis").set_value("axisCol");
            ptFieldNode.append_attribute("compact").set_value("0");
            ptFieldNode.append_attribute("outline").set_value("0");
            ptFieldNode.append_attribute("showAll").set_value("0");
            ptFieldNode.append_attribute("defaultSubtotal").set_value("1");
            
            const XLPivotField* cf = findColField(h);
            buildItemsNode(ptFieldNode, i, cf ? cf->selectedItems : std::vector<std::string>{});
        }
        else if (isFilter) {
            ptFieldNode.append_attribute("axis").set_value("axisPage");
            ptFieldNode.append_attribute("compact").set_value("0");
            ptFieldNode.append_attribute("outline").set_value("0");
            ptFieldNode.append_attribute("showAll").set_value("0");
            ptFieldNode.append_attribute("defaultSubtotal").set_value("1");
            
            const XLPivotField* ff = findFilterField(h);
            buildItemsNode(ptFieldNode, i, ff ? ff->selectedItems : std::vector<std::string>{});
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
        size_t rowFieldsCount = rowIndices.size() + (hasVirtualDataField && virtualDataOnRows ? 1 : 0);
        rowFieldsNode.append_attribute("count").set_value(rowFieldsCount);
        for (int idx : rowIndices) { rowFieldsNode.append_child("field").append_attribute("x").set_value(idx); }
        if (hasVirtualDataField && virtualDataOnRows) {
            rowFieldsNode.append_child("field").append_attribute("x").set_value("-2"); // -2 represents the Values field in Excel
        }

        rowItemsNode = ptRoot.insert_child_after("rowItems", rowFieldsNode);
        rowItemsNode.append_attribute("count").set_value("1");
        // Minimal placeholder row item — Excel rebuilds on refresh (refreshOnLoad=1)
        XMLNode rowItem = rowItemsNode.append_child("i");
        rowItem.append_child("x");
    }

    if (!colIndices.empty() || (hasVirtualDataField && !virtualDataOnRows)) {
        XMLNode insertAfter   = rowItemsNode.empty() ? (rowFieldsNode.empty() ? pivotFieldsNode : rowFieldsNode) : rowItemsNode;
        XMLNode colFieldsNode = ptRoot.insert_child_after("colFields", insertAfter);
        size_t colFieldsCount = colIndices.size() + (hasVirtualDataField && !virtualDataOnRows ? 1 : 0);
        colFieldsNode.append_attribute("count").set_value(colFieldsCount);
        for (int idx : colIndices) { colFieldsNode.append_child("field").append_attribute("x").set_value(idx); }
        if (hasVirtualDataField && !virtualDataOnRows) {
            colFieldsNode.append_child("field").append_attribute("x").set_value("-2"); // -2 represents the Values field in Excel
        }

        colItemsNode = ptRoot.insert_child_after("colItems", colFieldsNode);
        colItemsNode.append_attribute("count").set_value("1");
        // Minimal placeholder col item — Excel rebuilds on refresh (refreshOnLoad=1)
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
            pageField.append_attribute("hier").set_value("0");  // required by Excel
        }
    }
    if (!dataIndices.empty()) {
        dataFieldsNode = ptRoot.insert_child_before("dataFields", ptRoot.child("pivotTableStyleInfo"));
        dataFieldsNode.append_attribute("count").set_value(options.data().size());

        for (const auto& dataFld : options.data()) {
            int idx = findFieldIndex(dataFld.name);
            if (idx < 0) continue;

            XMLNode dfield = dataFieldsNode.append_child("dataField");

            std::string prefix = "Sum of ";
            std::string subType = "sum";
            switch (dataFld.subtotal) {
                case XLPivotSubtotal::Average:
                    subType = "average";
                    prefix = "Average of ";
                    break;
                case XLPivotSubtotal::Count:
                    subType = "count";
                    prefix = "Count of ";
                    break;
                case XLPivotSubtotal::Max:
                    subType = "max";
                    prefix = "Max of ";
                    break;
                case XLPivotSubtotal::Min:
                    subType = "min";
                    prefix = "Min of ";
                    break;
                case XLPivotSubtotal::Product:
                    subType = "product";
                    prefix = "Product of ";
                    break;
                case XLPivotSubtotal::CountNums:
                    subType = "countNums";
                    prefix = "Count Nums of ";
                    break;
                case XLPivotSubtotal::StdDev:
                    subType = "stdDev";
                    prefix = "StdDev of ";
                    break;
                case XLPivotSubtotal::StdDevP:
                    subType = "stdDevp";
                    prefix = "StdDevP of ";
                    break;
                case XLPivotSubtotal::Var:
                    subType = "var";
                    prefix = "Var of ";
                    break;
                case XLPivotSubtotal::VarP:
                    subType = "varp";
                    prefix = "VarP of ";
                    break;
                default:
                    subType = "sum";
                    prefix = "Sum of ";
                    break;
            }

            std::string dName = dataFld.customName.empty() ? (prefix + dataFld.name) : dataFld.customName;
            dfield.append_attribute("name").set_value(dName.c_str());
            dfield.append_attribute("fld").set_value(idx);

            // baseField/baseItem: 0/0 matches what Excel writes for normal sum/aggregate.
            // The magic value -1/1048832 used previously is rejected by Excel's validator.
            if (dataFld.subtotal != XLPivotSubtotal::Sum) {
                dfield.append_attribute("subtotal").set_value(subType.c_str());
            }
            dfield.append_attribute("baseField").set_value("0");
            dfield.append_attribute("baseItem").set_value("0");

            if (dataFld.numFmtId != 0) {
                dfield.append_attribute("numFmtId").set_value(dataFld.numFmtId);
            }
        }
    }

    // Calculate the dimensions of the PivotTable
    uint32_t numFilters = filterIndices.size();
    
    uint32_t rowItemsCount = 0;
    for (int idx : rowIndices) {
        const auto& rf = findRowField(headers[idx]);
        if (rf && !rf->selectedItems.empty()) {
            rowItemsCount += rf->selectedItems.size();
        } else if (static_cast<size_t>(idx) < sharedItemStrings.size()) {
            rowItemsCount += sharedItemStrings[idx].size();
        }
    }
    if (rowItemsCount == 0) rowItemsCount = 1;

    uint32_t colItemsCount = 0;
    for (int idx : colIndices) {
        const auto& cf = findColField(headers[idx]);
        if (cf && !cf->selectedItems.empty()) {
            colItemsCount += cf->selectedItems.size();
        } else if (static_cast<size_t>(idx) < sharedItemStrings.size()) {
            colItemsCount += sharedItemStrings[idx].size();
        }
    }
    if (colItemsCount == 0) colItemsCount = 1;

    uint32_t dataFieldsCount = dataIndices.size();
    uint32_t colMultiplier = (hasVirtualDataField && !virtualDataOnRows) ? dataFieldsCount : 1;
    uint32_t rowMultiplier = (hasVirtualDataField && virtualDataOnRows) ? dataFieldsCount : 1;

    uint32_t totalDataCols = colItemsCount * colMultiplier;
    if (options.colGrandTotals()) totalDataCols += colMultiplier;

    uint32_t totalDataRows = rowItemsCount * rowMultiplier;
    if (options.rowGrandTotals()) totalDataRows += rowMultiplier;

    // Header height (H) and Row labels width (R_cols)
    uint32_t H = colIndices.size() + ((hasVirtualDataField && !virtualDataOnRows) ? 1 : 0);
    if (H == 0) H = 1;

    uint32_t R_cols = rowIndices.size() + ((hasVirtualDataField && virtualDataOnRows) ? 1 : 0);
    if (R_cols == 0) R_cols = 1;

    uint32_t firstHeaderRow = 1;
    uint32_t firstDataRow = H + 1;
    uint32_t firstDataCol = R_cols + 1;

    uint32_t gridWidth = R_cols + totalDataCols;
    uint32_t gridHeight = H + totalDataRows;

    XLCellReference targetRef(options.targetCell());
    uint32_t startRow = targetRef.row();
    uint16_t startCol = targetRef.column();

    // The grid area starts after filter fields plus a blank spacer row
    uint32_t gridStartRow = startRow;
    if (numFilters > 0) {
        gridStartRow += numFilters + 1;
    }

    XLCellReference gridStartRef(gridStartRow, startCol);
    XLCellReference gridEndRef(gridStartRow + gridHeight - 1, startCol + gridWidth - 1);
    std::string ptRange = gridStartRef.address() + ":" + gridEndRef.address();

    locNode.attribute("ref").set_value(ptRange.c_str());
    
    // Update or append location attributes
    auto updateOrAppend = [](XMLNode node, const char* attrName, uint32_t val) {
        if (!node.attribute(attrName)) {
            node.append_attribute(attrName).set_value(val);
        } else {
            node.attribute(attrName).set_value(val);
        }
    };
    
    updateOrAppend(locNode, "firstHeaderRow", firstHeaderRow);
    updateOrAppend(locNode, "firstDataRow", firstDataRow);
    updateOrAppend(locNode, "firstDataCol", firstDataCol);

    if (numFilters > 0) {
        updateOrAppend(locNode, "rowPageCount", numFilters);
        updateOrAppend(locNode, "colPageCount", 1);
    } else {
        if (locNode.attribute("rowPageCount")) locNode.remove_attribute("rowPageCount");
        if (locNode.attribute("colPageCount")) locNode.remove_attribute("colPageCount");
    }

    return pivotTable;
}

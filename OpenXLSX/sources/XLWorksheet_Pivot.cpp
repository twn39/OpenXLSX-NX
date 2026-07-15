#include "XLDocument.hpp"
#include "XLUtilities.hpp"
#include "XLWorkbook.hpp"
#include "XLWorksheet.hpp"
#include <pugixml.hpp>

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
#include "XLException.hpp"
#include <algorithm>
#include <cctype>
#include <unordered_set>

using namespace OpenXLSX;

namespace {

std::string stripLeadingSlash(std::string p)
{
    if (!p.empty() && p.front() == '/') p.erase(0, 1);
    return p;
}

std::string unquoteSheetName(std::string sheet)
{
    if (sheet.size() >= 2 && sheet.front() == '\'' && sheet.back() == '\'')
        return sheet.substr(1, sheet.size() - 2);
    // Excel escapes doubled single-quotes inside quoted sheet names
    std::string out;
    out.reserve(sheet.size());
    for (size_t i = 0; i < sheet.size(); ++i) {
        if (sheet[i] == '\'' && i + 1 < sheet.size() && sheet[i + 1] == '\'') {
            out.push_back('\'');
            ++i;
        }
        else {
            out.push_back(sheet[i]);
        }
    }
    return out;
}

std::string normalizeRangeKey(std::string range)
{
    auto bang = range.find('!');
    if (bang == std::string::npos) return range;
    std::string sheet = unquoteSheetName(range.substr(0, bang));
    std::string ref   = range.substr(bang + 1);
    // drop $
    ref.erase(std::remove(ref.begin(), ref.end(), '$'), ref.end());
    return sheet + "!" + ref;
}

/** Resolve a relationship Target relative to a package part path into an absolute package path. */
std::string resolveRelTarget(std::string_view sourcePartPath, std::string target)
{
    if (target.empty()) return {};
    if (target.front() == '/') return stripLeadingSlash(std::string(target));

    std::string source(sourcePartPath);
    auto slash = source.find_last_of('/');
    std::string baseDir = (slash == std::string::npos) ? std::string() : source.substr(0, slash + 1);
    std::string combined = baseDir + target;
    std::string abs = eliminateDotAndDotDotFromPath(combined);
    return stripLeadingSlash(abs);
}

struct PivotSourceResolved {
    std::string identityKey;   ///< cache reuse key
    std::string sheet;         ///< worksheet name for header/data scan
    std::string ref;           ///< A1:D10 style (no sheet)
    std::string nameAttr;      ///< worksheetSource/@name when table or defined name
    bool        namedSource{false};
};

bool looksLikeSheetRange(std::string_view s)
{
    return s.find('!') != std::string_view::npos;
}

bool parseSheetAndRef(std::string_view formula, std::string& sheet, std::string& ref)
{
    std::string f(formula);
    if (!f.empty() && f.front() == '=') f.erase(0, 1);
    // trim
    while (!f.empty() && std::isspace(static_cast<unsigned char>(f.front()))) f.erase(0, 1);
    while (!f.empty() && std::isspace(static_cast<unsigned char>(f.back()))) f.pop_back();

    auto bang = f.find('!');
    if (bang == std::string::npos) return false;
    sheet = unquoteSheetName(f.substr(0, bang));
    ref   = f.substr(bang + 1);
    ref.erase(std::remove(ref.begin(), ref.end(), '$'), ref.end());
    return !sheet.empty() && ref.find(':') != std::string::npos;
}

PivotSourceResolved resolvePivotSource(XLWorksheet& self, const std::string& source)
{
    PivotSourceResolved out;
    auto& doc = self.parentDoc();
    auto  wb  = doc.workbook();

    if (source.empty())
        throw XLInputError("XLWorksheet::addPivotTable: sourceRange is empty");

    if (looksLikeSheetRange(source)) {
        std::string sheet, ref;
        if (!parseSheetAndRef(source, sheet, ref))
            throw XLInputError("XLWorksheet::addPivotTable: invalid source range \"" + source + "\"");
        out.sheet        = sheet;
        out.ref          = ref;
        out.namedSource  = false;
        out.identityKey  = "range:" + normalizeRangeKey(sheet + "!" + ref);
        return out;
    }

    // Excel Table name (search all worksheets)
    for (const auto& sheetName : wb.worksheetNames()) {
        XLWorksheet wks = wb.worksheet(sheetName);
        auto&       tc  = wks.tables();
        for (size_t i = 0; i < tc.count(); ++i) {
            XLTable t = tc[i];
            if (t.name() == source || t.displayName() == source) {
                out.sheet       = sheetName;
                out.ref         = t.rangeReference();
                out.nameAttr    = std::string(source);
                // Prefer canonical table name for cache identity
                out.nameAttr    = t.name();
                out.namedSource = true;
                out.identityKey = "table:" + t.name();
                if (out.ref.find(':') == std::string::npos)
                    throw XLInputError("XLWorksheet::addPivotTable: table \"" + t.name() + "\" has invalid range");
                return out;
            }
        }
    }

    // Defined name
    if (wb.definedNames().exists(source)) {
        auto dn = wb.definedNames().get(source);
        std::string sheet, ref;
        if (!parseSheetAndRef(dn.refersTo(), sheet, ref))
            throw XLInputError("XLWorksheet::addPivotTable: defined name \"" + source +
                               "\" does not resolve to a worksheet range");
        out.sheet       = sheet;
        out.ref         = ref;
        out.nameAttr    = std::string(source);
        out.namedSource = true;
        out.identityKey = "name:" + std::string(source);
        return out;
    }

    throw XLInputError(
        "XLWorksheet::addPivotTable: source \"" + source +
        "\" is not a Sheet!A1:B2 range, Excel Table name, or defined name");
}

std::string relsPathForPart(std::string_view partPath)
{
    std::string p = stripLeadingSlash(std::string(partPath));
    auto        slash = p.find_last_of('/');
    if (slash == std::string::npos) return "_rels/" + p + ".rels";
    return p.substr(0, slash) + "/_rels/" + p.substr(slash + 1) + ".rels";
}

std::string cachePathFromPivotTable(XLDocument& doc, const XLPivotTable& pt)
{
    const std::string ptPath = pt.getXmlPath();
    // Prefer workbook pivotCaches by cacheId (stable)
    try {
        return pt.cacheDefinition().getXmlPath();
    }
    catch (...) {
    }

    // Fallback: pivot table relationships (xmlRelationships may create empty .rels if missing).
    auto ptRels = doc.xmlRelationships(ptPath);
    for (const auto& rel : ptRels.relationships()) {
        if (rel.type() == XLRelationshipType::PivotCacheDefinition) {
            return resolveRelTarget(ptPath, rel.target());
        }
    }
    return {};
}

int countPivotsUsingCache(XLDocument& doc, const std::string& cachePath)
{
    const std::string want = stripLeadingSlash(cachePath);
    int               n    = 0;
    // Use relationship-first public discovery on each worksheet (matches ECMA-376).
    for (const auto& sheetName : doc.workbook().worksheetNames()) {
        XLWorksheet wks = doc.workbook().worksheet(sheetName);
        for (auto& pt : wks.pivotTables()) {
            try {
                if (stripLeadingSlash(pt.cacheDefinition().getXmlPath()) == want) ++n;
            }
            catch (...) {
            }
        }
    }
    return n;
}

void removeWorksheetPivotIndexEntry(XLWorksheet& wks, std::string_view rId)
{
    XMLNode ptNode = wks.xmlDocument().document_element().child("pivotTables");
    if (ptNode.empty()) return;
    for (auto pt = ptNode.child("pivotTable"); pt;) {
        auto next = pt.next_sibling("pivotTable");
        if (std::string(pt.attribute("r:id").value()) == rId) {
            ptNode.remove_child(pt);
        }
        pt = next;
    }
    if (ptNode.first_child().empty()) {
        wks.xmlDocument().document_element().remove_child(ptNode);
    }
}

void deleteOrphanPivotCache(XLDocument& doc, const std::string& cachePath, uint32_t cacheId)
{
    if (cachePath.empty()) return;

    // Remove workbook pivotCaches entry + relationship
    auto      wb      = doc.workbook();
    XMLNode   wbRoot  = wb.xmlDocument().document_element();
    XMLNode   caches  = wbRoot.child("pivotCaches");
    std::string rId;
    if (!caches.empty()) {
        for (auto pc = caches.child("pivotCache"); pc;) {
            auto next = pc.next_sibling("pivotCache");
            if (pc.attribute("cacheId").as_uint() == cacheId) {
                rId = pc.attribute("r:id").value();
                caches.remove_child(pc);
            }
            pc = next;
        }
        // Prefer element-count: leftover whitespace text nodes must not keep an empty pivotCaches.
        if (caches.child("pivotCache").empty()) wbRoot.remove_child(caches);
    }
    if (!rId.empty()) {
        try {
            doc.workbookRelationships().deleteRelationship(rId);
        }
        catch (...) {
        }
    }

    // Delete cache records if linked (only when a .rels part already exists).
    const std::string cacheRels = relsPathForPart(cachePath);
    if (doc.archive().hasEntry(cacheRels)) {
        auto crels = doc.xmlRelationships(cachePath);
        for (const auto& rel : crels.relationships()) {
            if (rel.type() == XLRelationshipType::PivotCacheRecords) {
                doc.deleteManagedXmlPart(resolveRelTarget(cachePath, rel.target()));
            }
        }
    }
    doc.deleteManagedXmlPart(cachePath);
}

}    // namespace

std::vector<XLPivotTable> XLWorksheet::pivotTables()
{
    std::vector<XLPivotTable> result;
    std::unordered_set<std::string> seenPaths;

    auto loadPivotAt = [&](const std::string& targetPath, const std::string& rId) {
        if (targetPath.empty() || !seenPaths.insert(targetPath).second) return;
        XLXmlData* xmlData = parentDoc().getXmlData(XLInternalAccess{}, targetPath, true);
        if (xmlData == nullptr) {
            // Load from archive if present
            if (parentDoc().archive().hasEntry(targetPath)) {
                xmlData = parentDoc().addXmlData(XLInternalAccess{}, targetPath, rId, XLContentType::PivotTable);
            }
            else {
                return;
            }
        }
        result.emplace_back(xmlData);
    };

    // Canonical discovery: worksheet package relationships (ECMA-376).
    for (const auto& rel : relationships().relationships()) {
        if (rel.type() != XLRelationshipType::PivotTable) continue;
        std::string targetPath = resolveRelTarget(getXmlPath(), rel.target());
        loadPivotAt(targetPath, rel.id());
    }

    // Legacy dual-index: worksheet <pivotTables> (OpenXLSX-written files).
    XMLNode ptNode = xmlDocument().document_element().child("pivotTables");
    if (!ptNode.empty()) {
        for (auto pt : ptNode.children("pivotTable")) {
            std::string rId = pt.attribute("r:id").value();
            if (rId.empty()) continue;
            try {
                std::string targetPath = resolveRelTarget(getXmlPath(), relationships().relationshipById(rId).target());
                loadPivotAt(targetPath, rId);
            }
            catch (...) {
            }
        }
    }

    return result;
}

bool XLWorksheet::deletePivotTable(std::string_view name)
{
    auto& doc = parentDoc();

    std::string targetRId;
    std::string ptPath;
    XLXmlData*  ptXml = nullptr;

    for (const auto& rel : relationships().relationships()) {
        if (rel.type() != XLRelationshipType::PivotTable) continue;
        std::string path = resolveRelTarget(getXmlPath(), rel.target());
        XLXmlData*  xml  = doc.getXmlData(XLInternalAccess{}, path, true);
        if (xml == nullptr && doc.archive().hasEntry(path)) {
            xml = doc.addXmlData(XLInternalAccess{}, path, rel.id(), XLContentType::PivotTable);
        }
        if (!xml) continue;
        XLPivotTable ptObj(xml);
        if (ptObj.name() == name) {
            targetRId = rel.id();
            ptPath    = path;
            ptXml     = xml;
            break;
        }
    }

    // Fallback: legacy index only
    if (!ptXml) {
        XMLNode ptNode = xmlDocument().document_element().child("pivotTables");
        for (auto pt : ptNode.children("pivotTable")) {
            std::string rId = pt.attribute("r:id").value();
            if (rId.empty()) continue;
            try {
                std::string path = resolveRelTarget(getXmlPath(), relationships().relationshipById(rId).target());
                XLXmlData*  xml  = doc.getXmlData(XLInternalAccess{}, path, true);
                if (!xml && doc.archive().hasEntry(path))
                    xml = doc.addXmlData(XLInternalAccess{}, path, rId, XLContentType::PivotTable);
                if (!xml) continue;
                XLPivotTable ptObj(xml);
                if (ptObj.name() == name) {
                    targetRId = rId;
                    ptPath    = path;
                    ptXml     = xml;
                    break;
                }
            }
            catch (...) {
            }
        }
    }

    if (!ptXml || targetRId.empty()) return false;

    XLPivotTable ptObj(ptXml);
    uint32_t     cacheId = ptObj.xmlDocument().document_element().attribute("cacheId").as_uint(0);
    std::string  cachePath;
    try {
        cachePath = ptObj.cacheDefinition().getXmlPath();
    }
    catch (...) {
        cachePath = cachePathFromPivotTable(doc, ptObj);
    }

    // Detach sheet relationship + optional non-standard index node.
    relationships().deleteRelationship(targetRId);
    removeWorksheetPivotIndexEntry(*this, targetRId);

    // Drop pivot table part (+ companion .rels) from package.
    doc.deleteManagedXmlPart(ptPath);

    // Orphan cache GC (reference count across workbook).
    if (!cachePath.empty() && cacheId != 0) {
        const int remaining = countPivotsUsingCache(doc, cachePath);
        if (remaining == 0) {
            deleteOrphanPivotCache(doc, cachePath, cacheId);
        }
    }

    return true;
}

XLPivotTable XLWorksheet::addPivotTable(const XLPivotTableOptions& options)
{
    auto& doc = parentDoc();
    auto  wb  = parentDoc().workbook();

    if (options.name().size() > 255)
        throw XLInputError("XLWorksheet::addPivotTable: pivot table name exceeds 255 characters");
    if (options.targetCell().empty())
        throw XLInputError("XLWorksheet::addPivotTable: targetCell is empty");

    // Validate axis field mutual exclusion (excelize: same field cannot be in rows/cols/filter together).
    {
        std::unordered_set<std::string> used;
        auto checkAxis = [&](const std::vector<XLPivotField>& fields, const char* axis) {
            for (const auto& f : fields) {
                if (f.name.empty())
                    throw XLInputError(std::string("XLWorksheet::addPivotTable: empty field name on ") + axis);
                if (!used.insert(f.name).second)
                    throw XLInputError("XLWorksheet::addPivotTable: field \"" + f.name +
                                       "\" cannot appear in more than one of rows/columns/filters");
            }
        };
        checkAxis(options.rows(), "rows");
        checkAxis(options.columns(), "columns");
        checkAxis(options.filters(), "filters");
    }

    const PivotSourceResolved src = resolvePivotSource(*this, options.sourceRange());

    XMLNode wbNode          = wb.xmlDocument().document_element();
    XMLNode pivotCachesNode = wbNode.child("pivotCaches");

    uint32_t newCacheId = 0;
    XLPivotCacheDefinition cacheDef(nullptr);

    // Reuse pivot cache when the resolved identity matches (range / table / defined name).
    if (!pivotCachesNode.empty()) {
        for (auto cache : pivotCachesNode.children("pivotCache")) {
            std::string rId = cache.attribute("r:id").value();
            if (rId.empty()) continue;
            std::string targetPath = doc.workbookRelationships().relationshipById(rId).target();
            if (!targetPath.empty() && targetPath[0] != '/') targetPath = "/xl/" + targetPath;
            targetPath = stripLeadingSlash(eliminateDotAndDotDotFromPath(targetPath));

            XLXmlData* xmlData = doc.getXmlData(XLInternalAccess{}, targetPath, true);
            if (xmlData == nullptr) continue;
            XLPivotCacheDefinition existingCache(xmlData);
            XMLNode                wsSrc = existingCache.xmlDocument().document_element().child("cacheSource").child("worksheetSource");
            if (wsSrc.empty()) continue;

            std::string existingKey;
            std::string nameAttr = wsSrc.attribute("name").value();
            if (!nameAttr.empty()) {
                existingKey = "table:" + nameAttr;    // also covers defined-name identity stored as @name
                // Prefer exact identityKey match for table: vs name:
                if (src.identityKey == "table:" + nameAttr || src.identityKey == "name:" + nameAttr ||
                    src.nameAttr == nameAttr) {
                    newCacheId = cache.attribute("cacheId").as_uint();
                    cacheDef   = existingCache;
                    break;
                }
            }
            else {
                std::string sheet = wsSrc.attribute("sheet").value();
                std::string ref   = wsSrc.attribute("ref").value();
                existingKey       = "range:" + normalizeRangeKey(sheet + "!" + ref);
                if (existingKey == src.identityKey) {
                    newCacheId = cache.attribute("cacheId").as_uint();
                    cacheDef   = existingCache;
                    break;
                }
            }
            (void)existingKey;
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

    // Optional dual-index for older OpenXLSX readers; discovery is relationship-first.
    XMLNode ptRelPathNode = xmlDocument().document_element().child("pivotTables");
    if (ptRelPathNode.empty()) {
        XMLNode docElement = xmlDocument().document_element();
        ptRelPathNode      = ensureChild(docElement, "pivotTables", m_nodeOrder);
    }

    XLRelationshipItem ptRel;
    if (!relationships().targetExists(ptRelPath)) { ptRel = relationships().addRelationship(XLRelationshipType::PivotTable, ptRelPath); }
    else {
        ptRel = relationships().relationshipByTarget(ptRelPath);
    }

    XMLNode ptEntry = ptRelPathNode.append_child("pivotTable");
    ptEntry.append_attribute("r:id").set_value(ptRel.id().c_str());

    XMLNode cacheDefRoot = cacheDef.xmlDocument().document_element();

    const std::string& sourceSheet = src.sheet;
    const std::string& sourceRef   = src.ref;

    if (isNewCache) {
        XMLNode sourceNode = cacheDefRoot.child("cacheSource").child("worksheetSource");
        if (src.namedSource) {
            if (sourceNode.attribute("name")) sourceNode.attribute("name").set_value(src.nameAttr.c_str());
            else sourceNode.append_attribute("name").set_value(src.nameAttr.c_str());
            // Keep sheet+ref when known so headers can be rebuilt without resolving names again.
            if (sourceNode.attribute("ref")) sourceNode.attribute("ref").set_value(sourceRef.c_str());
            else sourceNode.append_attribute("ref").set_value(sourceRef.c_str());
            if (!sourceSheet.empty()) {
                if (sourceNode.attribute("sheet")) sourceNode.attribute("sheet").set_value(sourceSheet.c_str());
                else sourceNode.append_attribute("sheet").set_value(sourceSheet.c_str());
            }
        }
        else {
            if (sourceNode.attribute("name")) sourceNode.remove_attribute("name");
            if (sourceNode.attribute("ref")) sourceNode.attribute("ref").set_value(sourceRef.c_str());
            else sourceNode.append_attribute("ref").set_value(sourceRef.c_str());
            if (!sourceSheet.empty()) {
                if (sourceNode.attribute("sheet")) sourceNode.attribute("sheet").set_value(sourceSheet.c_str());
                else sourceNode.append_attribute("sheet").set_value(sourceSheet.c_str());
            }
        }
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
    for (const auto& f : options.rows()) axisFieldNames.insert(f.name);
    for (const auto& f : options.columns()) axisFieldNames.insert(f.name);
    for (const auto& f : options.filters()) axisFieldNames.insert(f.name);

    // Deep shared-item lists only when needed (excelize-style on-demand scan).
    std::unordered_set<std::string> deepSharedFields;
    for (const auto& f : options.filters()) deepSharedFields.insert(f.name);
    for (const auto& f : options.rows())
        if (!f.selectedItems.empty()) deepSharedFields.insert(f.name);
    for (const auto& f : options.columns())
        if (!f.selectedItems.empty()) deepSharedFields.insert(f.name);
    for (const auto& f : options.data()) {
        const auto t = f.showValuesAs.type;
        if (t == XLPivotShowValuesAs::Normal) continue;
        if (pivotShowValuesAsNeedsBaseField(t)) {
            if (f.showValuesAs.baseField.empty())
                throw XLInputError("XLWorksheet::addPivotTable: ShowValuesAs requires baseField for field \"" +
                                   f.name + "\"");
            deepSharedFields.insert(f.showValuesAs.baseField);
        }
        if (pivotShowValuesAsNeedsBaseItem(t) && f.showValuesAs.baseItem.empty())
            throw XLInputError("XLWorksheet::addPivotTable: ShowValuesAs requires baseItem for field \"" + f.name +
                               "\"");
    }
    // Filters always need deep lists for dropdowns; axis fields without selection use placeholders.

    if (isNewCache) {
        if (sourceSheet.empty() || sourceRef.find(':') == std::string::npos) {
            throw XLInputError("XLWorksheet::addPivotTable: resolved source range is invalid");
        }
        XLWorksheet     srcWks = wb.worksheet(sourceSheet);
        XLCellReference startRef(sourceRef.substr(0, sourceRef.find(':')));
        XLCellReference endRef(sourceRef.substr(sourceRef.find(':') + 1));
        if (endRef.row() < startRef.row() || endRef.column() < startRef.column()) {
            throw XLInputError("XLWorksheet::addPivotTable: source range is inverted or empty");
        }

        for (uint16_t col = startRef.column(); col <= endRef.column(); ++col) {
            XLCellValue val = srcWks.cell(startRef.row(), col).value();
            std::string headerName;
            if (val.type() == XLValueType::String) headerName = val.get<std::string>();
            else if (val.type() == XLValueType::Empty)
                throw XLInputError("XLWorksheet::addPivotTable: header cell is empty in column " + std::to_string(col));
            else
                headerName = val.getString();
            if (headerName.empty())
                throw XLInputError("XLWorksheet::addPivotTable: empty header name in source range");
            headers.push_back(std::move(headerName));
        }
        if (headers.empty())
            throw XLInputError("XLWorksheet::addPivotTable: source has no header columns");

        std::vector<std::vector<XLCellValue>> uniqueValuesPerColumn(headers.size());
        sharedItemStrings.resize(headers.size());
        columnIsNumeric.resize(headers.size(), false);

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
            // Keep distinct value lists only when deep shared items are required.
            const bool needDeep = deepSharedFields.count(headers[colIdx]) > 0;
            if (needDeep) {
                uniqueValuesPerColumn[colIdx] = std::move(uniques);
                sharedItemStrings[colIdx]     = std::move(uniqueStrs);
            }
            else if (columnIsNumeric[colIdx]) {
                // Pure numeric data fields: keep uniques for min/max metadata only.
                uniqueValuesPerColumn[colIdx] = std::move(uniques);
            }
            else {
                uniqueValuesPerColumn[colIdx].clear();
                sharedItemStrings[colIdx].clear();
            }
            (void)isAxisField;
            colIdx++;
        }

        XMLNode cacheFieldsNode = cacheDefRoot.child("cacheFields");
        if (cacheFieldsNode.empty()) cacheFieldsNode = cacheDefRoot.append_child("cacheFields");
        cacheFieldsNode.remove_children();    // clean up template
        if (!cacheFieldsNode.attribute("count")) cacheFieldsNode.append_attribute("count").set_value(headers.size());
        else cacheFieldsNode.attribute("count").set_value(headers.size());

        size_t fieldColIdx = 0;
        for (const auto& h : headers) {
            XMLNode fieldNode = cacheFieldsNode.append_child("cacheField");
            fieldNode.append_attribute("name").set_value(h.c_str());
            fieldNode.append_attribute("numFmtId").set_value("0");

            XMLNode sharedItemsNode = fieldNode.append_child("sharedItems");
            const auto& uniques  = uniqueValuesPerColumn[fieldColIdx];
            const bool  needDeep = deepSharedFields.count(h) > 0;

            if (!needDeep && !columnIsNumeric[fieldColIdx]) {
                // excelize-style lightweight placeholder
                sharedItemsNode.append_attribute("containsBlank").set_value("1");
                sharedItemsNode.append_child("m");
            } else if (uniques.empty()) {
                // No data: blank placeholder
                sharedItemsNode.append_attribute("containsBlank").set_value("1");
                sharedItemsNode.append_attribute("count").set_value("0");
                sharedItemsNode.append_child("m");
            } else if (columnIsNumeric[fieldColIdx] && !needDeep) {
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
                bool     hasBlank  = false;
                bool     hasNumber = false;
                bool     hasString = false;
                bool     allInt    = true;
                double   minVal    = std::numeric_limits<double>::max();
                double   maxVal    = std::numeric_limits<double>::lowest();
                uint32_t valCount  = 0;

                for (const auto& val : uniques) {
                    if (val.type() == XLValueType::Empty) {
                        hasBlank = true;
                        sharedItemsNode.append_child("m");
                    } else if (val.type() == XLValueType::Boolean) {
                        auto bNode = sharedItemsNode.append_child("b");
                        bNode.append_attribute("v").set_value(val.get<bool>() ? "1" : "0");
                        allInt = false;
                        valCount++;
                    } else if (val.type() == XLValueType::Integer) {
                        auto nNode = sharedItemsNode.append_child("n");
                        nNode.append_attribute("v").set_value(std::to_string(val.get<int64_t>()).c_str());
                        double d = static_cast<double>(val.get<int64_t>());
                        if (d < minVal) minVal = d;
                        if (d > maxVal) maxVal = d;
                        hasNumber = true;
                        valCount++;
                    } else if (val.type() == XLValueType::Float) {
                        auto nNode = sharedItemsNode.append_child("n");
                        nNode.append_attribute("v").set_value(fmt::format("{}", val.get<double>()).c_str());
                        double d = val.get<double>();
                        if (d < minVal) minVal = d;
                        if (d > maxVal) maxVal = d;
                        hasNumber = true;
                        allInt    = false;
                        valCount++;
                    } else {
                        auto sNode = sharedItemsNode.append_child("s");
                        sNode.append_attribute("v").set_value(val.get<std::string>().c_str());
                        hasString = true;
                        allInt    = false;
                        valCount++;
                    }
                }
                // For all-numeric axis fields (e.g. Year as a filter), add Excel-style
                // numeric metadata attributes — must match what Excel writes after repair.
                if (hasNumber && !hasString && !hasBlank) {
                    auto fmtNum = [](double v) -> std::string {
                        if (v == std::floor(v) && std::abs(v) < 1e15)
                            return std::to_string(static_cast<int64_t>(v));
                        return fmt::format("{}", v);
                    };
                    sharedItemsNode.append_attribute("containsSemiMixedTypes").set_value("0");
                    sharedItemsNode.append_attribute("containsString").set_value("0");
                    sharedItemsNode.append_attribute("containsNumber").set_value("1");
                    if (allInt) sharedItemsNode.append_attribute("containsInteger").set_value("1");
                    sharedItemsNode.append_attribute("minValue").set_value(fmtNum(minVal).c_str());
                    sharedItemsNode.append_attribute("maxValue").set_value(fmtNum(maxVal).c_str());
                } else if (hasBlank) {
                    sharedItemsNode.append_attribute("containsBlank").set_value("1");
                }
                sharedItemsNode.append_attribute("count").set_value(valCount);
            }
            fieldColIdx++;
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
                    } else if (tagName == "b" || tagName == "n" || tagName == "s" || tagName == "d" || tagName == "e") {
                        // OOXML shared items use @v (not @val).
                        const char* v = item.attribute("v").value();
                        if (!v || !*v) v = item.attribute("val").value();    // tolerate legacy mis-writes
                        uniqueStrs.emplace_back(v ? v : "");
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

    auto findFieldIndex = [&](const std::string& fieldName) -> int {
        auto it = std::find(headers.begin(), headers.end(), fieldName);
        if (it != headers.end()) return gsl::narrow_cast<int>(std::distance(headers.begin(), it));
        return -1;
    };
    auto requireFieldIndex = [&](const std::string& fieldName, const char* role) -> int {
        int idx = findFieldIndex(fieldName);
        if (idx < 0)
            throw XLInputError(std::string("XLWorksheet::addPivotTable: ") + role + " field \"" + fieldName +
                               "\" not found in source headers");
        return idx;
    };

    for (const auto& rowFld : options.rows()) {
        rowIndices.push_back(requireFieldIndex(rowFld.name, "row"));
    }
    for (const auto& colFld : options.columns()) {
        colIndices.push_back(requireFieldIndex(colFld.name, "column"));
    }
    for (const auto& dataFld : options.data()) {
        dataIndices.push_back(requireFieldIndex(dataFld.name, "data"));
    }
    for (const auto& filterFld : options.filters()) {
        filterIndices.push_back(requireFieldIndex(filterFld.name, "filter"));
    }

    // SelectedItems must exist in the field's distinct values (when shared items are available).
    auto validateSelected = [&](const XLPivotField& fld, int fieldIdx) {
        if (fld.selectedItems.empty() || fieldIdx < 0 ||
            static_cast<size_t>(fieldIdx) >= sharedItemStrings.size())
            return;
        const auto& uniques = sharedItemStrings[static_cast<size_t>(fieldIdx)];
        if (uniques.empty()) return;
        for (const auto& sel : fld.selectedItems) {
            if (std::find(uniques.begin(), uniques.end(), sel) == uniques.end()) {
                throw XLInputError("XLWorksheet::addPivotTable: selected item \"" + sel + "\" not found in field \"" +
                                   fld.name + "\"");
            }
        }
    };
    for (const auto& rf : options.rows()) validateSelected(rf, findFieldIndex(rf.name));
    for (const auto& cf : options.columns()) validateSelected(cf, findFieldIndex(cf.name));
    for (const auto& ff : options.filters()) validateSelected(ff, findFieldIndex(ff.name));

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

    auto applyAxisLayout = [&](XMLNode ptFieldNode, const XLPivotField* fld) {
        const bool classic = options.classicLayout();
        const bool compact = classic ? false : (fld ? fld->compact : false);
        const bool outline = classic ? false : (fld ? fld->outline : false);
        const bool showAll = fld ? fld->showAll : false;
        ptFieldNode.append_attribute("compact").set_value(compact ? "1" : "0");
        ptFieldNode.append_attribute("outline").set_value(outline ? "1" : "0");
        ptFieldNode.append_attribute("showAll").set_value(showAll ? "1" : "0");
        if (fld && fld->insertBlankRow) ptFieldNode.append_attribute("insertBlankRow").set_value("1");
        if (fld && !fld->defaultSubtotal) ptFieldNode.append_attribute("defaultSubtotal").set_value("0");
        if (fld && !fld->customName.empty()) ptFieldNode.append_attribute("name").set_value(fld->customName.c_str());
    };

    int i = 0;
    for (const auto& h : headers) {
        XMLNode ptFieldNode = pivotFieldsNode.append_child("pivotField");

        bool isRow    = std::find(rowIndices.begin(), rowIndices.end(), i) != rowIndices.end();
        bool isCol    = std::find(colIndices.begin(), colIndices.end(), i) != colIndices.end();
        bool isData   = std::find(dataIndices.begin(), dataIndices.end(), i) != dataIndices.end();
        bool isFilter = std::find(filterIndices.begin(), filterIndices.end(), i) != filterIndices.end();

        if (isRow) {
            ptFieldNode.append_attribute("axis").set_value("axisRow");
            const XLPivotField* rf = findRowField(h);
            applyAxisLayout(ptFieldNode, rf);
            buildItemsNode(ptFieldNode, i, rf ? rf->selectedItems : std::vector<std::string>{});
        }
        else if (isCol) {
            ptFieldNode.append_attribute("axis").set_value("axisCol");
            const XLPivotField* cf = findColField(h);
            applyAxisLayout(ptFieldNode, cf);
            buildItemsNode(ptFieldNode, i, cf ? cf->selectedItems : std::vector<std::string>{});
        }
        else if (isFilter) {
            ptFieldNode.append_attribute("axis").set_value("axisPage");
            const XLPivotField* ff = nullptr;
            for (const auto& f : options.filters())
                if (f.name == h) {
                    ff = &f;
                    break;
                }
            applyAxisLayout(ptFieldNode, ff);

            // Page fields: list member indices when deep shared items exist.
            size_t memberCount =
                (static_cast<size_t>(i) < sharedItemStrings.size()) ? sharedItemStrings[i].size() : 0;
            XMLNode  itemsNode = ptFieldNode.append_child("items");
            uint32_t itemCount = static_cast<uint32_t>(memberCount) + 1;
            itemsNode.append_attribute("count").set_value(itemCount);
            for (size_t xi = 0; xi < memberCount; ++xi) {
                XMLNode item = itemsNode.append_child("item");
                item.append_attribute("x").set_value(static_cast<uint32_t>(xi));
            }
            itemsNode.append_child("item").append_attribute("t").set_value("default");
        }
        else if (isData) {
            ptFieldNode.append_attribute("dataField").set_value("1");
            ptFieldNode.append_attribute("showAll").set_value("0");
            if (options.classicLayout()) {
                ptFieldNode.append_attribute("compact").set_value("0");
                ptFieldNode.append_attribute("outline").set_value("0");
            }
        }
        else {
            ptFieldNode.append_attribute("showAll").set_value("0");
            if (options.classicLayout()) {
                ptFieldNode.append_attribute("compact").set_value("0");
                ptFieldNode.append_attribute("outline").set_value("0");
            }
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

        // Build proper rowItems: one entry per visible row + grand total
        // Excel's repaired format: <i><x/></i> per data row, <i t="grand"><x/></i> for total
        uint32_t rItemCount = 0;
        for (int idx : rowIndices) {
            size_t memberCount = (static_cast<size_t>(idx) < sharedItemStrings.size())
                                     ? sharedItemStrings[idx].size() : 1;
            // Check if there are selected (filtered) items
            const XLPivotField* rf = findRowField(headers[idx]);
            size_t visibleCount = memberCount;
            if (rf && !rf->selectedItems.empty()) {
                visibleCount = rf->selectedItems.size();
            }
            rItemCount += static_cast<uint32_t>(visibleCount);
        }
        if (options.rowGrandTotals()) rItemCount += (hasVirtualDataField && virtualDataOnRows ? dataIndices.size() : 1);
        if (rItemCount == 0) rItemCount = 1;
        rowItemsNode.append_attribute("count").set_value(rItemCount);

        // Add placeholder data rows
        uint32_t rowDataItems = rItemCount - (options.rowGrandTotals() ? 1 : 0);
        for (uint32_t ri = 0; ri < rowDataItems; ++ri) {
            XMLNode rowItem = rowItemsNode.append_child("i");
            rowItem.append_child("x");
        }
        // Add grand total row
        if (options.rowGrandTotals()) {
            XMLNode grandRow = rowItemsNode.append_child("i");
            grandRow.append_attribute("t").set_value("grand");
            grandRow.append_child("x");
        }
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

        // Build proper colItems matching Excel's repaired format.
        // colMultiplier = number of data fields when Values is on columns axis (multiple data fields)
        uint32_t localColMultiplier = (hasVirtualDataField && !virtualDataOnRows)
                                          ? static_cast<uint32_t>(dataIndices.size()) : 1;

        // Compute visible column items count (same logic as in grid-size calculation below)
        uint32_t localColItemsCount = 0;
        for (int idx : colIndices) {
            const XLPivotField* cf = findColField(headers[idx]);
            if (cf && !cf->selectedItems.empty()) {
                localColItemsCount += static_cast<uint32_t>(cf->selectedItems.size());
            } else if (static_cast<size_t>(idx) < sharedItemStrings.size()) {
                localColItemsCount += static_cast<uint32_t>(sharedItemStrings[idx].size());
            }
        }
        if (localColItemsCount == 0) localColItemsCount = 1;

        uint32_t colDataItems = localColItemsCount * localColMultiplier;
        uint32_t totalColItems = colDataItems + (options.colGrandTotals() ? localColMultiplier : 0);
        if (totalColItems == 0) totalColItems = 1;
        colItemsNode.append_attribute("count").set_value(totalColItems);

        // First item: all-zero indices
        {
            XMLNode colItem = colItemsNode.append_child("i");
            for (size_t ci = 0; ci < colIndices.size(); ++ci) colItem.append_child("x");
            if (hasVirtualDataField && !virtualDataOnRows) colItem.append_child("x");
        }
        // Remaining data items
        for (uint32_t ci = 1; ci < colDataItems; ++ci) {
            XMLNode colItem = colItemsNode.append_child("i");
            uint32_t dataIdx = ci % localColMultiplier;
            if (dataIdx > 0) {
                colItem.append_attribute("r").set_value("1");
                colItem.append_attribute("i").set_value(dataIdx);
                XMLNode xNode = colItem.append_child("x");
                xNode.append_attribute("v").set_value(dataIdx);
            } else {
                XMLNode xNode = colItem.append_child("x");
                xNode.append_attribute("v").set_value(ci / localColMultiplier);
                colItem.append_child("x");
            }
        }
        // Grand total items
        for (uint32_t gi = 0; options.colGrandTotals() && gi < localColMultiplier; ++gi) {
            XMLNode colItem = colItemsNode.append_child("i");
            colItem.append_attribute("t").set_value("grand");
            if (gi > 0) colItem.append_attribute("i").set_value(gi);
            colItem.append_child("x");
        }
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

            const char* subType = pivotSubtotalToString(dataFld.subtotal);
            std::string prefix  = "Sum of ";
            switch (dataFld.subtotal) {
                case XLPivotSubtotal::Average: prefix = "Average of "; break;
                case XLPivotSubtotal::Count: prefix = "Count of "; break;
                case XLPivotSubtotal::Max: prefix = "Max of "; break;
                case XLPivotSubtotal::Min: prefix = "Min of "; break;
                case XLPivotSubtotal::Product: prefix = "Product of "; break;
                case XLPivotSubtotal::CountNums: prefix = "Count Nums of "; break;
                case XLPivotSubtotal::StdDev: prefix = "StdDev of "; break;
                case XLPivotSubtotal::StdDevP: prefix = "StdDevP of "; break;
                case XLPivotSubtotal::Var: prefix = "Var of "; break;
                case XLPivotSubtotal::VarP: prefix = "VarP of "; break;
                default: prefix = "Sum of "; break;
            }

            std::string dName = dataFld.customName.empty() ? (prefix + dataFld.name) : dataFld.customName;
            dfield.append_attribute("name").set_value(dName.c_str());
            dfield.append_attribute("fld").set_value(idx);

            if (dataFld.subtotal != XLPivotSubtotal::Sum) {
                dfield.append_attribute("subtotal").set_value(subType);
            }

            // Show Values As
            const auto sva = dataFld.showValuesAs;
            int        baseFieldIdx = 0;
            int        baseItemIdx  = 0;
            if (sva.type != XLPivotShowValuesAs::Normal) {
                if (pivotShowValuesAsNeedsBaseField(sva.type)) {
                    baseFieldIdx = findFieldIndex(sva.baseField);
                    if (baseFieldIdx < 0)
                        throw XLInputError("XLWorksheet::addPivotTable: ShowValuesAs baseField \"" + sva.baseField +
                                           "\" not found");
                }
                if (pivotShowValuesAsNeedsBaseItem(sva.type)) {
                    if (baseFieldIdx < 0 || static_cast<size_t>(baseFieldIdx) >= sharedItemStrings.size())
                        throw XLInputError("XLWorksheet::addPivotTable: ShowValuesAs baseItem requires shared items");
                    const auto& labels = sharedItemStrings[static_cast<size_t>(baseFieldIdx)];
                    auto        it     = std::find(labels.begin(), labels.end(), sva.baseItem);
                    if (it == labels.end())
                        throw XLInputError("XLWorksheet::addPivotTable: ShowValuesAs baseItem \"" + sva.baseItem +
                                           "\" not found in field \"" + sva.baseField + "\"");
                    baseItemIdx = static_cast<int>(std::distance(labels.begin(), it));
                }

                const char* showStr = pivotShowValuesAsToString(sva.type);
                if (pivotShowValuesAsNeedsX14(sva.type)) {
                    // Modern types live in x14:dataField@pivotShowAs
                    XMLNode extLst = dfield.append_child("extLst");
                    XMLNode ext    = extLst.append_child("ext");
                    ext.append_attribute("uri").set_value("{E15A36E0-9728-4e99-A89B-3F7291B0FE68}");
                    ext.append_attribute("xmlns:x14")
                        .set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
                    XMLNode x14 = ext.append_child("x14:dataField");
                    x14.append_attribute("pivotShowAs").set_value(showStr);
                }
                else {
                    dfield.append_attribute("showDataAs").set_value(showStr);
                }
                dfield.append_attribute("baseField").set_value(baseFieldIdx);
                dfield.append_attribute("baseItem").set_value(baseItemIdx);
            }
            else {
                // Excel-compatible defaults for normal aggregates
                dfield.append_attribute("baseField").set_value("0");
                dfield.append_attribute("baseItem").set_value("0");
            }

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

    // For the ref bounding box, use total member count (including hidden rows/cols),
    // as Excel includes hidden members in the ref range.
    uint32_t totalMemberRows = 0;
    for (int idx : rowIndices) {
        if (static_cast<size_t>(idx) < sharedItemStrings.size())
            totalMemberRows += static_cast<uint32_t>(sharedItemStrings[idx].size());
        else
            totalMemberRows += rowItemsCount;
    }
    if (totalMemberRows == 0) totalMemberRows = rowItemsCount;
    uint32_t totalDataRows = totalMemberRows * rowMultiplier;
    if (options.rowGrandTotals()) totalDataRows += rowMultiplier;

    // Header height (H) and Row labels width (R_cols)
    uint32_t H = colIndices.size() + ((hasVirtualDataField && !virtualDataOnRows) ? 1 : 0);
    if (H == 0) H = 1;

    uint32_t R_cols = rowIndices.size() + ((hasVirtualDataField && virtualDataOnRows) ? 1 : 0);
    if (R_cols == 0) R_cols = 1;

    uint32_t firstHeaderRow = 1;
    uint32_t firstDataRow = H + 1;
    // firstDataCol: Excel uses 1 (data starts at first column of the ref area).
    // Our row-label columns are part of the grid but Excel's firstDataCol=1 means
    // the first data value column index (1-based within the ref).
    uint32_t firstDataCol = 1;

    uint32_t gridWidth = R_cols + totalDataCols;
    uint32_t gridHeight = H + totalDataRows;

    // Explicit location range (excelize PivotTableRange) wins over estimated box.
    std::string ptRange = options.locationRange();
    if (ptRange.empty() && options.targetCell().find(':') != std::string::npos)
        ptRange = options.targetCell();

    if (ptRange.empty()) {
        XLCellReference targetRef(options.targetCell());
        uint32_t        startRow = targetRef.row();
        uint16_t        startCol = targetRef.column();

        // The grid area starts after filter fields plus a blank spacer row
        uint32_t gridStartRow = startRow;
        if (numFilters > 0) {
            gridStartRow += numFilters + 1;
        }

        XLCellReference gridStartRef(gridStartRow, startCol);
        XLCellReference gridEndRef(gridStartRow + gridHeight - 1, startCol + gridWidth - 1);
        ptRange = gridStartRef.address() + ":" + gridEndRef.address();
    }

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

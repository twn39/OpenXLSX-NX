// ===== External Includes ===== //
#include <algorithm>
#include <cstring>
#include <pugixml.hpp>
#include <unordered_map>

// ===== OpenXLSX Includes ===== //
#include "XLContentTypes.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLPivotTable.hpp"
#include "XLRelationships.hpp"
#include "XLXmlData.hpp"

namespace OpenXLSX
{
    const char* pivotShowValuesAsToString(XLPivotShowValuesAs t)
    {
        switch (t) {
            case XLPivotShowValuesAs::PercentOfGrandTotal: return "percentOfTotal";
            case XLPivotShowValuesAs::PercentOfColumnTotal: return "percentOfCol";
            case XLPivotShowValuesAs::PercentOfRowTotal: return "percentOfRow";
            case XLPivotShowValuesAs::PercentOf: return "percent";
            case XLPivotShowValuesAs::PercentOfParentRowTotal: return "percentOfParentRow";
            case XLPivotShowValuesAs::PercentOfParentColumnTotal: return "percentOfParentCol";
            case XLPivotShowValuesAs::PercentOfParentTotal: return "percentOfParent";
            case XLPivotShowValuesAs::DifferenceFrom: return "difference";
            case XLPivotShowValuesAs::PercentDifferenceFrom: return "percentDiff";
            case XLPivotShowValuesAs::RunningTotalIn: return "runTotal";
            case XLPivotShowValuesAs::PercentRunningTotalIn: return "percentOfRunningTotal";
            case XLPivotShowValuesAs::RankSmallestToLargest: return "rankAscending";
            case XLPivotShowValuesAs::RankLargestToSmallest: return "rankDescending";
            case XLPivotShowValuesAs::Index: return "index";
            default: return "normal";
        }
    }

    XLPivotShowValuesAs pivotShowValuesAsFromString(std::string_view s)
    {
        if (s == "percentOfTotal") return XLPivotShowValuesAs::PercentOfGrandTotal;
        if (s == "percentOfCol") return XLPivotShowValuesAs::PercentOfColumnTotal;
        if (s == "percentOfRow") return XLPivotShowValuesAs::PercentOfRowTotal;
        if (s == "percent") return XLPivotShowValuesAs::PercentOf;
        if (s == "percentOfParentRow") return XLPivotShowValuesAs::PercentOfParentRowTotal;
        if (s == "percentOfParentCol") return XLPivotShowValuesAs::PercentOfParentColumnTotal;
        if (s == "percentOfParent") return XLPivotShowValuesAs::PercentOfParentTotal;
        if (s == "difference") return XLPivotShowValuesAs::DifferenceFrom;
        if (s == "percentDiff") return XLPivotShowValuesAs::PercentDifferenceFrom;
        if (s == "runTotal") return XLPivotShowValuesAs::RunningTotalIn;
        if (s == "percentOfRunningTotal") return XLPivotShowValuesAs::PercentRunningTotalIn;
        if (s == "rankAscending") return XLPivotShowValuesAs::RankSmallestToLargest;
        if (s == "rankDescending") return XLPivotShowValuesAs::RankLargestToSmallest;
        if (s == "index") return XLPivotShowValuesAs::Index;
        return XLPivotShowValuesAs::Normal;
    }

    bool pivotShowValuesAsNeedsX14(XLPivotShowValuesAs t)
    {
        switch (t) {
            case XLPivotShowValuesAs::PercentOfParentRowTotal:
            case XLPivotShowValuesAs::PercentOfParentColumnTotal:
            case XLPivotShowValuesAs::PercentOfParentTotal:
            case XLPivotShowValuesAs::PercentRunningTotalIn:
            case XLPivotShowValuesAs::RankSmallestToLargest:
            case XLPivotShowValuesAs::RankLargestToSmallest:
            case XLPivotShowValuesAs::Index:
                return true;
            default:
                return false;
        }
    }

    bool pivotShowValuesAsNeedsBaseField(XLPivotShowValuesAs t)
    {
        switch (t) {
            case XLPivotShowValuesAs::PercentOf:
            case XLPivotShowValuesAs::PercentOfParentTotal:
            case XLPivotShowValuesAs::DifferenceFrom:
            case XLPivotShowValuesAs::PercentDifferenceFrom:
            case XLPivotShowValuesAs::RunningTotalIn:
            case XLPivotShowValuesAs::PercentRunningTotalIn:
            case XLPivotShowValuesAs::RankSmallestToLargest:
            case XLPivotShowValuesAs::RankLargestToSmallest:
                return true;
            default:
                return false;
        }
    }

    bool pivotShowValuesAsNeedsBaseItem(XLPivotShowValuesAs t)
    {
        switch (t) {
            case XLPivotShowValuesAs::PercentOf:
            case XLPivotShowValuesAs::DifferenceFrom:
            case XLPivotShowValuesAs::PercentDifferenceFrom:
                return true;
            default:
                return false;
        }
    }

    const char* pivotSubtotalToString(XLPivotSubtotal s)
    {
        switch (s) {
            case XLPivotSubtotal::Average: return "average";
            case XLPivotSubtotal::Count: return "count";
            case XLPivotSubtotal::Max: return "max";
            case XLPivotSubtotal::Min: return "min";
            case XLPivotSubtotal::Product: return "product";
            case XLPivotSubtotal::CountNums: return "countNums";
            case XLPivotSubtotal::StdDev: return "stdDev";
            case XLPivotSubtotal::StdDevP: return "stdDevp";
            case XLPivotSubtotal::Var: return "var";
            case XLPivotSubtotal::VarP: return "varp";
            default: return "sum";
        }
    }

    XLPivotSubtotal pivotSubtotalFromString(std::string_view s)
    {
        if (s == "average") return XLPivotSubtotal::Average;
        if (s == "count") return XLPivotSubtotal::Count;
        if (s == "max") return XLPivotSubtotal::Max;
        if (s == "min") return XLPivotSubtotal::Min;
        if (s == "product") return XLPivotSubtotal::Product;
        if (s == "countNums") return XLPivotSubtotal::CountNums;
        if (s == "stdDev") return XLPivotSubtotal::StdDev;
        if (s == "stdDevp") return XLPivotSubtotal::StdDevP;
        if (s == "var") return XLPivotSubtotal::Var;
        if (s == "varp") return XLPivotSubtotal::VarP;
        return XLPivotSubtotal::Sum;
    }

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
        if (locNode) {
            std::string refVal = locNode.attribute("ref").value();
            auto colonPos = refVal.find(':');
            if (colonPos != std::string::npos) {
                return refVal.substr(0, colonPos);
            }
            return refVal;
        }
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
            // Package port: relationships + managed XML parts (no full document façade).
            XLPackageServices& pkg = const_cast<XLPivotTable*>(this)->package();
            std::string cacheTargetPath = pkg.workbookRelationships().relationshipById(rId).target();
            if (!cacheTargetPath.empty() && cacheTargetPath[0] != '/') cacheTargetPath = "/xl/" + cacheTargetPath;
            std::string targetPath = !cacheTargetPath.empty() && cacheTargetPath[0] == '/' ? cacheTargetPath.substr(1) : cacheTargetPath;

            XLXmlData* xmlData = pkg.findXmlPart(targetPath, /*doNotThrow=*/true);
            if (xmlData == nullptr) {
                xmlData = pkg.emplaceXmlPart(targetPath, rId, XLContentType::PivotCacheDefinition);
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
            XLPackageServices& pkg = package();
            std::string cacheTargetPath = pkg.workbookRelationships().relationshipById(rId).target();
            // Resolve relative path to absolute
            if (!cacheTargetPath.empty() && cacheTargetPath[0] != '/') cacheTargetPath = "/xl/" + cacheTargetPath;
            std::string targetPath = !cacheTargetPath.empty() && cacheTargetPath[0] == '/' ? cacheTargetPath.substr(1) : cacheTargetPath;

            // Get the xml document for cache via package part store
            XLPivotCacheDefinition cacheDef(pkg.findXmlPart(targetPath));
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
        if (!sourceNode) return {};

        // Table / defined-name sources use @name (excelize / Excel dynamic sources).
        std::string nameAttr = sourceNode.attribute("name").value();
        if (!nameAttr.empty()) return nameAttr;

        std::string sheet = sourceNode.attribute("sheet").value();
        std::string ref   = sourceNode.attribute("ref").value();
        if (sheet.empty()) return ref;
        return sheet + "!" + ref;
    }

    void XLPivotCacheDefinition::changeSourceRange(std::string_view newRange)
    {
        XMLNode sourceNode = xmlDocument().document_element().child("cacheSource").child("worksheetSource");
        if (!sourceNode) throw XLInternalError("Could not locate worksheetSource in pivot cache.");

        std::string rangeStr(newRange);
        auto        bangPos = rangeStr.find('!');
        if (bangPos == std::string::npos) {
            // Treat as table / defined name binding.
            if (sourceNode.attribute("name")) sourceNode.attribute("name").set_value(rangeStr.c_str());
            else sourceNode.append_attribute("name").set_value(rangeStr.c_str());
            if (sourceNode.attribute("ref")) sourceNode.remove_attribute("ref");
            if (sourceNode.attribute("sheet")) sourceNode.remove_attribute("sheet");
            return;
        }

        std::string sheet = rangeStr.substr(0, bangPos);
        std::string ref   = rangeStr.substr(bangPos + 1);

        if (sheet.length() >= 2 && sheet.front() == '\'' && sheet.back() == '\'') {
            sheet = sheet.substr(1, sheet.length() - 2);
        }

        if (sourceNode.attribute("name")) sourceNode.remove_attribute("name");
        if (sourceNode.attribute("sheet")) sourceNode.attribute("sheet").set_value(sheet.c_str());
        else sourceNode.append_attribute("sheet").set_value(sheet.c_str());
        if (sourceNode.attribute("ref")) sourceNode.attribute("ref").set_value(ref.c_str());
        else sourceNode.append_attribute("ref").set_value(ref.c_str());
    }

    std::vector<std::string> XLPivotCacheDefinition::fieldNames() const
    {
        std::vector<std::string> names;
        XMLNode fields = xmlDocument().document_element().child("cacheFields");
        for (auto cf : fields.children("cacheField")) {
            names.emplace_back(cf.attribute("name").value());
        }
        return names;
    }

    namespace
    {
        std::vector<std::string> readSharedItemLabels(XMLNode sharedItems)
        {
            std::vector<std::string> labels;
            if (sharedItems.empty()) return labels;
            for (auto item : sharedItems.children()) {
                std::string tag = item.name();
                if (tag == "m") {
                    labels.emplace_back("");
                }
                else if (tag == "b" || tag == "n" || tag == "s" || tag == "d" || tag == "e") {
                    const char* v = item.attribute("v").value();
                    if (!v || !*v) v = item.attribute("val").value();
                    labels.emplace_back(v ? v : "");
                }
            }
            return labels;
        }

        void extractSelectedItems(XMLNode pivotField, const std::vector<std::string>& labels, XLPivotField& out)
        {
            XMLNode items = pivotField.child("items");
            if (items.empty() || labels.empty()) return;
            for (auto item : items.children("item")) {
                if (item.attribute("t")) continue;    // default/grand etc.
                if (!item.attribute("x")) continue;
                const unsigned x = item.attribute("x").as_uint(UINT32_MAX);
                if (x >= labels.size()) continue;
                // hidden items have h="1"; selected = not hidden
                if (item.attribute("h").as_bool(false)) continue;
                out.selectedItems.push_back(labels[x]);
            }
            // If every non-default item was visible, selectedItems equals full set — leave empty
            // to mean "all selected" when all visible.
            if (out.selectedItems.size() == labels.size()) out.selectedItems.clear();
        }

        XLPivotShowValuesAsOptions extractShowValuesAs(XMLNode dataField, const std::vector<std::string>& cacheNames)
        {
            XLPivotShowValuesAsOptions sva;
            std::string show = dataField.attribute("showDataAs").value();
            if (show.empty() || show == "normal") {
                // x14:dataField pivotShowAs
                XMLNode extLst = dataField.child("extLst");
                for (auto ext : extLst.children("ext")) {
                    for (auto child : ext.children()) {
                        std::string n = child.name();
                        if (n.find("dataField") != std::string::npos) {
                            const char* ps = child.attribute("pivotShowAs").value();
                            if (ps && *ps) show = ps;
                        }
                    }
                }
            }
            if (!show.empty() && show != "normal") sva.type = pivotShowValuesAsFromString(show);

            if (dataField.attribute("baseField")) {
                int bf = dataField.attribute("baseField").as_int(-1);
                if (bf >= 0 && static_cast<size_t>(bf) < cacheNames.size()) sva.baseField = cacheNames[static_cast<size_t>(bf)];
            }
            // baseItem index → label requires shared items of base field; leave empty if unknown
            (void)cacheNames;
            return sva;
        }
    }    // namespace

    XLPivotTableOptions XLPivotTable::options() const
    {
        XMLNode root = xmlDocument().document_element();
        XLPivotTableOptions opts(name(), sourceRange(), targetCell());

        XMLNode loc = root.child("location");
        if (loc && loc.attribute("ref")) {
            opts.setLocationRangeInternal(loc.attribute("ref").value());
            std::string ref = loc.attribute("ref").value();
            auto        colon = ref.find(':');
            opts.setTargetCellInternal(colon == std::string::npos ? ref : ref.substr(0, colon));
        }

        bool classic = root.attribute("gridDropZones").as_bool(false);
        opts.setBoolFlags(
            /*dataOnRows*/ false,    // refined below
            root.attribute("rowGrandTotals").as_bool(true),
            root.attribute("colGrandTotals").as_bool(true),
            root.attribute("showDrill").as_bool(true),
            root.attribute("useAutoFormatting").as_bool(false),
            root.attribute("pageOverThenDown").as_bool(false),
            root.attribute("mergeItem").as_bool(false),
            root.attribute("compactData").as_bool(true),
            root.attribute("showError").as_bool(false),
            true,
            true,
            false,
            false,
            false,
            classic,
            root.attribute("fieldPrintTitles").as_bool(false),
            root.attribute("itemPrintTitles").as_bool(false));

        XMLNode style = root.child("pivotTableStyleInfo");
        if (style) {
            if (style.attribute("name")) opts.setPivotTableStyleNameInternal(style.attribute("name").value());
            opts.setBoolFlags(opts.dataOnRows(),
                              opts.rowGrandTotals(),
                              opts.colGrandTotals(),
                              opts.showDrill(),
                              opts.useAutoFormatting(),
                              opts.pageOverThenDown(),
                              opts.mergeItem(),
                              opts.compactData(),
                              opts.showError(),
                              style.attribute("showRowHeaders").as_bool(true),
                              style.attribute("showColHeaders").as_bool(true),
                              style.attribute("showRowStripes").as_bool(false),
                              style.attribute("showColStripes").as_bool(false),
                              style.attribute("showLastColumn").as_bool(false),
                              opts.classicLayout(),
                              opts.fieldPrintTitles(),
                              opts.itemPrintTitles());
        }

        std::vector<std::string> cacheNames;
        std::vector<std::vector<std::string>> sharedLabels;
        try {
            auto cache = cacheDefinition();
            cacheNames = cache.fieldNames();
            XMLNode cfields = cache.xmlDocument().document_element().child("cacheFields");
            for (auto cf : cfields.children("cacheField")) {
                sharedLabels.push_back(readSharedItemLabels(cf.child("sharedItems")));
            }
        }
        catch (...) {
        }

        // Map cache index → pivotField node
        std::vector<XMLNode> pivotFieldNodes;
        for (auto pf : root.child("pivotFields").children("pivotField")) pivotFieldNodes.push_back(pf);

        auto fieldAt = [&](int idx) -> XLPivotField {
            XLPivotField f;
            if (idx >= 0 && static_cast<size_t>(idx) < cacheNames.size()) f.name = cacheNames[static_cast<size_t>(idx)];
            if (idx >= 0 && static_cast<size_t>(idx) < pivotFieldNodes.size()) {
                XMLNode pf = pivotFieldNodes[static_cast<size_t>(idx)];
                f.compact         = pf.attribute("compact").as_bool(false);
                f.outline         = pf.attribute("outline").as_bool(false);
                f.showAll         = pf.attribute("showAll").as_bool(false);
                f.insertBlankRow  = pf.attribute("insertBlankRow").as_bool(false);
                f.defaultSubtotal = pf.attribute("defaultSubtotal").as_bool(true);
                if (pf.attribute("name")) f.customName = pf.attribute("name").value();
                if (static_cast<size_t>(idx) < sharedLabels.size())
                    extractSelectedItems(pf, sharedLabels[static_cast<size_t>(idx)], f);
            }
            return f;
        };

        std::vector<XLPivotField> rows, cols, filters, data;
        for (auto fld : root.child("rowFields").children("field")) {
            int x = fld.attribute("x").as_int(-999);
            if (x == -2) {
                // Values virtual field on rows
                opts.setBoolFlags(true,
                                  opts.rowGrandTotals(),
                                  opts.colGrandTotals(),
                                  opts.showDrill(),
                                  opts.useAutoFormatting(),
                                  opts.pageOverThenDown(),
                                  opts.mergeItem(),
                                  opts.compactData(),
                                  opts.showError(),
                                  opts.showRowHeaders(),
                                  opts.showColHeaders(),
                                  opts.showRowStripes(),
                                  opts.showColStripes(),
                                  opts.showLastColumn(),
                                  opts.classicLayout(),
                                  opts.fieldPrintTitles(),
                                  opts.itemPrintTitles());
                continue;
            }
            if (x >= 0) rows.push_back(fieldAt(x));
        }
        for (auto fld : root.child("colFields").children("field")) {
            int x = fld.attribute("x").as_int(-999);
            if (x == -2) continue;
            if (x >= 0) cols.push_back(fieldAt(x));
        }
        for (auto pf : root.child("pageFields").children("pageField")) {
            int fld = pf.attribute("fld").as_int(-1);
            if (fld >= 0) filters.push_back(fieldAt(fld));
        }
        for (auto df : root.child("dataFields").children("dataField")) {
            int fld = df.attribute("fld").as_int(-1);
            XLPivotField f = fieldAt(fld);
            if (df.attribute("name")) f.customName = df.attribute("name").value();
            if (df.attribute("subtotal")) f.subtotal = pivotSubtotalFromString(df.attribute("subtotal").value());
            if (df.attribute("numFmtId")) f.numFmtId = df.attribute("numFmtId").as_uint(0);
            f.showValuesAs = extractShowValuesAs(df, cacheNames);
            // Resolve baseItem label when possible
            if (pivotShowValuesAsNeedsBaseItem(f.showValuesAs.type) && !f.showValuesAs.baseField.empty()) {
                int bi = df.attribute("baseItem").as_int(-1);
                auto it = std::find(cacheNames.begin(), cacheNames.end(), f.showValuesAs.baseField);
                if (it != cacheNames.end() && bi >= 0) {
                    size_t bfi = static_cast<size_t>(std::distance(cacheNames.begin(), it));
                    if (bfi < sharedLabels.size() && static_cast<size_t>(bi) < sharedLabels[bfi].size())
                        f.showValuesAs.baseItem = sharedLabels[bfi][static_cast<size_t>(bi)];
                }
            }
            data.push_back(std::move(f));
        }

        opts.setRowsInternal(std::move(rows));
        opts.setColumnsInternal(std::move(cols));
        opts.setFiltersInternal(std::move(filters));
        opts.setDataInternal(std::move(data));
        return opts;
    }

    XLPivotCacheRecords::XLPivotCacheRecords(XLXmlData* xmlData) : XLXmlFile(xmlData)
    {
        if (xmlData && xmlData->getXmlType() != XLContentType::PivotCacheRecords) {
            throw XLInternalError("XLPivotCacheRecords constructor: Invalid XML data.");
        }
    }

    XLRelationships XLPivotTable::relationships() { return package().xmlRelationships(getXmlPath()); }

    XLRelationships XLPivotCacheDefinition::relationships() { return package().xmlRelationships(getXmlPath()); }
}    // namespace OpenXLSX


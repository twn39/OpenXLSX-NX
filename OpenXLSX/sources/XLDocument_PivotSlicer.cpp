// ===== XLDocument pivot table & slicer package factories ===== //
// Extracted from XLDocument.cpp: pivot caches/tables/records and slicer/slicerCache creation.

#include "XLDocument.hpp"

#include "XLContentTypes.hpp"
#include "XLDocument_Templates.hpp"
#include "XLException.hpp"
#include "XLPivotTable.hpp"
#include "XLRelationships.hpp"
#include "XLUtilities.hpp"

#include <fmt/format.h>
#include <pugixml.hpp>
#include <string>
#include <string_view>
#include <vector>

using namespace OpenXLSX;

XLPivotTable XLDocument::createPivotTable()
{
    using namespace std::literals::string_literals;

    uint32_t    num      = 1;
    std::string filename = fmt::format("xl/pivotTables/pivotTable{}.xml", num);
    while (m_archive.hasEntry(filename)) {
        ++num;
        filename = fmt::format("xl/pivotTables/pivotTable{}.xml", num);
    }

    m_archive.addEntry(filename, std::string(xlPivotTableTemplate));
    m_contentTypes.addOverride("/" + filename, XLContentType::PivotTable);

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(filename, DO_NOT_THROW);
    if (xmlData == nullptr) {
        xmlData = &m_data.emplace_back(this, filename, std::string(xlPivotTableTemplate), XLContentType::PivotTable);
    }

    return XLPivotTable(xmlData);
}

XLPivotCacheDefinition XLDocument::createPivotCacheDefinition()
{
    using namespace std::literals::string_literals;

    uint32_t    num      = 1;
    std::string filename = fmt::format("xl/pivotCache/pivotCacheDefinition{}.xml", num);
    while (m_archive.hasEntry(filename)) {
        ++num;
        filename = fmt::format("xl/pivotCache/pivotCacheDefinition{}.xml", num);
    }

    m_archive.addEntry(filename, std::string(xlPivotCacheDefinitionTemplate));
    m_contentTypes.addOverride("/" + filename, XLContentType::PivotCacheDefinition);

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(filename, DO_NOT_THROW);
    if (xmlData == nullptr) {
        xmlData = &m_data.emplace_back(this, filename, std::string(xlPivotCacheDefinitionTemplate), XLContentType::PivotCacheDefinition);
    }

    return XLPivotCacheDefinition(xmlData);
}

std::string XLDocument::createTableSlicerCache(uint32_t tableId, uint32_t tableColumnId, std::string_view name, std::string_view sourceName)
{
    using namespace std::literals::string_literals;

    uint32_t    num      = 1;
    std::string filename = fmt::format("xl/slicerCaches/slicerCache{}.xml", num);
    while (m_archive.hasEntry(filename)) {
        ++num;
        filename = fmt::format("xl/slicerCaches/slicerCache{}.xml", num);
    }

    std::string templateStr = fmt::format(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x xr10" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" name="{0}" xr10:uid="{{00000000-0013-0000-FFFF-FFFF0{4}000000}}" sourceName="{1}"><extLst><x:ext uri="{{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:tableSlicerCache tableId="{2}" column="{3}"/></x:ext></extLst>
</slicerCacheDefinition>)",
                                          name,
                                          sourceName,
                                          tableId,
                                          tableColumnId,
                                          num);

    m_archive.addEntry(filename, templateStr);
    m_contentTypes.addOverride("/" + filename, XLContentType::SlicerCache);

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(filename, DO_NOT_THROW);
    if (xmlData == nullptr) { m_data.emplace_back(this, filename, templateStr, XLContentType::SlicerCache); }

    // Add relationship to workbook.xml (relative path, no leading /)
    std::string relTarget = filename.substr(3);  // strip "xl/" prefix → "slicerCaches/slicerCache1.xml"
    m_wbkRelationships.addRelationship(XLRelationshipType::SlicerCache, relTarget);
    std::string rId = m_wbkRelationships.relationshipByTarget(relTarget).id();

    // Add extLst to workbook.xml to reference the slicer cache
    // TABLE slicers are X15 features (Excel 2013+): use ExtURISlicerCachesX15 = {46BE6895-7355-4a93-B00E-2C351335B9C9}
    // with x15: prefixes throughout. This is DIFFERENT from pivot slicers which use {BBE1A952...} (X14).
    XMLNode wbkNode = m_workbook.xmlDocument().document_element();
    XMLNode extLst  = wbkNode.child("extLst");
    // Table slicers need xmlns:x15 on workbook root (X15 = Excel 2013+)
    if (!wbkNode.attribute("xmlns:x15"))
        wbkNode.append_attribute("xmlns:x15").set_value("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
    if (!wbkNode.attribute("xmlns:mc"))
        wbkNode.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
    // mc:Ignorable must list x15 so older readers skip unknown x15 elements
    if (!wbkNode.attribute("mc:Ignorable")) {
        wbkNode.append_attribute("mc:Ignorable").set_value("x15");
    } else {
        std::string ign = wbkNode.attribute("mc:Ignorable").value();
        if (ign.find("x15") == std::string::npos)
            wbkNode.attribute("mc:Ignorable").set_value((ign + " x15").c_str());
    }

    if (extLst.empty()) extLst = wbkNode.append_child("extLst");

    // Excel always writes x14:workbookPr ext AFTER {BBE1A952} pivot cache and BEFORE {46BE6895} table caches.
    // So we append it (it gets pushed before the table cache which is appended after).
    if (extLst.find_child_by_attribute("uri", "{79F54976-1DA5-4618-B147-4CDE4B953A38}").empty()) {
        XMLNode wbkPrExt = extLst.append_child("ext");
        wbkPrExt.append_attribute("uri").set_value("{79F54976-1DA5-4618-B147-4CDE4B953A38}");
        wbkPrExt.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
        wbkPrExt.append_child("x14:workbookPr");
    }

    // URI {46BE6895-7355-4a93-B00E-2C351335B9C9}: table slicer cache list in workbook.xml
    // Excel: uri first, then xmlns:x15; x15:slicerCaches contains x14:slicerCache with xmlns:x14 inline
    XMLNode ext = extLst.find_child_by_attribute("uri", "{46BE6895-7355-4a93-B00E-2C351335B9C9}");
    if (ext.empty()) {
        ext = extLst.append_child("ext");
        ext.append_attribute("uri").set_value("{46BE6895-7355-4a93-B00E-2C351335B9C9}");
        ext.append_attribute("xmlns:x15").set_value("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
    }

    XMLNode slicerCaches = ext.child("x15:slicerCaches");
    if (slicerCaches.empty()) {
        slicerCaches = ext.append_child("x15:slicerCaches");
        // Excel declares xmlns:x14 inline on x15:slicerCaches
        slicerCaches.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
    }

    // Excel uses x14:slicerCache inside x15:slicerCaches (not x15:slicerCache)
    slicerCaches.append_child("x14:slicerCache").append_attribute("r:id").set_value(rId.c_str());

    return filename;
}

std::string XLDocument::createPivotSlicerCache(uint32_t         pivotCacheId,
                                               uint32_t         sheetId,
                                               std::string_view pivotTableName,
                                               std::string_view name,
                                               std::string_view sourceName)
{
    using namespace std::literals::string_literals;

    uint32_t    num      = 1;
    std::string filename = fmt::format("xl/slicerCaches/slicerCache{}.xml", num);
    while (m_archive.hasEntry(filename)) {
        ++num;
        filename = fmt::format("xl/slicerCaches/slicerCache{}.xml", num);
    }

    std::string templateStr = fmt::format(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x xr10" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" name="{0}" xr10:uid="{{00000000-0013-0000-FFFF-FFFF0{5}000000}}" sourceName="{1}">
  <pivotTables>
    <pivotTable tabId="{2}" name="{3}"/>
  </pivotTables>
  <data>
    <tabular pivotCacheId="{4}" showMissing="0">
      <items count="1">
        <i x="0" s="1"/>
      </items>
    </tabular>
  </data>
</slicerCacheDefinition>)",
                                          name,
                                          sourceName,
                                          sheetId,
                                          pivotTableName,
                                          pivotCacheId,
                                          num);

    m_archive.addEntry(filename, templateStr);
    m_contentTypes.addOverride("/" + filename, XLContentType::SlicerCache);

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(filename, DO_NOT_THROW);
    if (xmlData == nullptr) { m_data.emplace_back(this, filename, templateStr, XLContentType::SlicerCache); }

    // Add relationship to workbook.xml (relative path, no leading /)
    std::string relTarget = filename.substr(3);  // strip "xl/" → "slicerCaches/slicerCache1.xml"
    m_wbkRelationships.addRelationship(XLRelationshipType::SlicerCache, relTarget);
    std::string rId = m_wbkRelationships.relationshipByTarget(relTarget).id();

    // Add extLst to workbook.xml to reference the slicer cache
    XMLNode wbkNode = m_workbook.xmlDocument().document_element();
    XMLNode extLst  = wbkNode.child("extLst");
    if (!wbkNode.attribute("xmlns:mc"))
        wbkNode.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
    if (!wbkNode.attribute("xmlns:x15"))
        wbkNode.append_attribute("xmlns:x15").set_value("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
    if (!wbkNode.attribute("mc:Ignorable")) { wbkNode.append_attribute("mc:Ignorable").set_value("x15"); }
    else {
        std::string ign = wbkNode.attribute("mc:Ignorable").value();
        if (ign.find("x15") == std::string::npos) wbkNode.attribute("mc:Ignorable").set_value((ign + " x15").c_str());
    }

    if (extLst.empty()) extLst = wbkNode.append_child("extLst");

    XMLNode ext = extLst.find_child_by_attribute("uri", "{BBE1A952-AA13-448e-AADC-164F8A28A991}");
    if (ext.empty()) {
        // {BBE1A952} must be FIRST in extLst (before workbookPr and table caches).
        // Use prepend_child so it appears before any existing ext nodes.
        ext = extLst.prepend_child("ext");
        ext.append_attribute("uri").set_value("{BBE1A952-AA13-448e-AADC-164F8A28A991}");
        ext.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
    }

    XMLNode slicerCaches = ext.child("x14:slicerCaches");
    if (slicerCaches.empty()) {
        slicerCaches = ext.append_child("x14:slicerCaches");
    }

    slicerCaches.append_child("x14:slicerCache").append_attribute("r:id").set_value(rId.c_str());

    return filename;
}

std::string XLDocument::findOrCreateTableSlicerCache(uint32_t tableId, uint32_t tableColumnId, std::string_view name, std::string_view sourceName)
{
    for (auto& item : m_data) {
        if (item.getXmlType() == XLContentType::SlicerCache) {
            auto root = item.getXmlDocument()->document_element();
            auto extLst = root.child("extLst");
            if (extLst) {
                auto ext = extLst.child("x:ext");
                if (ext.empty()) ext = extLst.child("ext");
                if (ext) {
                    auto cacheDef = ext.child("x15:tableSlicerCache");
                    if (cacheDef.empty()) cacheDef = ext.child("tableSlicerCache");
                    if (cacheDef && cacheDef.attribute("tableId").as_uint() == tableId 
                                 && cacheDef.attribute("column").as_uint() == tableColumnId) {
                        return root.attribute("name").value();
                    }
                }
            }
        }
    }

    std::string uniqueName = std::string(name);
    uint32_t suffix = 1;
    while (true) {
        bool nameExists = false;
        for (const auto& item : m_data) {
            if (item.getXmlType() == XLContentType::SlicerCache) {
                auto root = item.getXmlDocument()->document_element();
                if (std::string_view(root.attribute("name").value()) == uniqueName) {
                    nameExists = true;
                    break;
                }
            }
        }
        if (!nameExists) break;
        uniqueName = std::string(name) + std::to_string(suffix++);
    }

    createTableSlicerCache(tableId, tableColumnId, uniqueName, sourceName);
    return uniqueName;
}

std::string XLDocument::findOrCreatePivotSlicerCache(uint32_t pivotCacheId, uint32_t sheetId, std::string_view pivotTableName, std::string_view name, std::string_view sourceName)
{
    for (auto& item : m_data) {
        if (item.getXmlType() == XLContentType::SlicerCache) {
            auto root = item.getXmlDocument()->document_element();
            auto dataNode = root.child("data");
            if (dataNode) {
                auto tabular = dataNode.child("tabular");
                if (tabular && tabular.attribute("pivotCacheId").as_uint() == pivotCacheId 
                            && std::string(root.attribute("sourceName").value()) == sourceName) {
                    auto pivotTables = root.child("pivotTables");
                    if (pivotTables) {
                        bool foundPt = false;
                        for (auto pt : pivotTables.children("pivotTable")) {
                            if (std::string(pt.attribute("name").value()) == pivotTableName) {
                                foundPt = true;
                                break;
                            }
                        }
                        if (!foundPt) {
                            auto newPt = pivotTables.append_child("pivotTable");
                            newPt.append_attribute("tabId").set_value(sheetId);
                            newPt.append_attribute("name").set_value(std::string(pivotTableName).c_str());
                        }
                    }
                    return root.attribute("name").value();
                }
            }
        }
    }

    std::string uniqueName = std::string(name);
    uint32_t suffix = 1;
    while (true) {
        bool nameExists = false;
        for (const auto& item : m_data) {
            if (item.getXmlType() == XLContentType::SlicerCache) {
                auto root = item.getXmlDocument()->document_element();
                if (std::string_view(root.attribute("name").value()) == uniqueName) {
                    nameExists = true;
                    break;
                }
            }
        }
        if (!nameExists) break;
        uniqueName = std::string(name) + std::to_string(suffix++);
    }

    createPivotSlicerCache(pivotCacheId, sheetId, pivotTableName, uniqueName, sourceName);
    return uniqueName;
}

void XLDocument::deleteSlicerFileAndOrphanCache(const std::string& name)
{
    std::string slicerPath;
    std::string cacheName;
    XLXmlData* targetSlicerXmlData = nullptr;
    XMLNode targetSlicerNode;

    for (const auto& item : m_contentTypes.getContentItems()) {
        if (item.type() == XLContentType::Slicer) {
            std::string relPath = item.path();
            if (relPath[0] == '/') relPath = relPath.substr(1);
            XLXmlData* xmlData = getXmlData(relPath, true);
            if (!xmlData) {
                xmlData = &m_data.emplace_back(this, relPath, "", XLContentType::Slicer);
            }
            if (xmlData) {
                auto root = xmlData->getXmlDocument()->document_element();
                for (auto slicerNode : root.children("slicer")) {
                    if (std::string(slicerNode.attribute("name").value()) == name) {
                        slicerPath = xmlData->getXmlPath();
                        cacheName = slicerNode.attribute("cache").value();
                        targetSlicerXmlData = xmlData;
                        targetSlicerNode = slicerNode;
                        break;
                    }
                }
                if (targetSlicerXmlData) break;
            }
        }
    }

    if (!targetSlicerXmlData) {
        return;
    }

    // 1. Delete slicer XML node from its parent
    auto root = targetSlicerXmlData->getXmlDocument()->document_element();
    root.remove_child(targetSlicerNode);

    // 2. Check if the slicer XML document has no more <slicer> nodes left
    bool fileEmpty = root.child("slicer").empty();
    if (fileEmpty) {
        m_archive.deleteEntry(slicerPath);
        if (m_contentTypes.hasOverride("/" + slicerPath)) {
            m_contentTypes.deleteOverride("/" + slicerPath);
        }
        m_unhandledEntries.erase(slicerPath);
        m_data.remove_if([&slicerPath](const XLXmlData& d) {
            return d.getXmlPath() == slicerPath;
        });
    }

    // 3. Check if the cache is still used by other slicers
    bool cacheStillUsed = false;
    for (const auto& item : m_contentTypes.getContentItems()) {
        if (item.type() == XLContentType::Slicer) {
            std::string relPath = item.path();
            if (relPath[0] == '/') relPath = relPath.substr(1);
            if (fileEmpty && relPath == slicerPath) continue;
            XLXmlData* xmlData = getXmlData(relPath, true);
            if (!xmlData) {
                xmlData = &m_data.emplace_back(this, relPath, "", XLContentType::Slicer);
            }
            if (xmlData) {
                auto sRoot = xmlData->getXmlDocument()->document_element();
                for (auto slicerNode : sRoot.children("slicer")) {
                    if (std::string(slicerNode.attribute("cache").value()) == cacheName) {
                        cacheStillUsed = true;
                        break;
                    }
                }
                if (cacheStillUsed) break;
            }
        }
    }

    if (!cacheStillUsed && !cacheName.empty()) {
        std::string cachePath;
        for (const auto& item : m_contentTypes.getContentItems()) {
            if (item.type() == XLContentType::SlicerCache) {
                std::string relPath = item.path();
                if (relPath[0] == '/') relPath = relPath.substr(1);
                XLXmlData* xmlData = getXmlData(relPath, true);
                if (!xmlData) {
                    xmlData = &m_data.emplace_back(this, relPath, "", XLContentType::SlicerCache);
                }
                if (xmlData) {
                    auto cRoot = xmlData->getXmlDocument()->document_element();
                    if (std::string(cRoot.attribute("name").value()) == cacheName) {
                        cachePath = xmlData->getXmlPath();
                        break;
                    }
                }
            }
        }

        if (!cachePath.empty()) {
            m_archive.deleteEntry(cachePath);
            if (m_contentTypes.hasOverride("/" + cachePath)) {
                m_contentTypes.deleteOverride("/" + cachePath);
            }
            m_unhandledEntries.erase(cachePath);
            
            std::string relId;
            auto rels = workbookRelationships().relationships();
            for (const auto& rel : rels) {
                std::string target = rel.target();
                if (!target.empty() && target[0] == '/') target = target.substr(1);
                if (target.rfind("xl/", 0) == 0) target = target.substr(3);

                std::string cleanCachePath = cachePath;
                if (cleanCachePath.rfind("xl/", 0) == 0) cleanCachePath = cleanCachePath.substr(3);

                if (target == cleanCachePath) {
                    relId = rel.id();
                    workbookRelationships().deleteRelationship(rel);
                    break;
                }
            }

            m_data.remove_if([&cachePath](const XLXmlData& d) {
                return d.getXmlPath() == cachePath;
            });

            if (!relId.empty()) {
                auto wbkNode = m_workbook.xmlDocument().document_element();
                auto extLst = wbkNode.child("extLst");
                if (extLst) {
                    for (auto ext : extLst.children("ext")) {
                        auto slicerCaches = ext.child("x15:slicerCaches");
                        if (slicerCaches.empty()) {
                            slicerCaches = ext.child("x14:slicerCaches");
                        }
                        if (slicerCaches) {
                            auto cacheNode = slicerCaches.find_child_by_attribute("r:id", relId.c_str());
                            if (cacheNode) {
                                slicerCaches.remove_child(cacheNode);
                            }
                            if (slicerCaches.first_child().empty()) {
                                ext.remove_child(slicerCaches);
                            }
                        }
                    }
                }
            }

            auto wbkNode = m_workbook.xmlDocument().document_element();
            auto definedNames = wbkNode.child("definedNames");
            if (definedNames) {
                auto defNameNode = definedNames.find_child_by_attribute("name", cacheName.c_str());
                if (defNameNode) {
                    definedNames.remove_child(defNameNode);
                }
                if (definedNames.first_child().empty()) {
                    wbkNode.remove_child(definedNames);
                }
            }
        }
    }
}


std::string XLDocument::createSlicer(std::string_view name,
                                     std::string_view cacheName,
                                     std::string_view caption,
                                     std::string_view existingSlicerFile)
{
    // ── Per-sheet slicer file sharing strategy ────────────────────────────────
    // OOXML requires: one slicer.xml per worksheet, containing ALL slicers on
    // that sheet as multiple <slicer> child nodes. (Reference: excelize addSlicer)
    // If an existing slicer file for this sheet is provided, append to it.
    // ─────────────────────────────────────────────────────────────────────────

    std::string filename;
    if (!existingSlicerFile.empty()) {
        // Reuse the existing file — just append a new <slicer> node below
        filename = std::string(existingSlicerFile);
    } else {
        // Create a new slicer file for this sheet
        uint32_t num = 1;
        filename     = fmt::format("xl/slicers/slicer{}.xml", num);
        while (m_archive.hasEntry(filename)) {
            ++num;
            filename = fmt::format("xl/slicers/slicer{}.xml", num);
        }

        std::string templateStr = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x xr10" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"></slicers>)";

        m_archive.addEntry(filename, templateStr);
        m_contentTypes.addOverride("/" + filename, XLContentType::Slicer);

        constexpr bool DO_NOT_THROW = true;
        XLXmlData*     xmlData      = getXmlData(filename, DO_NOT_THROW);
        if (xmlData == nullptr) {
            xmlData = &m_data.emplace_back(this, filename, templateStr, XLContentType::Slicer);
        }
    }

    // Count existing slicers workbook-wide to generate a unique xr10:uid index
    uint32_t slicerIdx = 1;
    for (auto& item : m_data) {
        if (item.getXmlType() == XLContentType::Slicer) {
            auto docPtr = item.getXmlDocument();
            if (docPtr) {
                auto root = docPtr->document_element();
                for (auto child = root.first_child_of_type(pugi::node_element); !child.empty();
                     child = child.next_sibling_of_type(pugi::node_element)) {
                    ++slicerIdx;
                }
            }
        }
    }

    XLXmlData* slicerXml = getXmlData(filename, /*doNotThrow=*/true);
    if (slicerXml) {
        XMLNode root = slicerXml->getXmlDocument()->document_element();
        // Attribute order must match Excel: name, xr10:uid, cache, caption, rowHeight
        // xr10:uid is REQUIRED by OOXML — its absence causes Excel to report corruption
        std::string uid = fmt::format("{{00000000-0014-0000-FFFF-FFFF{:02d}000000}}", slicerIdx);
        XMLNode node = root.append_child("slicer");
        node.append_attribute("name").set_value(std::string(name).c_str());
        node.append_attribute("xr10:uid").set_value(uid.c_str());
        node.append_attribute("cache").set_value(std::string(cacheName).c_str());
        node.append_attribute("caption").set_value(std::string(caption).c_str());
        node.append_attribute("rowHeight").set_value("251883");
    }

    return filename;
}


XLPivotCacheRecords XLDocument::createPivotCacheRecords(std::string_view cacheDefPath)
{
    using namespace std::literals::string_literals;

    std::string defPathStr(cacheDefPath);
    size_t      start      = defPathStr.find("Definition") + 10;
    size_t      end        = defPathStr.find(".xml");
    std::string cacheNoStr = defPathStr.substr(start, end - start);

    std::string filename = "xl/pivotCache/pivotCacheRecords"s + cacheNoStr + ".xml"s;

    m_archive.addEntry(filename, std::string(xlPivotCacheRecordsTemplate));
    m_contentTypes.addOverride("/" + filename, XLContentType::PivotCacheRecords);

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(filename, DO_NOT_THROW);
    if (xmlData == nullptr) {
        xmlData = &m_data.emplace_back(this, filename, std::string(xlPivotCacheRecordsTemplate), XLContentType::PivotCacheRecords);
    }

    return XLPivotCacheRecords(xmlData);
}


// ===== External Includes ===== //
#include <list>
#include <string>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLRelationships.hpp"
#include "XLXmlData.hpp"

using namespace OpenXLSX;

namespace
{
    bool pathInManagedParts(const std::list<XLXmlData>& data, std::string_view path)
    {
        for (const auto& item : data) {
            if (item.getXmlPath() == path) return true;
        }
        return false;
    }

    bool contentTypesHasOverride(XLContentTypes& ct, std::string_view partName)
    {
        return !ct.contentItem(partName).path().empty();
    }

    /** Normalize workbook relationship Target to package-relative path (no leading slash). */
    std::string normalizePackagePath(std::string target)
    {
        if (target.empty()) return target;
        if (target.front() == '/') target.erase(0, 1);
        // Workbook rels targets are relative to xl/
        if (target.rfind("worksheets/", 0) == 0 || target.rfind("chartsheets/", 0) == 0 || target.rfind("theme/", 0) == 0 ||
            target.rfind("styles", 0) == 0 || target.rfind("sharedStrings", 0) == 0 || target.rfind("tables/", 0) == 0 ||
            target.rfind("drawings/", 0) == 0 || target.rfind("charts/", 0) == 0 || target.rfind("pivotTables/", 0) == 0 ||
            target.rfind("pivotCache/", 0) == 0)
        {
            target = "xl/" + target;
        }
        return target;
    }
}    // namespace

/**
 * @details Hard checks that catch structural corruption before ZIP commit.
 */
void XLDocument::validatePackageInvariants() const
{
    using namespace std::literals::string_literals;

    // const method: Content_Types helpers are non-const but do not mutate when only querying.
    auto& contentTypes = const_cast<XLContentTypes&>(m_contentTypes);
    auto& wbkRels      = const_cast<XLRelationships&>(m_wbkRelationships);

    // ----- Required OPC roots -----
    if (!m_contentTypes.valid()) {
        throw XLException("Package invariant: [Content_Types].xml is not loaded");
    }

    const bool hasWorkbookPart = pathInManagedParts(m_data, "xl/workbook.xml");
    if (!hasWorkbookPart) {
        throw XLException("Package invariant: missing workbook part (xl/workbook.xml)");
    }

    if (!contentTypesHasOverride(contentTypes, "/xl/workbook.xml")) {
        // Some packages use macro-enabled content type path still named workbook.xml
        bool found = false;
        for (const auto& item : contentTypes.getContentItems()) {
            const auto p = item.path();
            if (p == "/xl/workbook.xml" || p.find("workbook") != std::string::npos) {
                found = true;
                break;
            }
        }
        if (!found) {
            throw XLException("Package invariant: [Content_Types].xml has no Override for workbook");
        }
    }

    // Styles: if Content_Types claims styles, the part must exist
    if (contentTypesHasOverride(contentTypes, "/xl/styles.xml")) {
        if (!pathInManagedParts(m_data, "xl/styles.xml") && !m_archive.hasEntry("xl/styles.xml")) {
            throw XLException("Package invariant: Content_Types references /xl/styles.xml but part is missing");
        }
    }

    // ----- Sheet relationships: every <sheet r:id> must resolve -----
    const XLXmlData* wbkData = nullptr;
    for (const auto& item : m_data) {
        if (item.getXmlPath() == "xl/workbook.xml") {
            wbkData = &item;
            break;
        }
    }
    if (!wbkData) return;

    auto wbkDoc = wbkData->getXmlDocument();
    if (!wbkDoc) return;
    XMLNode sheets = wbkDoc->document_element().child("sheets");
    if (sheets.empty()) return;

    if (!m_wbkRelationships.valid()) {
        if (!sheets.first_child().empty()) {
            throw XLException("Package invariant: workbook has sheets but xl/_rels/workbook.xml.rels is missing");
        }
        return;
    }

    for (XMLNode sheet = sheets.child("sheet"); !sheet.empty(); sheet = sheet.next_sibling("sheet")) {
        const std::string rId = sheet.attribute("r:id").value();
        if (rId.empty()) {
            throw XLException("Package invariant: workbook sheet '"s + sheet.attribute("name").value() + "' has empty r:id");
        }

        XLRelationshipItem rel = wbkRels.relationshipById(rId);
        if (rel.empty()) {
            throw XLException("Package invariant: workbook sheet r:id '"s + rId + "' not found in workbook relationships");
        }

        const std::string target = normalizePackagePath(rel.target());
        if (target.empty()) {
            throw XLException("Package invariant: workbook relationship '"s + rId + "' has empty Target");
        }

        const bool inData    = pathInManagedParts(m_data, target);
        const bool inArchive = m_archive.hasEntry(target);
        if (!inData && !inArchive) {
            throw XLException("Package invariant: sheet relationship '"s + rId + "' target '" + target +
                              "' not found in package");
        }
    }
}

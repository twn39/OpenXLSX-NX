#include <algorithm>
#include <vector>
#include <string>
#include <fmt/format.h>

#include "XLDocument.hpp"
#include "XLSheet.hpp"
#include "XLException.hpp"
#include "XLContentTypes.hpp"
#include "XLRelationships.hpp"

constexpr bool THROW_ON_INVALID = true;

using namespace OpenXLSX;

/**
 * @details Provides a centralized mutation interface for document-level changes. Abstracted as commands to decouple the UI/API layers from the internal DOM routing and dependency management logic.
 */
bool XLDocument::execCommand(const XLCommand& command)
{
    switch (command.type()) {
        case XLCommandType::SetSheetName:
            validateSheetName(command.getParam<std::string>("newName"), THROW_ON_INVALID);
            m_appProperties.setSheetName(command.getParam<std::string>("sheetName"), command.getParam<std::string>("newName"));
            m_workbook.setSheetName(command.getParam<std::string>("sheetID"), command.getParam<std::string>("newName"));
            break;

        case XLCommandType::SetSheetVisibility:
            m_workbook.setSheetVisibility(command.getParam<std::string>("sheetID"), command.getParam<std::string>("sheetVisibility"));
            break;

        case XLCommandType::SetSheetIndex: {
            XLQuery    qry(XLQueryType::QuerySheetName);
            const auto sheetName = execQuery(qry.setParam("sheetID", command.getParam<std::string>("sheetID"))).result<std::string>();
            m_workbook.setSheetIndex(sheetName, command.getParam<uint16_t>("sheetIndex"));
        } break;

        case XLCommandType::SetSheetActive:
            return m_workbook.setSheetActive(command.getParam<std::string>("sheetID"));
            // break;

        case XLCommandType::ResetCalcChain: {
            m_archive.deleteEntry("xl/calcChain.xml");
            const auto item = std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& theItem) {
                return theItem.getXmlPath() == "xl/calcChain.xml";
            });

            if (item != m_data.end()) m_data.erase(item);
        } break;
        case XLCommandType::SetFullCalcOnLoad: {
            m_workbook.setFullCalculationOnLoad();
        } break;
        case XLCommandType::CheckAndFixCoreProperties: {    // does nothing if core properties are in good shape
            // ===== If _rels/.rels has no entry for docProps/core.xml
            if (!m_docRelationships.targetExists("docProps/core.xml"))
                m_docRelationships.addRelationship(XLRelationshipType::CoreProperties, "docProps/core.xml");    // Fix m_docRelationships

            // ===== If docProps/core.xml is missing
            if (!m_archive.hasEntry("docProps/core.xml"))
                m_archive.addEntry("docProps/core.xml",
                                   "<?xml version=\"1.0\" encoding=\"UTF-8\"standalone=\"yes\"?>");    // create empty docProps/core.xml
            // ===== XLProperties constructor will take care of adding template content

            // ===== If [Content Types].xml has no relationship for docProps/core.xml
            if (!hasXmlData("docProps/core.xml")) {
                m_contentTypes.addOverride("/docProps/core.xml", XLContentType::CoreProperties);    // add content types entry
                m_data.emplace_back(                                                                // store new entry in m_data
                    /* parentDoc */ this,
                    /* xmlPath   */ "docProps/core.xml",
                    /* xmlID     */ m_docRelationships.relationshipByTarget("docProps/core.xml").id(),
                    /* xmlType   */ XLContentType::CoreProperties);
            }
        } break;
        case XLCommandType::CheckAndFixExtendedProperties: {    // does nothing if extended properties are in good shape
            // ===== If _rels/.rels has no entry for docProps/app.xml
            if (!m_docRelationships.targetExists("docProps/app.xml"))
                m_docRelationships.addRelationship(XLRelationshipType::ExtendedProperties, "docProps/app.xml");    // Fix m_docRelationships

            // ===== If docProps/app.xml is missing
            if (!m_archive.hasEntry("docProps/app.xml"))
                m_archive.addEntry("docProps/app.xml",
                                   "<?xml version=\"1.0\" encoding=\"UTF-8\"standalone=\"yes\"?>");    // create empty docProps/app.xml
            // ===== XLAppProperties constructor will take care of adding template content

            // ===== If [Content Types].xml has no relationship for docProps/app.xml
            if (!hasXmlData("docProps/app.xml")) {
                m_contentTypes.addOverride("/docProps/app.xml", XLContentType::ExtendedProperties);    // add content types entry
                m_data.emplace_back(                                                                   // store new entry in m_data
                    /* parentDoc */ this,
                    /* xmlPath   */ "docProps/app.xml",
                    /* xmlID     */ m_docRelationships.relationshipByTarget("docProps/app.xml").id(),
                    /* xmlType   */ XLContentType::ExtendedProperties);
            }
        } break;
        case XLCommandType::CheckAndFixCustomProperties: {    // does nothing if custom properties are in good shape
            // ===== If _rels/.rels has no entry for docProps/custom.xml
            if (!m_docRelationships.targetExists("docProps/custom.xml"))
                m_docRelationships.addRelationship(XLRelationshipType::CustomProperties, "docProps/custom.xml");

            // ===== If docProps/custom.xml is missing
            if (!m_archive.hasEntry("docProps/custom.xml"))
                m_archive.addEntry("docProps/custom.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");

            // ===== If [Content Types].xml has no relationship for docProps/custom.xml
            if (!hasXmlData("docProps/custom.xml")) {
                m_contentTypes.addOverride("/docProps/custom.xml", XLContentType::CustomProperties);
                m_data.emplace_back(this,
                                    "docProps/custom.xml",
                                    m_docRelationships.relationshipByTarget("docProps/custom.xml").id(),
                                    XLContentType::CustomProperties);
            }
            m_customProperties = XLCustomProperties(getXmlData("docProps/custom.xml"));
        } break;
        case XLCommandType::AddSharedStrings: {
            m_contentTypes.addOverride("/xl/sharedStrings.xml", XLContentType::SharedStrings);
            m_wbkRelationships.addRelationship(XLRelationshipType::SharedStrings, "sharedStrings.xml");
            // ===== Add empty archive entry for shared strings, XLSharedStrings constructor will create a default document when no document
            // element is found
            m_archive.addEntry("xl/sharedStrings.xml", "");
        } break;
        case XLCommandType::AddWorksheet: {
            validateSheetName(command.getParam<std::string>("sheetName"), THROW_ON_INVALID);
            const std::string emptyWorksheet{
                "<?xml version=\"1.0\" encoding=\"UTF-8\"standalone=\"yes\"?>\n"
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\""
                " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">"
                "<dimension ref=\"A1\"/>"
                "<sheetViews>"
                "<sheetView workbookViewId=\"0\"/>"
                "</sheetViews>"
                "<sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"16\" x14ac:dyDescent=\"0.2\"/>"
                "<sheetData/>"
                "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
                "</worksheet>"};
            m_contentTypes.addOverride(command.getParam<std::string>("sheetPath"), XLContentType::Worksheet);
            m_wbkRelationships.addRelationship(XLRelationshipType::Worksheet, command.getParam<std::string>("sheetPath").substr(4));
            m_appProperties.appendSheetName(command.getParam<std::string>("sheetName"));
            m_archive.addEntry(command.getParam<std::string>("sheetPath").substr(1), emptyWorksheet);
            m_data.emplace_back(
                /* parentDoc */ this,
                /* xmlPath   */ command.getParam<std::string>("sheetPath").substr(1),
                /* xmlID     */ m_wbkRelationships.relationshipByTarget(command.getParam<std::string>("sheetPath").substr(4)).id(),
                /* xmlType   */ XLContentType::Worksheet);
        } break;
        case XLCommandType::AddChartsheet: {
            validateSheetName(command.getParam<std::string>("sheetName"), THROW_ON_INVALID);
            const std::string emptyChartsheet{
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                "<sheetViews>"
                "<sheetView workbookViewId=\"0\" zoomToFit=\"1\"/>"
                "</sheetViews>"
                "</chartsheet>"};
            m_contentTypes.addOverride(command.getParam<std::string>("sheetPath"), XLContentType::Chartsheet);
            m_wbkRelationships.addRelationship(XLRelationshipType::Chartsheet, command.getParam<std::string>("sheetPath").substr(4));
            m_appProperties.appendSheetName(command.getParam<std::string>("sheetName"));
            m_archive.addEntry(command.getParam<std::string>("sheetPath").substr(1), emptyChartsheet);
            m_data.emplace_back(
                /* parentDoc */ this,
                /* xmlPath   */ command.getParam<std::string>("sheetPath").substr(1),
                /* xmlID     */ m_wbkRelationships.relationshipByTarget(command.getParam<std::string>("sheetPath").substr(4)).id(),
                /* xmlType   */ XLContentType::Chartsheet);
        } break;
        case XLCommandType::DeleteSheet: {
            m_appProperties.deleteSheetName(command.getParam<std::string>("sheetName"));
            std::string sheetPath = m_wbkRelationships.relationshipById(command.getParam<std::string>("sheetID")).target();
            if (sheetPath.substr(0, 4) != "/xl/") sheetPath = "/xl/" + sheetPath;    // 2024-12-15: respect absolute sheet path
            m_archive.deleteEntry(sheetPath.substr(1));
            m_contentTypes.deleteOverride(sheetPath);
            m_wbkRelationships.deleteRelationship(command.getParam<std::string>("sheetID"));
            m_data.erase(std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) {
                return item.getXmlPath() == sheetPath.substr(1);
            }));
        } break;
        case XLCommandType::CloneSheet: {
            validateSheetName(command.getParam<std::string>("cloneName"), THROW_ON_INVALID);
            const auto internalID = m_workbook.createInternalSheetID();
            const auto sheetPath  = fmt::format("/xl/worksheets/sheet{}.xml", internalID);
            if (m_workbook.sheetExists(command.getParam<std::string>("cloneName")))
                throw XLInternalError("Sheet named \"" + command.getParam<std::string>("cloneName") + "\" already exists.");

            // ===== 2024-12-15: handle absolute sheet path: ensure relative sheet path
            std::string sheetToClonePath = m_wbkRelationships.relationshipById(command.getParam<std::string>("sheetID")).target();
            if (sheetToClonePath.substr(0, 4) == "/xl/") sheetToClonePath = sheetToClonePath.substr(4);

            if (m_wbkRelationships.relationshipById(command.getParam<std::string>("sheetID")).type() == XLRelationshipType::Worksheet) {
                m_contentTypes.addOverride(sheetPath, XLContentType::Worksheet);
                m_wbkRelationships.addRelationship(XLRelationshipType::Worksheet, sheetPath.substr(4));
                m_appProperties.appendSheetName(command.getParam<std::string>("cloneName"));
                m_archive.addEntry(sheetPath.substr(1), std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& data) {
                                                            return data.getXmlPath().substr(3) ==
                                                                   sheetToClonePath;    // 2024-12-15: ensure relative sheet path
                                                        })->getRawData());
                m_data.emplace_back(
                    /* parentDoc */ this,
                    /* xmlPath   */ sheetPath.substr(1),
                    /* xmlID     */ m_wbkRelationships.relationshipByTarget(sheetPath.substr(4)).id(),
                    /* xmlType   */ XLContentType::Worksheet);
            }
            else {
                m_contentTypes.addOverride(sheetPath, XLContentType::Chartsheet);
                m_wbkRelationships.addRelationship(XLRelationshipType::Chartsheet, sheetPath.substr(4));
                m_appProperties.appendSheetName(command.getParam<std::string>("cloneName"));
                m_archive.addEntry(sheetPath.substr(1), std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& data) {
                                                            return data.getXmlPath().substr(3) ==
                                                                   sheetToClonePath;    // 2024-12-15: ensure relative sheet path
                                                        })->getRawData());
                m_data.emplace_back(
                    /* parentDoc */ this,
                    /* xmlPath   */ sheetPath.substr(1),
                    /* xmlID     */ m_wbkRelationships.relationshipByTarget(sheetPath.substr(4)).id(),
                    /* xmlType   */ XLContentType::Chartsheet);
            }

            m_workbook.prepareSheetMetadata(command.getParam<std::string>("cloneName"), internalID);
        } break;
        case XLCommandType::AddStyles: {
            m_contentTypes.addOverride("/xl/styles.xml", XLContentType::Styles);
            m_wbkRelationships.addRelationship(XLRelationshipType::Styles, "styles.xml");
            // ===== Add empty archive entry for styles, XLStyles constructor will create a default document when no document element is
            // found
            m_archive.addEntry("xl/styles.xml", "");
        } break;
    }

    return true;    // default: command claims success unless otherwise implemented in switch-clauses
}

/**
 * @details Provides a centralized read interface for querying document state. Decouples consumers from traversing the relationship graphs directly.
 */
XLQuery XLDocument::execQuery(const XLQuery& query) const
{
    switch (query.type()) {
        case XLQueryType::QuerySheetName:
            return XLQuery(query).setResult(m_workbook.sheetName(query.getParam<std::string>("sheetID")));

        case XLQueryType::QuerySheetIndex: {    // 2025-01-13: implemented query - previously no index was determined at all
            std::string queriedSheetName = m_workbook.sheetName(query.getParam<std::string>("sheetID"));
            for (uint16_t sheetIndex = 1; sheetIndex <= workbook().sheetCount(); ++sheetIndex) {
                if (workbook().sheet(sheetIndex).name() == queriedSheetName) return XLQuery(query).setResult(std::to_string(sheetIndex));
            }

            {    // if loop failed to locate queriedSheetName:
                using namespace std::literals::string_literals;
                throw XLInternalError("Could not determine a sheet index for sheet named \"" + queriedSheetName + "\"");
            }
        }
        case XLQueryType::QuerySheetVisibility:
            return XLQuery(query).setResult(m_workbook.sheetVisibility(query.getParam<std::string>("sheetID")));

        case XLQueryType::QuerySheetType: {
            const XLRelationshipType t = m_wbkRelationships.relationshipById(query.getParam<std::string>("sheetID")).type();
            if (t == XLRelationshipType::Worksheet) return XLQuery(query).setResult(XLContentType::Worksheet);
            if (t == XLRelationshipType::Chartsheet) return XLQuery(query).setResult(XLContentType::Chartsheet);
            return XLQuery(query).setResult(XLContentType::Unknown);
        }
        case XLQueryType::QuerySheetIsActive:
            return XLQuery(query).setResult(m_workbook.sheetIsActive(query.getParam<std::string>("sheetID")));

        case XLQueryType::QuerySheetID:
            return XLQuery(query).setResult(m_workbook.sheetVisibility(query.getParam<std::string>("sheetID")));

        case XLQueryType::QuerySheetRelsID:
            return XLQuery(query).setResult(
                m_wbkRelationships.relationshipByTarget(query.getParam<std::string>("sheetPath").substr(4)).id());

        case XLQueryType::QuerySheetRelsTarget:
            // ===== 2024-12-15: XLRelationshipItem::target() returns the unmodified Relationship "Target" property
            //                     - can be absolute or relative and must be handled by the caller
            //                   The only invocation as of today is in XLWorkbook::sheet(const std::string& sheetName) and handles this
            return XLQuery(query).setResult(m_wbkRelationships.relationshipById(query.getParam<std::string>("sheetID")).target());

        case XLQueryType::QuerySharedStrings:
            return XLQuery(query).setResult(m_sharedStrings);

        case XLQueryType::QueryXmlData: {
            const auto result = std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) {
                return item.getXmlPath() == query.getParam<std::string>("xmlPath");
            });
            if (result == m_data.end())
                throw XLInternalError("Path does not exist in zip archive (" + query.getParam<std::string>("xmlPath") + ")");
            return XLQuery(query).setResult(&*result);
        }
        default:
            throw XLInternalError("XLDocument::execQuery: unknown query type " + std::to_string(static_cast<uint8_t>(query.type())));
    }

    return query;    // Needed in order to suppress compiler warning
}

/**
 * @details Non-const overload for querying document state.
 */
XLQuery XLDocument::execQuery(const XLQuery& query) { return static_cast<const XLDocument&>(*this).execQuery(query); }

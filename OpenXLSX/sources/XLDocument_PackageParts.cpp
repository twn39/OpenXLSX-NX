// ===== XLDocument package-part accessors and factories ===== //
// Extracted from XLDocument.cpp: sheet relationships, drawings, charts, images,
// comments, tables, persons, and generic XML relationship helpers.

#include "XLDocument.hpp"

#include "XLChart.hpp"
#include "XLComments.hpp"
#include "XLContentTypes.hpp"
#include "XLDrawing.hpp"
#include "XLException.hpp"
#include "XLRelationships.hpp"
#include "XLTables.hpp"
#include "XLThreadedComments.hpp"
#include "XLUtilities.hpp"

#include <fmt/format.h>
#include <string>
#include <string_view>
#include <vector>

using namespace OpenXLSX;

bool XLDocument::hasSheetRelationships(uint16_t sheetXmlNo, bool isChartsheet) const
{
    using namespace std::literals::string_literals;
    if (isChartsheet) { return m_archive.hasEntry(fmt::format("xl/chartsheets/_rels/sheet{}.xml.rels", sheetXmlNo)); }
    return m_archive.hasEntry(fmt::format("xl/worksheets/_rels/sheet{}.xml.rels", sheetXmlNo));
}

/**
 * @details Probes the archive index for drawings to verify the presence of visual elements before attempting to parse them.
 */
bool XLDocument::hasSheetDrawing(uint16_t sheetXmlNo) const
{
    using namespace std::literals::string_literals;
    return m_archive.hasEntry(fmt::format("xl/drawings/drawing{}.xml", sheetXmlNo));
}

/**
 * @details Probes the archive index for legacy VML drawings (often used for comments or form controls) to avoid redundant DOM creation.
 */
bool XLDocument::hasSheetVmlDrawing(uint16_t sheetXmlNo) const
{
    using namespace std::literals::string_literals;
    return m_archive.hasEntry(fmt::format("xl/drawings/vmlDrawing{}.vml", sheetXmlNo));
}

/**
 * @details Probes the archive index for comments to conditionally parse them, optimizing load times for sheets without annotations.
 */
bool XLDocument::hasSheetComments(uint16_t sheetXmlNo) const
{
    using namespace std::literals::string_literals;
    return m_archive.hasEntry(fmt::format("xl/comments{}.xml", sheetXmlNo));
}

bool XLDocument::hasSheetThreadedComments(uint16_t sheetXmlNo) const
{
    using namespace std::literals::string_literals;
    return m_archive.hasEntry(fmt::format("xl/threadedComments/threadedComment{}.xml", sheetXmlNo));
}

/**
 * @details Probes the archive index for table definitions, delaying XML parsing until table data is explicitly requested.
 */
bool XLDocument::hasSheetTables(uint16_t sheetXmlNo) const
{
    using namespace std::literals::string_literals;
    return m_archive.hasEntry(fmt::format("xl/tables/table{}.xml", sheetXmlNo));
}

/**
 * @details Lazily resolves or bootstraps sheet-level relationships. This is necessary to construct valid linkages before adding interactive
 * or visual elements to a sheet.
 */
XLRelationships XLDocument::sheetRelationships(uint16_t sheetXmlNo, bool isChartsheet)
{
    using namespace std::literals::string_literals;
    std::string relsFilename;
    if (isChartsheet) { relsFilename = "xl/chartsheets/_rels/sheet"s + std::to_string(sheetXmlNo) + ".xml.rels"s; }
    else {
        relsFilename = "xl/worksheets/_rels/sheet"s + std::to_string(sheetXmlNo) + ".xml.rels"s;
    }

    if (!m_archive.hasEntry(relsFilename)) { m_archive.addEntry(relsFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"); }
    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(relsFilename, DO_NOT_THROW);
    if (xmlData == nullptr) xmlData = &m_data.emplace_back(this, relsFilename, "", XLContentType::Relationships);

    return XLRelationships(xmlData, relsFilename);
}

/**
 * @details Lazily resolves or bootstraps drawing-level relationships, which are required when linking images or charts within a drawing
 * surface.
 */
XLRelationships XLDocument::drawingRelationships(std::string_view drawingPath)
{
    using namespace std::literals::string_literals;
    size_t      lastSlash    = drawingPath.find_last_of('/');
    std::string folder       = std::string(drawingPath.substr(0, lastSlash));
    std::string filename     = std::string(drawingPath.substr(lastSlash + 1));
    std::string relsFilename = folder + "/_rels/"s + filename + ".rels"s;

    if (!m_archive.hasEntry(relsFilename)) { m_archive.addEntry(relsFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"); }

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(relsFilename, DO_NOT_THROW);
    if (xmlData == nullptr) xmlData = &m_data.emplace_back(this, relsFilename, "", XLContentType::Relationships);

    return XLRelationships(xmlData, relsFilename);
}

/**
 * @details Initializes or retrieves the drawing canvas for a sheet, serving as the root container for images, charts, and shapes.
 */
XLDrawing XLDocument::sheetDrawing(uint16_t sheetXmlNo)
{
    using namespace std::literals::string_literals;
    std::string drawingFilename = fmt::format("xl/drawings/drawing{}.xml", sheetXmlNo);

    if (!m_archive.hasEntry(drawingFilename)) {
        // Match Excel's actual drawing template: only xdr+a namespaces on root.
        // mc and sle namespaces are declared inline on mc:AlternateContent/mc:Choice child nodes.
        m_archive.addEntry(drawingFilename,
                           "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<xdr:wsDr "
                           "xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
                           "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
                           "></xdr:wsDr>");
        m_contentTypes.addOverride("/" + drawingFilename, XLContentType::Drawing);
    }
    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(drawingFilename, DO_NOT_THROW);
    if (xmlData == nullptr) xmlData = &m_data.emplace_back(this, drawingFilename, "", XLContentType::Drawing);

    return XLDrawing(xmlData);
}

XLDrawing XLDocument::createDrawing()
{
    using namespace std::literals::string_literals;

    uint32_t    drawingNo       = 1;
    std::string drawingFilename = fmt::format("xl/drawings/drawing{}.xml", drawingNo);
    while (m_archive.hasEntry(drawingFilename)) {
        ++drawingNo;
        drawingFilename = fmt::format("xl/drawings/drawing{}.xml", drawingNo);
    }

    m_archive.addEntry(drawingFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    m_contentTypes.addOverride("/" + drawingFilename, XLContentType::Drawing);

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(drawingFilename, DO_NOT_THROW);
    if (xmlData == nullptr) { xmlData = &m_data.emplace_back(this, drawingFilename, "", XLContentType::Drawing); }

    return XLDrawing(xmlData);
}

XLDrawing XLDocument::drawing(std::string_view path) { return XLDrawing(getXmlData(path)); }

XLChart XLDocument::createChart(XLChartType type)
{
    using namespace std::literals::string_literals;

    // Find a new chart number by looking at existing chart entries
    uint32_t    chartNo       = 1;
    std::string chartFilename = fmt::format("xl/charts/chart{}.xml", chartNo);
    while (m_archive.hasEntry(chartFilename)) {
        ++chartNo;
        chartFilename = fmt::format("xl/charts/chart{}.xml", chartNo);
    }

    m_archive.addEntry(chartFilename, "");
    m_contentTypes.addOverride("/" + chartFilename, XLContentType::Chart);

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(chartFilename, DO_NOT_THROW);
    if (xmlData == nullptr) { xmlData = &m_data.emplace_back(this, chartFilename, "", XLContentType::Chart); }

    XLChart chart(xmlData);
    chart.initXml(type);
    return chart;
}

/**
 * @details Initializes or retrieves the legacy VML drawing canvas, primarily required for anchoring cell comments to the UI.
 */
XLVmlDrawing XLDocument::sheetVmlDrawing(uint16_t sheetXmlNo)
{
    using namespace std::literals::string_literals;
    std::string vmlDrawingFilename = fmt::format("xl/drawings/vmlDrawing{}.vml", sheetXmlNo);

    if (!m_archive.hasEntry(vmlDrawingFilename)) {
        m_archive.addEntry(vmlDrawingFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\"standalone=\"yes\"?>");
        m_contentTypes.addDefault("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing");
        m_contentTypes.addOverride("/" + vmlDrawingFilename, XLContentType::VMLDrawing);
    }
    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(vmlDrawingFilename, DO_NOT_THROW);
    if (xmlData == nullptr) xmlData = &m_data.emplace_back(this, vmlDrawingFilename, "", XLContentType::VMLDrawing);

    return XLVmlDrawing(xmlData);
}

/**
 * @details Injects binary image payloads into the internal media folder and registers their MIME type, laying the groundwork for visual
 * sheet annotations.
 */
std::string XLDocument::addImage(std::string_view name, std::string_view data)
{
    using namespace std::literals::string_literals;
    std::string internalPath = "xl/media/"s + std::string(name);
    if (!m_archive.hasEntry(internalPath)) {
        m_archive.addEntry(internalPath, std::string(data));

        size_t dotPos = name.find_last_of('.');
        if (dotPos != std::string_view::npos) {
            std::string_view ext = name.substr(dotPos + 1);
            if (ext == "png")
                m_contentTypes.addDefault("png", "image/png");
            else if (ext == "jpg" or ext == "jpeg")
                m_contentTypes.addDefault(std::string(ext), "image/jpeg");
        }
    }
    return internalPath;
}

std::string XLDocument::addImage(std::string_view name, gsl::span<const uint8_t> data)
{
    using namespace std::literals::string_literals;
    std::string internalPath = "xl/media/"s + std::string(name);
    if (!m_archive.hasEntry(internalPath)) {
        m_archive.addEntry(internalPath, std::string(reinterpret_cast<const char*>(data.data()), data.size()));

        size_t dotPos = name.find_last_of('.');
        if (dotPos != std::string_view::npos) {
            std::string_view ext = name.substr(dotPos + 1);
            if (ext == "png")
                m_contentTypes.addDefault("png", "image/png");
            else if (ext == "jpg" or ext == "jpeg")
                m_contentTypes.addDefault(std::string(ext), "image/jpeg");
        }
    }
    return internalPath;
}

/**
 * @details Retrieves the raw binary payload of an embedded image, enabling external rendering or extraction of worksheet media.
 */
std::string XLDocument::getImage(std::string_view path) const
{
    if (m_archive.hasEntry(std::string(path))) { return m_archive.getEntry(std::string(path)); }
    return "";
}

/**
 * @details Bootstraps or retrieves the comment thread storage for a sheet, required to read or write cell-anchored notes.
 */
XLComments XLDocument::sheetComments(uint16_t sheetXmlNo)
{
    using namespace std::literals::string_literals;
    std::string commentsFilename = fmt::format("xl/comments{}.xml", sheetXmlNo);

    if (!m_archive.hasEntry(commentsFilename)) {
        m_archive.addEntry(commentsFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\"standalone=\"yes\"?>");
        m_contentTypes.addOverride("/" + commentsFilename, XLContentType::Comments);
    }
    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(commentsFilename, DO_NOT_THROW);
    if (xmlData == nullptr) xmlData = &m_data.emplace_back(this, commentsFilename, "", XLContentType::Comments);

    return XLComments(xmlData);
}

/**
 * @details Bootstraps or retrieves the table registry for a sheet, facilitating structured data manipulation and formatting.
 */
XLThreadedComments XLDocument::sheetThreadedComments(uint16_t sheetXmlNo)
{
    using namespace std::literals::string_literals;
    std::string commentsFilename = fmt::format("xl/threadedComments/threadedComment{}.xml", sheetXmlNo);

    if (!m_archive.hasEntry(commentsFilename)) {
        m_archive.addEntry(commentsFilename,
                           "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<ThreadedComments "
                           "xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" "
                           "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
        m_contentTypes.addOverride("/" + commentsFilename, XLContentType::ThreadedComments);
    }
    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(commentsFilename, DO_NOT_THROW);
    if (xmlData == nullptr) xmlData = &m_data.emplace_back(this, commentsFilename, "", XLContentType::ThreadedComments);

    return XLThreadedComments(xmlData);
}

XLTableCollection XLDocument::sheetTables(uint16_t sheetXmlNo)
{
    using namespace std::literals::string_literals;
    std::string tablesFilename = fmt::format("xl/tables/table{}.xml", sheetXmlNo);

    if (!m_archive.hasEntry(tablesFilename)) {
        m_archive.addEntry(tablesFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        m_contentTypes.addOverride("/" + tablesFilename, XLContentType::Table);
    }
    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(tablesFilename, DO_NOT_THROW);
    if (xmlData == nullptr) xmlData = &m_data.emplace_back(this, tablesFilename, "", XLContentType::Table);

    // This method is used for old compatibility or internal fetch.
    // For multiple tables, it's better to use XLWorksheet::tables()
    return XLTableCollection(xmlData->getXmlDocument()->document_element(), nullptr);
}

/**
 * @details Worksheet names cannot:
 *     Be blank.
 *     Contain more than 31 characters.
 *     Contain any of the following characters: / \ ? * : [ ]
 *     For example, 02/17/2016 would not be a valid worksheet name, but 02-17-2016 would work fine.
 *     Begin or end with an apostrophe ('), but they can be used in between text or numbers in a name.
 *     Be named "History". This is a reserved word Excel uses internally.
 */
/**
 * @details Worksheet names must adhere to strict Excel naming rules to ensure compatibility
 * across different spreadsheet applications and prevent OOXML schema violations.
 * Key restrictions include length (31 chars), forbidden characters (/ \ ? * : [ ]),
 * and the reserved name "History".
 */
bool XLDocument::validateSheetName(std::string_view sheetName, bool throwOnInvalid)
{
    using namespace std::literals::string_literals;
    auto reportError = [sheetName, throwOnInvalid](std::string_view err) {
        if (throwOnInvalid)
            throw XLInputError("Sheet name \""s + std::string(sheetName) + "\" violates naming rules: sheet name can not "s +
                               std::string(err));
        return false;
    };

    if (sheetName.length() > 31) return reportError("contain more than 31 characters");

    size_t pos = 0;
    while (pos < sheetName.length() && (sheetName[pos] == ' ' || sheetName[pos] == '\t')) ++pos;
    if (pos == sheetName.length()) return reportError("be blank");

    if (!sheetName.empty() && (sheetName.front() == '\'' || sheetName.back() == '\''))
        return reportError("begin or end with an apostrophe (')");

    if (sheetName == "History") return reportError("be named \"History\" (Excel reserves this word for internal use)");

    for (char c : sheetName) {
        switch (c) {
            case '/':
            case '\\':
            case '?':
            case '*':
            case ':':
            case '[':
            case ']':
                return reportError("contain any of the following characters: / \\ ? * : [ ]");
            default:;
        }
    }
    return true;
}

/**
 * @details Overrides the default XML header attributes, allowing consumers to dictate XML standalone status or encoding during
 * serialization.
 */
bool XLDocument::hasPersons() const { return m_archive.hasEntry("xl/persons/person.xml"); }

XLPersons& XLDocument::persons()
{
    if (!m_persons.valid()) {
        std::string personsFilename = "xl/persons/person.xml";
        if (!m_archive.hasEntry(personsFilename)) {
            m_archive.addEntry(personsFilename,
                               "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<personList "
                               "xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" "
                               "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
            m_contentTypes.addOverride("/" + personsFilename, XLContentType::Persons);
            m_wbkRelationships.addRelationship(XLRelationshipType::Person, "persons/person.xml");
        }
        XLXmlData* xmlData = getXmlData(personsFilename, true);
        if (xmlData == nullptr) { xmlData = &m_data.emplace_back(this, personsFilename, "", XLContentType::Persons); }
        m_persons = XLPersons(xmlData);
    }
    return m_persons;
}

void XLDocument::setSavingDeclaration(XLXmlSavingDeclaration const& savingDeclaration) { m_xmlSavingDeclaration = savingDeclaration; }

/**
 * @details Shared String Table (SST) cleanup.
 * Rationale: Over time, the SST can accumulate unused strings as cells are modified or deleted.
 * This method performs a full workbook traversal to identify strings currently in use,
 * remaps their indices to a new, compact table, and moves the string data to avoid copies.
 * Index 0 is reserved for empty strings by convention.
 */

XLRelationships XLDocument::xmlRelationships(std::string_view xmlPath)
{
    using namespace std::literals::string_literals;
    size_t      lastSlash    = xmlPath.find_last_of('/');
    std::string folder       = std::string(xmlPath.substr(0, lastSlash));
    std::string filename     = std::string(xmlPath.substr(lastSlash + 1));
    std::string relsFilename = folder + "/_rels/"s + filename + ".rels"s;

    if (!m_archive.hasEntry(relsFilename)) { m_archive.addEntry(relsFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"); }

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(relsFilename, DO_NOT_THROW);
    if (xmlData == nullptr) xmlData = &m_data.emplace_back(this, relsFilename, "", XLContentType::Relationships);

    return XLRelationships(xmlData, relsFilename);
}

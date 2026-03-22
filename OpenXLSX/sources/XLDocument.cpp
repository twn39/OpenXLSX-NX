// ===== External Includes ===== //
#include <algorithm>
#include <filesystem>
#include <gsl/gsl>
#if defined(_WIN32)
#    include <random>
#endif
#include <fstream>
#include <pugixml.hpp>
#include <sys/stat.h>    // for stat, to test if a file exists and if a file is a directory
#include <vector>        // std::vector

// ===== OpenXLSX Includes ===== //
#include "XLContentTypes.hpp"
#include "XLDocument.hpp"
#include "XLSheet.hpp"
#include "XLStyles.hpp"
#include "XLUtilities.hpp"
#include "XLChart.hpp"
#include "XLPivotTable.hpp"

using namespace OpenXLSX;
#include "XLDocument_Templates.hpp"

/**
 * @details Allows creating or manipulating an in-memory or custom-provided ZIP archive, decoupling the document from direct disk I/O.
 */
XLDocument::XLDocument(const IZipArchive& zipArchive) : m_xmlSavingDeclaration{}, m_archive(zipArchive) {}

/**
 * @details Provides an optimized initialization path that avoids string copies when reading an existing workbook from disk.
 */
XLDocument::XLDocument(std::string_view docPath, const IZipArchive& zipArchive) : m_xmlSavingDeclaration{}, m_archive(zipArchive)
{ open(docPath); }

/**
 * @details Maintains backward compatibility for existing codebases relying on std::string, ensuring smooth upgrades.
 */
XLDocument::XLDocument(const std::string& docPath, const IZipArchive& zipArchive) : m_xmlSavingDeclaration{}, m_archive(zipArchive)
{ open(std::string_view(docPath)); }

/**
 * @details Prevents resource leaks and ensures any pending archive handles or DOM structures are properly disposed of when the object goes out of scope.
 */
XLDocument::~XLDocument()
{
    if (isOpen()) close();
}

/**
 * @details Enables stdout logging for non-critical schema deviations (e.g., unknown xml relationships), which is vital for debugging compatibility issues.
 */
void XLDocument::showWarnings() { m_suppressWarnings = false; }

/**
 * @details Silences non-critical schema deviation warnings to keep the console output clean during production use.
 */
void XLDocument::suppressWarnings() { m_suppressWarnings = true; }

/**
 * @details Establishes the document's internal structure by parsing its root relationships and mandatory XML parts. This serves as the primary entry point for accessing and mutating any existing workbook.
 */
void XLDocument::open(std::string_view fileName)
{
    // Check if a document is already open. If yes, close it.
    if (m_archive.isOpen()) close();
    m_filePath = std::string(fileName);
    m_archive.open(m_filePath);

    // ===== Add and open the Relationships and [Content_Types] files for the document level.
    std::string relsFilename = "_rels/.rels";
    m_data.emplace_back(this, "[Content_Types].xml");
    m_data.emplace_back(this, relsFilename);

    m_contentTypes            = XLContentTypes(getXmlData("[Content_Types].xml"));
    m_docRelationships        = XLRelationships(getXmlData(relsFilename), relsFilename);
    std::string workbookPath  = "xl/workbook.xml";
    bool        workbookAdded = false;
    for (auto& item : m_docRelationships.relationships()) {
        if (item.type() == XLRelationshipType::Workbook) {
            workbookPath = item.target();
            if (workbookPath[0] == '/') workbookPath = workbookPath.substr(1);    // NON STANDARD FORMATS: strip leading '/'
            m_data.emplace_back(this, workbookPath, item.id(), XLContentType::Workbook);
            workbookAdded = true;
            break;
        }
    }

    // ===== Determine workbook relationships path based on workbookPath, and construct m_wbkRelationships
    size_t pos = workbookPath.find_last_of('/');
    if (pos == std::string::npos) {
        using namespace std::literals::string_literals;
        throw XLInputError("workbook path from "s + relsFilename + " has no folder name: "s + workbookPath);
    }
    std::string workbookRelsFilename = "xl/_rels/" + workbookPath.substr(pos + 1) + ".rels";
    m_data.emplace_back(this, workbookRelsFilename);
    m_wbkRelationships = XLRelationships(getXmlData(workbookRelsFilename), workbookRelsFilename);

    // ===== Create xl/styles.xml if missing
    if (m_contentTypes.contentItem("/xl/styles.xml").path().empty()) execCommand(XLCommand(XLCommandType::AddStyles));

    // ===== Create xl/sharedStrings.xml from scratch if missing
    if (m_contentTypes.contentItem("/xl/sharedStrings.xml").path().empty()) execCommand(XLCommand(XLCommandType::AddSharedStrings));

    // ===== Add remaining spreadsheet elements to the vector of XLXmlData objects.
    for (auto& item : m_contentTypes.getContentItems()) {
        // ===== 2024-07-26 BUGFIX: ignore content item entries for relationship files, that have already been read above
        if (item.path().substr(1) == relsFilename) continue;            // always ignore relsFilename - would have thrown above if not found
        if (item.path().substr(1) == workbookRelsFilename) continue;    // always ignore workbookRelsFilename - would have thrown

        // ===== Test if item is not in a known and handled subfolder (e.g. /customXml/*)
        if (size_t subdirectoryPos = item.path().substr(1).find_first_of('/'); subdirectoryPos != std::string::npos) {
            std::string subdirectory = item.path().substr(1, subdirectoryPos);
            if (subdirectory != "xl" and subdirectory != "docProps") continue;    // ignore items in unhandled subfolders
        }

        bool isWorkbookPath = (item.path().substr(1) == workbookPath);    // determine once, use thrice
        if (!isWorkbookPath and item.path().substr(0, 4) == "/xl/") {
            if ((item.path().substr(4, 7) == "comment") or (item.path().substr(4, 12) == "tables/table") ||
                (item.path().substr(4, 19) == "drawings/vmlDrawing") or (item.path().substr(4, 22) == "worksheets/_rels/sheet"))
            {
                // no-op - worksheet dependencies will be loaded on access through the worksheet
            }
            else if ((item.path().substr(4, 16) == "worksheets/sheet") or (item.path().substr(4) == "sharedStrings.xml") ||
                     (item.path().substr(4) == "styles.xml") or (item.path().substr(4, 11) == "theme/theme"))
            {
                m_data.emplace_back(/* parentDoc */ this,
                                    /* xmlPath   */ item.path().substr(1),
                                    /* xmlID     */ m_wbkRelationships.relationshipByTarget(item.path().substr(4)).id(),
                                    /* xmlType   */ item.type());
            }
            else {
                if (!m_suppressWarnings) std::cout << "ignoring currently unhandled workbook item " << item.path() << std::endl;
            }
        }
        else if (!isWorkbookPath or !workbookAdded) {    // do not re-add workbook if it was previously added via m_docRelationships
            if (isWorkbookPath) {        // if workbook is found but workbook relationship did not exist in m_docRelationships
                workbookAdded = true;    // 2024-09-30 bugfix: was set to true after checking item.path() == workbookPath, not
                                         // item.path().substr(1) as above
                std::cerr << "adding missing workbook relationship to _rels/.rels" << std::endl;
                m_docRelationships.addRelationship(XLRelationshipType::Workbook,
                                                   workbookPath);    // Pull request #185: Fix missing workbook relationship
            }
            m_data.emplace_back(/* parentDoc */ this,
                                /* xmlPath   */ item.path().substr(1),
                                /* xmlID     */ m_docRelationships.relationshipByTarget(item.path().substr(1)).id(),
                                /* xmlType   */ item.type());
        }
    }

    // ===== Cache unhandled archive entries (e.g., vbaProject.bin, ctrlProps)
    for (const auto& entryName : m_archive.entryNames()) {
        bool handled = false;
        // Ignore known directories or trailing slashes (e.g. xl/media/)
        if (!entryName.empty() && entryName.back() == '/') {
            handled = true;
        } else {
            for (const auto& item : m_data) {
                if (item.getXmlPath() == entryName) {
                    handled = true;
                    break;
                }
            }
        }
        
        if (!handled && entryName != "docProps/core.xml" && entryName != "docProps/app.xml" && entryName != "docProps/custom.xml") {
            // Wait, we need to extract from m_archive since the zip saveAs copies the original zip file, 
            // but if saveAs is called without the original zip (e.g. memory manipulation), it might not?
            // Actually XLZipArchive::save(path) handles this. So we just cache them so they can be explicitly added back if needed.
            m_unhandledEntries[entryName] = m_archive.getEntry(entryName);
        }
    }

    // ===== Read shared strings table.
    XMLDocument* sharedStrings = getXmlData("xl/sharedStrings.xml")->getXmlDocument();
    if (not sharedStrings->document_element().attribute("uniqueCount").empty())
        sharedStrings->document_element().remove_attribute(
            "uniqueCount");    // pull request #192 -> remove count & uniqueCount as they are optional
    if (not sharedStrings->document_element().attribute("count").empty())
        sharedStrings->document_element().remove_attribute(
            "count");    // pull request #192 -> remove count & uniqueCount as they are optional

    XMLNode node =
        sharedStrings->document_element().first_child_of_type(pugi::node_element);    // pull request #186: Skip non-element nodes in sst.
    while (not node.empty()) {
        // ===== Validate si node name.
        using namespace std::literals::string_literals;
        if (node.name() != "si"s) throw XLInputError("xl/sharedStrings.xml sst node name \""s + node.name() + "\" is not \"si\""s);

        // ===== 2024-09-01 Refactored code to tolerate a mix of <t> and <r> tags within a shared string entry.
        // This simplifies the loop while not doing any harm (obsolete inner loops for rich text and text elements removed).

        // ===== Find first node_element child of si node.
        XMLNode     elem = node.first_child_of_type(pugi::node_element);
        std::string result{};    // assemble a shared string entry here
        while (not elem.empty()) {
            // 2024-09-01: support a string composed of multiple <t> nodes in the same way as rich text <r> nodes, because LibreOffice
            // accepts it

            std::string elementName = elem.name();         // assign name to a string once, for string comparisons using operator==
            if (elementName == "t")                        // If elem is a regular string
                result += elem.text().get();               // append the tag value to result
            else if (elementName == "r")                   // If elem is rich text
                result += elem.child("t").text().get();    // append the <t> node value to result
            // ===== Ignore phonetic property tags
            else if (elementName == "rPh" or elementName == "phoneticPr") {}
            else    // For all other (unexpected) tags, throw an exception
                throw XLInputError("xl/sharedStrings.xml si node \""s + elementName +
                                   "\" is none of \"r\", \"t\", \"rPh\", \"phoneticPr\""s);

            elem = elem.next_sibling_of_type(pugi::node_element);    // advance to next child of <si>
        }
        // ===== Append an empty string even if elem.empty(), to keep the index aligned with the <si> tag index in the shared strings table
        // <sst>
        m_sharedStringCache.emplace_back(
            result);    // 2024-09-01 TBC BUGFIX: previously, a shared strings table entry that had neither <t> nor
        /**/            //     <r> nodes would not have appended to m_sharedStringCache, causing an index misalignment

        node = node.next_sibling_of_type(pugi::node_element);
    }

    // ===== Open the workbook and document property items
    m_workbook = XLWorkbook(getXmlData(workbookPath));
    // 2024-05-31: moved XLWorkbook object creation up in code worksheets info can be used for XLAppProperties generation from scratch

    // ===== 2024-06-03: creating core and extended properties if they do not exist
    execCommand(XLCommand(XLCommandType::CheckAndFixCoreProperties));        // checks & fixes consistency of docProps/core.xml related data
    execCommand(XLCommand(XLCommandType::CheckAndFixExtendedProperties));    // checks & fixes consistency of docProps/app.xml related data

    if (!hasXmlData("docProps/core.xml") or !hasXmlData("docProps/app.xml"))
        throw XLInternalError("Failed to repair docProps (core.xml and/or app.xml)");

    m_coreProperties = XLProperties(getXmlData("docProps/core.xml"));
    m_appProperties  = XLAppProperties(getXmlData("docProps/app.xml"), m_workbook.xmlDocument());
    if (hasXmlData("docProps/custom.xml")) m_customProperties = XLCustomProperties(getXmlData("docProps/custom.xml"));
    // ===== 2024-09-02: ensure that all worksheets are contained in app.xml <TitlesOfParts> and reflected in <HeadingPairs> value for
    // Worksheets
    m_appProperties.alignWorksheets(m_workbook.sheetNames());

    m_sharedStrings = XLSharedStrings(getXmlData("xl/sharedStrings.xml"), &m_sharedStringCache, &m_sharedStringIndex);
    m_styles = XLStyles(getXmlData("xl/styles.xml"), m_suppressWarnings);    // 2024-10-14: forward supress warnings setting to XLStyles
}

namespace
{
    /**
     * @brief Test if path exists as either a file or a directory
     * @param path Check for existence of this
     * @return true if path exists as a file or directory
     */
    bool pathExists(std::string_view path) { return std::filesystem::exists(std::filesystem::u8path(path)); }
#ifdef __GNUC__    // conditionally enable GCC specific pragmas to suppress unused function warning
#    pragma GCC diagnostic push
#    pragma GCC diagnostic ignored "-Wunused-function"
#endif    // __GNUC__
    /**
     * @brief Test if fileName exists and is not a directory
     * @param fileName The path to check for existence (as a file)
     * @return true if fileName exists and is a file, otherwise false
     */
    bool fileExists(std::string_view fileName)
    {
        return std::filesystem::exists(std::filesystem::u8path(fileName)) &&
               !std::filesystem::is_directory(std::filesystem::u8path(fileName));
    }
    bool isDirectory(std::string_view fileName)
    {
        return std::filesystem::exists(std::filesystem::u8path(fileName)) &&
               std::filesystem::is_directory(std::filesystem::u8path(fileName));
    }
#ifdef __GNUC__    // conditionally enable GCC specific pragmas to suppress unused function warning
#    pragma GCC diagnostic pop
#endif    // __GNUC__
}    // anonymous namespace

/**
 * @details Bootstraps a minimal, valid OOXML document structure from scratch. Uses hardcoded string templates instead of a binary payload to prevent carrying opaque blobs in the executable.
 */
void XLDocument::create(std::string_view fileName, bool forceOverwrite)
{
    if (!forceOverwrite and pathExists(fileName)) {
        using namespace std::literals::string_literals;
        throw XLException("XLDocument::create: refusing to overwrite existing file "s + std::string(fileName));
    }

    if (m_archive.isOpen()) close();

    if (pathExists(fileName) && forceOverwrite) { std::filesystem::remove(std::filesystem::u8path(fileName)); }

    m_filePath = std::string(fileName);
    m_archive.open(m_filePath);

    m_archive.addEntry("[Content_Types].xml", std::string(templateContentTypes));
    m_archive.addEntry("_rels/.rels", std::string(templateRootRels));
    m_archive.addEntry("xl/workbook.xml", std::string(templateWorkbook));
    m_archive.addEntry("xl/_rels/workbook.xml.rels", std::string(templateWorkbookRels));
    m_archive.addEntry("xl/worksheets/sheet1.xml", std::string(templateSheet));
    m_archive.addEntry("xl/styles.xml", std::string(templateStyles));
    m_archive.addEntry("xl/theme/theme1.xml", std::string(templateTheme));
    m_archive.addEntry("xl/sharedStrings.xml", std::string(templateSharedStrings));
    m_archive.addEntry("docProps/app.xml", std::string(templateApp));
    m_archive.addEntry("docProps/core.xml", std::string(templateCore));

    m_archive.save(std::string(fileName));
    m_archive.close();

    open(fileName);
}

/**
 * @details Maintains backward compatibility for creating new documents using std::string, ensuring smooth upgrades.
 */
void XLDocument::create(const std::string& fileName) { create(std::string_view(fileName), XLForceOverwrite); }

/**
 * @details Safely drops the internal DOM representations and archive handles, returning the XLDocument to a pristine, uninitialized state to prevent cross-contamination between files.
 */
void XLDocument::close()
{
    if (m_archive.isValid()) m_archive.close();
    m_filePath.clear();
    m_xmlSavingDeclaration = XLXmlSavingDeclaration("1.0", "UTF-8", false);
    m_data.clear();
    m_sharedStringCache.clear();
    m_sharedStringIndex.clear();
    m_sharedStrings    = XLSharedStrings();
    m_docRelationships = XLRelationships();
    m_wbkRelationships = XLRelationships();
    m_contentTypes     = XLContentTypes();
    m_appProperties    = XLAppProperties();
    m_coreProperties   = XLProperties();
    m_customProperties = XLCustomProperties();
    m_styles           = XLStyles();
    m_workbook         = XLWorkbook();
    m_sharedFormulas.clear();
}

/**
 * @details Commits the current DOM state to the archive, updating dimensions and calculation chains before disk flush to ensure Excel interprets the modified file as valid.
 */
void XLDocument::save() { saveAs(m_filePath, XLForceOverwrite); }

void XLDocument::addStreamedFile(const std::string& pathInZip, const std::string& tempFilePath) { m_archive.addEntryFromFile(pathInZip, tempFilePath); }

/**
 * @details Serializes the modified DOM objects into their respective XML strings and constructs a new ZIP archive. Caches and restores unhandled media/VBA entries to prevent data loss in macro-enabled files.
 */
void XLDocument::saveAs(std::string_view fileName, bool forceOverwrite)
{
    if (!forceOverwrite and pathExists(fileName)) {
        using namespace std::literals::string_literals;
        throw XLException("XLDocument::saveAs: refusing to overwrite existing file "s + std::string(fileName));
    }

    m_filePath = std::string(fileName);
    workbook().updateWorksheetDimensions();
    execCommand(XLCommand(XLCommandType::ResetCalcChain));

    if (m_filePath.size() >= 5 && m_filePath.substr(m_filePath.size() - 5) == ".xlsm") {
        m_contentTypes.updateOverride("/xl/workbook.xml", XLContentType::WorkbookMacroEnabled);
    } else if (m_filePath.size() >= 5 && m_filePath.substr(m_filePath.size() - 5) == ".xlsx") {
        m_contentTypes.updateOverride("/xl/workbook.xml", XLContentType::Workbook);
    }

    m_archive.addEntry("[Content_Types].xml", m_contentTypes.xmlData(m_xmlSavingDeclaration));
    m_archive.addEntry("_rels/.rels", m_docRelationships.xmlData(m_xmlSavingDeclaration));
    if (m_wbkRelationships.valid()) {
        m_archive.addEntry("xl/_rels/workbook.xml.rels", m_wbkRelationships.xmlData(m_xmlSavingDeclaration));
    }

    for (auto& item : m_data) {
        if (item.getXmlPath() == "[Content_Types].xml" or item.getXmlPath() == "_rels/.rels" ||
            item.getXmlPath() == "xl/_rels/workbook.xml.rels")
            continue;

        bool xmlIsStandalone = m_xmlSavingDeclaration.standalone_as_bool();
        if ((item.getXmlPath() == "docProps/core.xml") or (item.getXmlPath() == "docProps/app.xml") or
            (item.getXmlPath() == "docProps/custom.xml"))
            xmlIsStandalone = XLXmlStandalone;
        
        if (item.m_isStreamed) {
            m_archive.addEntryFromFile(item.getXmlPath(), item.m_streamFilePath);
        } else {
            m_archive.addEntry(
                item.getXmlPath(),
                item.getRawData(XLXmlSavingDeclaration(m_xmlSavingDeclaration.version(), m_xmlSavingDeclaration.encoding(), xmlIsStandalone)));
        }

    }

    for (const auto& entry : m_unhandledEntries) {
        if (std::none_of(m_data.begin(), m_data.end(), [&](const XLXmlData& item) { return item.getXmlPath() == entry.first; })) {
            m_archive.addEntry(entry.first, entry.second);
        }
    }

    m_archive.save(m_filePath);
}

/**
 * @details Maintains backward compatibility for saving documents using std::string, ensuring smooth upgrades.
 */
void XLDocument::saveAs(const std::string& fileName) { saveAs(std::string_view(fileName), XLForceOverwrite); }

/**
 * @details Provides the base filename to facilitate identification in UI, logging, or reporting contexts.
 */
std::string XLDocument::name() const
{
    size_t pos = m_filePath.find_last_of('/');
    if (pos != std::string::npos)
        return m_filePath.substr(pos + 1);
    else
        return m_filePath;
}

/**
 * @details Exposes the content types manager, which is required to register new XML parts (like sheets or drawings) so Excel recognizes their purpose.
 */
XLContentTypes& XLDocument::contentTypes() { return m_contentTypes; }

/**
 * @details Provides access to custom properties, enabling the storage of application-specific metadata within the OOXML container.
 */
XLCustomProperties& XLDocument::customProperties()
{
    if (!hasXmlData("docProps/custom.xml")) execCommand(XLCommand(XLCommandType::CheckAndFixCustomProperties));
    return m_customProperties;
}

/**
 * @details Returns the active file path to allow consumers to verify or display the target destination.
 */
const std::string& XLDocument::path() const { return m_filePath; }

/**
 * @details Calculates the next available auto-incrementing ID across all tables to ensure unique table references within the document, a strict OOXML requirement.
 */
uint32_t XLDocument::nextTableId() const
{
    uint32_t maxId = 0;
    for (const auto& item : m_data) {
        if (item.getXmlType() == XLContentType::Table) {
            auto id = item.getXmlDocument()->document_element().attribute("id").as_uint();
            if (id > maxId) maxId = id;
        }
    }
    return maxId + 1;
}

/**
 * @details Provides access to the core workbook component, serving as the gateway to manage worksheets, chartsheets, and global definitions.
 */
XLWorkbook XLDocument::workbook() const { return m_workbook; }


/**
 * @details Allows fast truthy checks to verify if the underlying zip archive is successfully bound and accessible.
 */
XLDocument::operator bool() const
{
    return m_archive.isValid();    // NOLINT
}

/**
 * @details Provides an explicit boolean check for the document's readiness state, preventing operations on uninitialized archives.
 */
bool XLDocument::isOpen() const { return this->operator bool(); }

/**
 * @details Provides access to the global stylesheet manager to retrieve or define fonts, fills, borders, and number formats used across the workbook.
 */
XLStyles& XLDocument::styles() { return m_styles; }

/**
 * @details Probes the archive index for sheet relationships to conditionally access dependencies, avoiding eager and expensive XML allocations for untouched components.
 */
bool XLDocument::hasSheetRelationships(uint16_t sheetXmlNo, bool isChartsheet) const
{
    using namespace std::literals::string_literals;
    if (isChartsheet) {
        return m_archive.hasEntry("xl/chartsheets/_rels/sheet"s + std::to_string(sheetXmlNo) + ".xml.rels"s);
    }
    return m_archive.hasEntry("xl/worksheets/_rels/sheet"s + std::to_string(sheetXmlNo) + ".xml.rels"s);
}

/**
 * @details Probes the archive index for drawings to verify the presence of visual elements before attempting to parse them.
 */
bool XLDocument::hasSheetDrawing(uint16_t sheetXmlNo) const
{
    using namespace std::literals::string_literals;
    return m_archive.hasEntry("xl/drawings/drawing"s + std::to_string(sheetXmlNo) + ".xml"s);
}

/**
 * @details Probes the archive index for legacy VML drawings (often used for comments or form controls) to avoid redundant DOM creation.
 */
bool XLDocument::hasSheetVmlDrawing(uint16_t sheetXmlNo) const
{
    using namespace std::literals::string_literals;
    return m_archive.hasEntry("xl/drawings/vmlDrawing"s + std::to_string(sheetXmlNo) + ".vml"s);
}

/**
 * @details Probes the archive index for comments to conditionally parse them, optimizing load times for sheets without annotations.
 */
bool XLDocument::hasSheetComments(uint16_t sheetXmlNo) const
{
    using namespace std::literals::string_literals;
    return m_archive.hasEntry("xl/comments"s + std::to_string(sheetXmlNo) + ".xml"s);
}

/**
 * @details Probes the archive index for table definitions, delaying XML parsing until table data is explicitly requested.
 */
bool XLDocument::hasSheetTables(uint16_t sheetXmlNo) const
{
    using namespace std::literals::string_literals;
    return m_archive.hasEntry("xl/tables/table"s + std::to_string(sheetXmlNo) + ".xml"s);
}

/**
 * @details Lazily resolves or bootstraps sheet-level relationships. This is necessary to construct valid linkages before adding interactive or visual elements to a sheet.
 */
XLRelationships XLDocument::sheetRelationships(uint16_t sheetXmlNo, bool isChartsheet)
{
    using namespace std::literals::string_literals;
    std::string relsFilename;
    if (isChartsheet) {
        relsFilename = "xl/chartsheets/_rels/sheet"s + std::to_string(sheetXmlNo) + ".xml.rels"s;
    } else {
        relsFilename = "xl/worksheets/_rels/sheet"s + std::to_string(sheetXmlNo) + ".xml.rels"s;
    }

    if (!m_archive.hasEntry(relsFilename)) { m_archive.addEntry(relsFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"); }
    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(relsFilename, DO_NOT_THROW);
    if (xmlData == nullptr) xmlData = &m_data.emplace_back(this, relsFilename, "", XLContentType::Relationships);

    return XLRelationships(xmlData, relsFilename);
}

/**
 * @details Lazily resolves or bootstraps drawing-level relationships, which are required when linking images or charts within a drawing surface.
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
    std::string drawingFilename = "xl/drawings/drawing"s + std::to_string(sheetXmlNo) + ".xml"s;

    if (!m_archive.hasEntry(drawingFilename)) {
        m_archive.addEntry(drawingFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\"standalone=\"yes\"?>");
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

    uint32_t drawingNo = 1;
    std::string drawingFilename = "xl/drawings/drawing"s + std::to_string(drawingNo) + ".xml"s;
    while (m_archive.hasEntry(drawingFilename)) {
        ++drawingNo;
        drawingFilename = "xl/drawings/drawing"s + std::to_string(drawingNo) + ".xml"s;
    }

    m_archive.addEntry(drawingFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    m_contentTypes.addOverride("/" + drawingFilename, XLContentType::Drawing);
    
    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(drawingFilename, DO_NOT_THROW);
    if (xmlData == nullptr) {
        xmlData = &m_data.emplace_back(this, drawingFilename, "", XLContentType::Drawing);
    }

    return XLDrawing(xmlData);
}

XLDrawing XLDocument::drawing(std::string_view path)
{
    return XLDrawing(getXmlData(path));
}

XLChart XLDocument::createChart(XLChartType type)
{
    using namespace std::literals::string_literals;

    // Find a new chart number by looking at existing chart entries
    uint32_t chartNo = 1;
    std::string chartFilename = "xl/charts/chart"s + std::to_string(chartNo) + ".xml"s;
    while (m_archive.hasEntry(chartFilename)) {
        ++chartNo;
        chartFilename = "xl/charts/chart"s + std::to_string(chartNo) + ".xml"s;
    }

    m_archive.addEntry(chartFilename, "");
    m_contentTypes.addOverride("/" + chartFilename, XLContentType::Chart);
    
    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(chartFilename, DO_NOT_THROW);
    if (xmlData == nullptr) {
        xmlData = &m_data.emplace_back(this, chartFilename, "", XLContentType::Chart);
    }

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
    std::string vmlDrawingFilename = "xl/drawings/vmlDrawing"s + std::to_string(sheetXmlNo) + ".vml"s;

    if (!m_archive.hasEntry(vmlDrawingFilename)) {
        m_archive.addEntry(vmlDrawingFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\"standalone=\"yes\"?>");
        m_contentTypes.addOverride("/" + vmlDrawingFilename, XLContentType::VMLDrawing);
    }
    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(vmlDrawingFilename, DO_NOT_THROW);
    if (xmlData == nullptr) xmlData = &m_data.emplace_back(this, vmlDrawingFilename, "", XLContentType::VMLDrawing);

    return XLVmlDrawing(xmlData);
}

/**
 * @details Injects binary image payloads into the internal media folder and registers their MIME type, laying the groundwork for visual sheet annotations.
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
    std::string commentsFilename = "xl/comments"s + std::to_string(sheetXmlNo) + ".xml"s;

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
XLTableCollection XLDocument::sheetTables(uint16_t sheetXmlNo)
{
    using namespace std::literals::string_literals;
    std::string tablesFilename = "xl/tables/table"s + std::to_string(sheetXmlNo) + ".xml"s;

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
bool           XLDocument::validateSheetName(std::string_view sheetName, bool throwOnInvalid)
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
 * @details Overrides the default XML header attributes, allowing consumers to dictate XML standalone status or encoding during serialization.
 */
void XLDocument::setSavingDeclaration(XLXmlSavingDeclaration const& savingDeclaration) { m_xmlSavingDeclaration = savingDeclaration; }

/**
 * @details Shared String Table (SST) cleanup.
 * Rationale: Over time, the SST can accumulate unused strings as cells are modified or deleted.
 * This method performs a full workbook traversal to identify strings currently in use,
 * remaps their indices to a new, compact table, and moves the string data to avoid copies.
 * Index 0 is reserved for empty strings by convention.
 */
void XLDocument::cleanupSharedStrings()
{
    const size_t         oldStringCount = m_sharedStringCache.size();
    std::vector<int32_t> indexMap(oldStringCount, -1);
    int32_t              newStringCount = 1;

    for (uint16_t wIndex = 1; wIndex <= m_workbook.worksheetCount(); ++wIndex) {
        XLWorksheet wks       = m_workbook.worksheet(wIndex);
        XLCellRange cellRange = wks.range();
        for (auto& cell : cellRange) {
            if (cell.value().type() == XLValueType::String) {
                XLCellValueProxy val = cell.value();
                const int32_t    si  = val.stringIndex();
                if (si < 0) continue;
                if (indexMap[static_cast<size_t>(si)] == -1) {
                    if (!m_sharedStringCache[static_cast<size_t>(si)].empty())
                        indexMap[static_cast<size_t>(si)] = newStringCount++;
                    else
                        indexMap[static_cast<size_t>(si)] = 0;
                }
                if (indexMap[static_cast<size_t>(si)] != si) val.setStringIndex(indexMap[static_cast<size_t>(si)]);
            }
        }
    }

    std::vector<std::string> newStringCache(static_cast<size_t>(newStringCount));
    newStringCache[0] = "";
    for (size_t oldIdx = 0; oldIdx < oldStringCount; ++oldIdx) {
        if (const int32_t newIdx = indexMap[oldIdx]; newIdx > 0)
            newStringCache[static_cast<size_t>(newIdx)] = std::move(m_sharedStringCache[oldIdx]);
    }

    m_sharedStringCache.clear();
    std::move(newStringCache.begin(), newStringCache.end(), std::back_inserter(m_sharedStringCache));
    if (static_cast<int32_t>(m_sharedStringCache.size()) != m_sharedStrings.rewriteXmlFromCache())
        throw XLInternalError("XLDocument::cleanupSharedStrings: failed to rewrite shared string table - document would be corrupted");
}

//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

/**
 * @details Serves as the foundational bridge between the raw ZIP layer and the parsed DOM layer, safely extracting content only when explicitly requested.
 */
std::string XLDocument::extractXmlFromArchive(std::string_view path)
{ return m_archive.hasEntry(std::string(path)) ? m_archive.getEntry(std::string(path)) : ""; }

/**
 * @details Returns a mutable pointer to the requested XML part, acting as the document's central DOM registry cache.
 */
XLXmlData* XLDocument::getXmlData(std::string_view path, bool doNotThrow)
{ return const_cast<XLXmlData*>(const_cast<XLDocument const*>(this)->getXmlData(path, doNotThrow)); }

/**
 * @details Returns a read-only pointer to the requested XML part from the DOM cache.
 */
const XLXmlData* XLDocument::getXmlData(std::string_view path, bool doNotThrow) const
{
    auto result = std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) { return item.getXmlPath() == path; });
    if (result == m_data.end()) {
        if (doNotThrow) return nullptr;
        throw XLInternalError("Path " + std::string(path) + " does not exist in zip archive.");
    }
    return &*result;
}

/**
 * @details Provides a safe lookup into the DOM cache to verify the existence of optional parts without triggering exceptions.
 */
bool XLDocument::hasXmlData(std::string_view path) const
{
    return std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) { return item.getXmlPath() == path; }) != m_data.end();
}




XLPivotTable XLDocument::createPivotTable()
{
    using namespace std::literals::string_literals;

    uint32_t num = 1;
    std::string filename = "xl/pivotTables/pivotTable"s + std::to_string(num) + ".xml"s;
    while (m_archive.hasEntry(filename)) {
        ++num;
        filename = "xl/pivotTables/pivotTable"s + std::to_string(num) + ".xml"s;
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

    uint32_t num = 1;
    std::string filename = "xl/pivotCache/pivotCacheDefinition"s + std::to_string(num) + ".xml"s;
    while (m_archive.hasEntry(filename)) {
        ++num;
        filename = "xl/pivotCache/pivotCacheDefinition"s + std::to_string(num) + ".xml"s;
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

XLPivotCacheRecords XLDocument::createPivotCacheRecords(std::string_view cacheDefPath)
{
    using namespace std::literals::string_literals;
    
    std::string defPathStr(cacheDefPath);
    size_t start = defPathStr.find("Definition") + 10;
    size_t end = defPathStr.find(".xml");
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

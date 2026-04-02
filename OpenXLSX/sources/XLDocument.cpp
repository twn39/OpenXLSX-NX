// ===== External Includes ===== //
#include <algorithm>
#include <filesystem>
#include <gsl/gsl>
#if defined(_WIN32)
#    include <random>
#endif
#include <fmt/format.h>
#include <fstream>
#include <mutex>
#include <pugixml.hpp>
#include <shared_mutex>
#include <sys/stat.h>    // for stat, to test if a file exists and if a file is a directory
#include <vector>        // std::vector

// ===== OpenXLSX Includes ===== //
#include "XLChart.hpp"
#include "XLContentTypes.hpp"
#include "XLDocument.hpp"

#include "XLCrypto.hpp"
#include <fstream>
#include <filesystem>
#include <random>

#include "XLPivotTable.hpp"
#include "XLSheet.hpp"
#include "XLStyles.hpp"
#include "XLUtilities.hpp"

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
 * @details Prevents resource leaks and ensures any pending archive handles or DOM structures are properly disposed of when the object goes
 * out of scope.
 */
XLDocument::~XLDocument()
{
    if (isOpen()) close();
}

/**
 * @details Enables stdout logging for non-critical schema deviations (e.g., unknown xml relationships), which is vital for debugging
 * compatibility issues.
 */
void XLDocument::showWarnings() { m_suppressWarnings = false; }

/**
 * @details Silences non-critical schema deviation warnings to keep the console output clean during production use.
 */
void XLDocument::suppressWarnings() { m_suppressWarnings = true; }

/**
 * @details Establishes the document's internal structure by parsing its root relationships and mandatory XML parts. This serves as the
 * primary entry point for accessing and mutating any existing workbook.
 */




void XLDocument::open(std::string_view fileName, const std::string& password)
{
    if (m_archive.isOpen()) close();

    std::ifstream file(std::string(fileName), std::ios::binary | std::ios::ate);
    if (!file) throw XLInternalError("Failed to open encrypted document");

    auto size = file.tellg();
    file.seekg(0, std::ios::beg);
    std::vector<uint8_t> buffer(size);
    if (!file.read(reinterpret_cast<char*>(buffer.data()), size)) {
        throw XLInternalError("Failed to read encrypted document");
    }

    if (!isEncryptedDocument(buffer)) {
        open(fileName);
        return;
    }

    auto decrypted = decryptDocument(buffer, password);
    if (decrypted.empty()) {
        throw XLInternalError("Decryption failed or not implemented");
    }

    std::random_device rd;
    auto tempPath = std::filesystem::temp_directory_path() / ("openxlsx_" + std::to_string(rd()) + ".xlsx");
    std::ofstream out(tempPath, std::ios::binary);
    out.write(reinterpret_cast<const char*>(decrypted.data()), decrypted.size());
    out.close();

    open(tempPath.string());
    
    m_isEncryptedSession = true;
    m_encryptionPassword = password;
    m_tempDecryptedPath = tempPath.string();
    m_filePath = std::string(fileName); 
}

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
            if ((item.path().substr(4, 7) == "comment") or (item.path().substr(4, 15) == "threadedComment") or (item.path().substr(4, 12) == "tables/table") ||
                (item.path().substr(4, 19) == "drawings/vmlDrawing") or (item.path().substr(4, 22) == "worksheets/_rels/sheet"))
            {
                // no-op - worksheet dependencies will be loaded on access through the worksheet
            }
            else if ((item.path().substr(4, 16) == "worksheets/sheet") or (item.path().substr(4) == "sharedStrings.xml") ||
                     (item.path().substr(4) == "styles.xml") or (item.path().substr(4, 11) == "theme/theme") ||
                     (item.path().substr(4, 14) == "persons/person"))
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
        if (!entryName.empty() && entryName.back() == '/') { handled = true; }
        else {
            for (const auto& item : m_data) {
                if (item.getXmlPath() == entryName) {
                    handled = true;
                    break;
                }
            }
        }

        if (!handled && entryName != "docProps/core.xml" && entryName != "docProps/app.xml" && entryName != "docProps/custom.xml" &&
            entryName != "[Content_Types].xml" && entryName != "_rels/.rels" && entryName != "xl/_rels/workbook.xml.rels")
        {
            // Wait, we need to extract from m_archive since the zip saveAs copies the original zip file,
            // but if saveAs is called without the original zip (e.g. memory manipulation), it might not?
            // Actually XLZipArchive::save(path) handles this. So we just cache them so they can be explicitly added back if needed.
            m_unhandledEntries[entryName] = m_archive.getEntry(entryName);
        }
    }

    // ===== Read shared strings table (Safely bypass if not present)
    XLXmlData*   sstData       = getXmlData("xl/sharedStrings.xml", true);
    XMLDocument* sharedStrings = sstData ? sstData->getXmlDocument() : nullptr;
    if (sharedStrings && sharedStrings->document_element()) {
        if (not sharedStrings->document_element().attribute("uniqueCount").empty())
            sharedStrings->document_element().remove_attribute("uniqueCount");
        if (not sharedStrings->document_element().attribute("count").empty()) sharedStrings->document_element().remove_attribute("count");
    }

    XMLNode node = (sharedStrings && sharedStrings->document_element())
                       ? sharedStrings->document_element().first_child_of_type(pugi::node_element)
                       : XMLNode();
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
        m_sharedStringCache.emplace_back(m_sharedStringArena.store(result));    // store result string in arena and get a persistent view
        // 2024-09-01 TBC BUGFIX: previously, a shared strings table entry that had neither <t> nor
        /**/    //     <r> nodes would not have appended to m_sharedStringCache, causing an index misalignment

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

    if (getXmlData("xl/sharedStrings.xml", true)) {
        m_sharedStrings = XLSharedStrings(getXmlData("xl/sharedStrings.xml"),
                                          &m_sharedStringArena,
                                          &m_sharedStringCache,
                                          &m_sharedStringIndex,
                                          m_sharedStringMutex.get());
    }
    else {
        m_sharedStrings = XLSharedStrings();
    }

    if (getXmlData("xl/persons/person.xml", true)) {
        m_persons = XLPersons(getXmlData("xl/persons/person.xml"));
    }

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
 * @details Bootstraps a minimal, valid OOXML document structure from scratch. Uses hardcoded string templates instead of a binary payload
 * to prevent carrying opaque blobs in the executable.
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
 * @details Safely drops the internal DOM representations and archive handles, returning the XLDocument to a pristine, uninitialized state
 * to prevent cross-contamination between files.
 */
void XLDocument::close()
{
    if (m_archive.isValid()) m_archive.close();
    m_filePath.clear();
    m_xmlSavingDeclaration = XLXmlSavingDeclaration("1.0", "UTF-8", false);

    // Clean up temporary streaming files
    for (auto& item : m_data) {
        if (item.m_isStreamed && !item.m_streamFilePath.empty()) {
            std::error_code ec;
            if (std::filesystem::exists(item.m_streamFilePath, ec)) { std::filesystem::remove(item.m_streamFilePath, ec); }
            item.m_isStreamed = false;
            item.m_streamFilePath.clear();
        }
    }

    m_data.clear();
    m_sharedStringArena.clear();
    m_sharedStringCache.clear();
    m_sharedStringIndex.clear();
    m_sharedStrings    = XLSharedStrings();
    m_persons          = XLPersons();
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
 * @details Commits the current DOM state to the archive, updating dimensions and calculation chains before disk flush to ensure Excel
 * interprets the modified file as valid.
 */
void XLDocument::save() { saveAs(m_filePath, XLForceOverwrite); }

void XLDocument::addStreamedFile(std::string_view pathInZip, std::string_view tempFilePath)
{ m_archive.addEntryFromFile(pathInZip, tempFilePath); }

/**
 * @details Serializes the modified DOM objects into their respective XML strings and constructs a new ZIP archive. Caches and restores
 * unhandled media/VBA entries to prevent data loss in macro-enabled files.
 */
void XLDocument::saveAs(std::string_view fileName, const std::string& password, bool forceOverwrite)
{
    m_isEncryptedSession = true;
    m_encryptionPassword = password;
    if (m_tempDecryptedPath.empty()) {
        std::random_device rd;
        m_tempDecryptedPath = (std::filesystem::temp_directory_path() / ("openxlsx_" + std::to_string(rd()) + ".xlsx")).string();
    }
    saveAs(fileName, forceOverwrite);
}

void XLDocument::saveAs(std::string_view fileName, bool forceOverwrite)
{
    std::unique_lock<std::shared_mutex> lock(*m_docMutex);

    if (!forceOverwrite and pathExists(fileName)) {
        using namespace std::literals::string_literals;
        throw XLException("XLDocument::saveAs: refusing to overwrite existing file "s + std::string(fileName));
    }



    // [CRITICAL FIX: Prevent source file corruption during saveAs]
    // If the target filename is different from the currently opened file, we must clone the archive FIRST.
    // Otherwise, libzip will commit all pending XML modifications into the original source file.
    if (!m_isEncryptedSession && m_archive.isOpen() && std::string(fileName) != m_filePath && pathExists(m_filePath)) {
        // 1. Close the current archive WITHOUT committing any changes.
        // Note: m_archive.close() might commit if isModified is true. We should ensure it doesn't, but currently we can't easily reset
        // isModified. Actually, the modifications (m_archive.addEntry) are executed LATER in this function! So at this exact moment,
        // m_archive has NO pending zip_source_buffer modifications yet! So it's perfectly safe to close it, copy the file, and reopen the
        // new one.
        m_archive.close();

        // 2. Safely copy the pure, untouched original file to the new destination.
        std::filesystem::copy_file(std::filesystem::u8path(m_filePath),
                                   std::filesystem::u8path(fileName),
                                   std::filesystem::copy_options::overwrite_existing);

        // 3. Re-open the archive, now pointing exclusively to the new target file.
        m_archive.open(std::string(fileName));
    }

    m_filePath = std::string(fileName);
    workbook().updateWorksheetDimensions();

    // Auto-apply calculation enforcement if any formula was written during this session
    if (m_formulaNeedsRecalculation) { execCommand(XLCommand(XLCommandType::SetFullCalcOnLoad)); }

    execCommand(XLCommand(XLCommandType::ResetCalcChain));

    if (m_filePath.size() >= 5 && m_filePath.substr(m_filePath.size() - 5) == ".xlsm") {
        m_contentTypes.updateOverride("/xl/workbook.xml", XLContentType::WorkbookMacroEnabled);
        if (hasMacro()) {
            // Need to make sure bin is registered
            m_contentTypes.addDefault("bin", "application/vnd.ms-office.vbaProject");

            // Note: some macros have a signature file (vbaProjectSignature.bin) but its type is
            // application/vnd.ms-office.vbaProjectSignature The unhandled entries logic will preserve the file, but without the default
            // type it might be ignored. Using Default Extension="bin" handles both.
        }
    }
    else if (m_filePath.size() >= 5 && m_filePath.substr(m_filePath.size() - 5) == ".xlsx") {
        m_contentTypes.updateOverride("/xl/workbook.xml", XLContentType::Workbook);
        if (hasMacro()) {
            // Security: If the user saves an .xlsm (or macro-enabled file) as .xlsx,
            // we must explicitly delete the macro payload and relationships to prevent Excel from reporting a corrupted extension mismatch.
            deleteMacro();
        }
    }

    m_archive.addEntry("[Content_Types].xml", m_contentTypes.xmlData(m_xmlSavingDeclaration));
    m_archive.addEntry("_rels/.rels", m_docRelationships.xmlData(m_xmlSavingDeclaration));
    if (m_wbkRelationships.valid()) {
        m_archive.addEntry("xl/_rels/workbook.xml.rels", m_wbkRelationships.xmlData(m_xmlSavingDeclaration));
    }

    // Ensure the shared strings pugi DOM is in sync with the in-memory cache.
    // appendString() uses a lazy-DOM path (skipping DOM mutation for performance),
    // so we always rebuild the DOM here before serializing it to the ZIP archive.
    if (m_sharedStrings.stringCount() > 0) { m_sharedStrings.rewriteXmlFromCache(); }

    for (auto& item : m_data) {
        if (item.getXmlPath() == "[Content_Types].xml" or item.getXmlPath() == "_rels/.rels" ||
            item.getXmlPath() == "xl/_rels/workbook.xml.rels")
            continue;

        bool xmlIsStandalone = m_xmlSavingDeclaration.standalone_as_bool();
        if ((item.getXmlPath() == "docProps/core.xml") or (item.getXmlPath() == "docProps/app.xml") or
            (item.getXmlPath() == "docProps/custom.xml"))
            xmlIsStandalone = XLXmlStandalone;

        if (item.m_isStreamed) { m_archive.addEntryFromFile(item.getXmlPath(), item.m_streamFilePath); }
        else {
            m_archive.addEntry(
                item.getXmlPath(),
                item.getRawData(
                    XLXmlSavingDeclaration(m_xmlSavingDeclaration.version(), m_xmlSavingDeclaration.encoding(), xmlIsStandalone)));
        }
    }

    for (const auto& entry : m_unhandledEntries) {
        if (std::none_of(m_data.begin(), m_data.end(), [&](const XLXmlData& item) { return item.getXmlPath() == entry.first; })) {
            m_archive.addEntry(entry.first, entry.second);
        }
    }

        if (m_isEncryptedSession) {
        m_archive.save(m_tempDecryptedPath);
        std::ifstream file(m_tempDecryptedPath, std::ios::binary | std::ios::ate);
        if (file) {
            auto size = file.tellg();
            file.seekg(0, std::ios::beg);
            std::vector<uint8_t> zipData(size);
            file.read(reinterpret_cast<char*>(zipData.data()), size);
            file.close();

            auto encryptedData = encryptDocument(zipData, m_encryptionPassword);
            std::ofstream out(m_filePath, std::ios::binary | std::ios::trunc);
            out.write(reinterpret_cast<const char*>(encryptedData.data()), encryptedData.size());
            out.close();
        }
    } else {
        m_archive.save(m_filePath);
    }
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
 * @details Exposes the content types manager, which is required to register new XML parts (like sheets or drawings) so Excel recognizes
 * their purpose.
 */
XLContentTypes& XLDocument::contentTypes() { return m_contentTypes; }

/**
 * @details Provides access to custom properties, enabling the storage of application-specific metadata within the OOXML container.
 */

bool XLDocument::hasMacro() const
{
    // Check in-memory cache first (fastest path)
    for (const auto& entry : m_unhandledEntries) {
        if (entry.first.find("vbaProject.bin") != std::string::npos) { return true; }
    }
    // Fall back to a direct archive probe. Using hasEntry() avoids iterating
    // all entries and is safe even after deleteMacro(), because zip_name_locate
    // correctly handles pending-deleted entries (unlike zip_get_name, which
    // returns nullptr for deleted entries and would cause UB in std::string).
    return m_archive.isOpen() && m_archive.hasEntry("xl/vbaProject.bin");
}

void XLDocument::deleteMacro()
{
    // 1. Remove from unhandled entries map (e.g. vbaProject.bin, vbaProjectSignature.bin, etc.)
    auto it = m_unhandledEntries.begin();
    while (it != m_unhandledEntries.end()) {
        if (it->first.find("vbaProject") != std::string::npos) { it = m_unhandledEntries.erase(it); }
        else {
            ++it;
        }
    }

    // 2. Remove from actual archive if it's already there
    std::vector<std::string> entriesToDelete;
    for (const auto& entry : m_archive.entryNames()) {
        if (entry.find("vbaProject") != std::string::npos) { entriesToDelete.push_back(entry); }
    }
    for (const auto& entry : entriesToDelete) { m_archive.deleteEntry(entry); }

    // 3. Remove ContentType declaration
    if (m_contentTypes.hasOverride("/xl/vbaProject.bin")) { m_contentTypes.deleteOverride("/xl/vbaProject.bin"); }
    if (m_contentTypes.hasOverride("/xl/vbaProjectSignature.bin")) { m_contentTypes.deleteOverride("/xl/vbaProjectSignature.bin"); }
    if (m_contentTypes.hasDefault("bin")) { m_contentTypes.deleteDefault("bin"); }

    // 4. Remove Workbook Relationship
    auto rels = workbookRelationships().relationships();
    for (const auto& rel : rels) {
        if (rel.type() == XLRelationshipType::VBAProject) { workbookRelationships().deleteRelationship(rel); }
    }

    // 5. Ensure the workbook root node does not have the 'macroEnabled' extension type
    // Not strictly necessary to delete the relationship, but Excel complains if we leave it hanging.
}

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
 * @details Calculates the next available auto-incrementing ID across all tables to ensure unique table references within the document, a
 * strict OOXML requirement.
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
 * @details Provides access to the core workbook component, serving as the gateway to manage worksheets, chartsheets, and global
 * definitions.
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
 * @details Provides access to the global stylesheet manager to retrieve or define fonts, fills, borders, and number formats used across the
 * workbook.
 */
XLStyles& XLDocument::styles() { return m_styles; }

/**
 * @details Probes the archive index for sheet relationships to conditionally access dependencies, avoiding eager and expensive XML
 * allocations for untouched components.
 */
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
        m_archive.addEntry(drawingFilename,
                           "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<xdr:wsDr "
                           "xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
                           "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
                           "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
                           "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                           "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"></xdr:wsDr>");
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
        m_archive.addEntry(commentsFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
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

XLPersons& XLDocument::persons() {
    if (!m_persons.valid()) {
        std::string personsFilename = "xl/persons/person.xml";
        if (!m_archive.hasEntry(personsFilename)) {
            m_archive.addEntry(personsFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<personList xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
            m_contentTypes.addOverride("/" + personsFilename, XLContentType::Persons);
            m_wbkRelationships.addRelationship(XLRelationshipType::Person, "persons/person.xml");
        }
        XLXmlData* xmlData = getXmlData(personsFilename, true);
        if (xmlData == nullptr) {
            xmlData = &m_data.emplace_back(this, personsFilename, "", XLContentType::Persons);
        }
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
void XLDocument::cleanupSharedStrings()
{
    std::unique_lock<std::shared_mutex> docLock(*m_docMutex);
    std::unique_lock<std::shared_mutex> strLock(*m_sharedStringMutex);
    const size_t                        oldStringCount = m_sharedStringCache.size();
    std::vector<int32_t>                indexMap(oldStringCount, -1);
    int32_t                             newStringCount = 1;

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

    XLStringArena                 newArena;
    std::vector<std::string_view> newStringCache(static_cast<size_t>(newStringCount));
    newStringCache[0] = newArena.store("");
    for (size_t oldIdx = 0; oldIdx < oldStringCount; ++oldIdx) {
        if (const int32_t newIdx = indexMap[oldIdx]; newIdx > 0)
            newStringCache[static_cast<size_t>(newIdx)] = newArena.store(m_sharedStringCache[oldIdx]);
    }

    m_sharedStringArena = std::move(newArena);
    m_sharedStringCache = std::move(newStringCache);
    // rebuilding index
    m_sharedStringIndex.clear();
    for (size_t i = 0; i < m_sharedStringCache.size(); ++i) {
        m_sharedStringIndex.emplace(m_sharedStringCache[i], static_cast<int32_t>(i));
    }

    if (static_cast<int32_t>(m_sharedStringCache.size()) != m_sharedStrings.rewriteXmlFromCache())
        throw XLInternalError("XLDocument::cleanupSharedStrings: failed to rewrite shared string table - document would be corrupted");
}

//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

/**
 * @details Serves as the foundational bridge between the raw ZIP layer and the parsed DOM layer, safely extracting content only when
 * explicitly requested.
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

    std::string templateStr = fmt::format(R"(<?xml version="1.0" encoding="UTF-8"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" name="{0}" sourceName="{1}">
  <extLst>
    <ext xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" uri="{{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}}">
      <x15:tableSlicerCache tableId="{2}" column="{3}"/>
    </ext>
  </extLst>
</slicerCacheDefinition>)",
                                          name,
                                          sourceName,
                                          tableId,
                                          tableColumnId);

    m_archive.addEntry(filename, templateStr);
    m_contentTypes.addOverride("/" + filename, XLContentType::SlicerCache);

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(filename, DO_NOT_THROW);
    if (xmlData == nullptr) { m_data.emplace_back(this, filename, templateStr, XLContentType::SlicerCache); }

    // Add relationship to workbook.xml
    m_wbkRelationships.addRelationship(XLRelationshipType::SlicerCache, "/" + filename);
    std::string rId = m_wbkRelationships.relationshipByTarget("/" + filename).id();

    // Add extLst to workbook.xml to reference the slicer cache
    XMLNode wbkNode = m_workbook.xmlDocument().document_element();
    XMLNode extLst  = wbkNode.child("extLst");
    // Ensure mc namespaces are present
    if (!wbkNode.attribute("xmlns:mc"))
        wbkNode.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
    if (!wbkNode.attribute("xmlns:x15"))
        wbkNode.append_attribute("xmlns:x15").set_value("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
    if (!wbkNode.attribute("mc:Ignorable")) wbkNode.append_attribute("mc:Ignorable").set_value("x15");

    if (extLst.empty()) extLst = wbkNode.append_child("extLst");

    XMLNode ext = extLst.find_child_by_attribute("uri", "{46BE6895-7355-4a93-B00E-2C351335B9C9}");
    if (ext.empty()) {
        ext = extLst.append_child("ext");
        ext.append_attribute("uri").set_value("{46BE6895-7355-4a93-B00E-2C351335B9C9}");
        ext.append_attribute("xmlns:x15").set_value("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
    }

    XMLNode slicerCaches = ext.child("x15:slicerCaches");
    if (slicerCaches.empty()) {
        slicerCaches = ext.append_child("x15:slicerCaches");
        // Must inject x14 at the top of workbook or at least correctly
        slicerCaches.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
    }
    // Excel ignores or strips if the mc:Ignorable is missing or something?

    slicerCaches.append_child("x14:slicerCache").append_attribute("r:id").set_value(rId.c_str());

    // Add definedName to workbook.xml
    XMLNode definedNames = wbkNode.child("definedNames");
    if (definedNames.empty()) {
        XMLNode calcPr = wbkNode.child("calcPr");
        if (!calcPr.empty()) { definedNames = wbkNode.insert_child_before("definedNames", calcPr); }
        else {
            definedNames = wbkNode.append_child("definedNames");
        }
    }
    XMLNode definedName = definedNames.append_child("definedName");
    definedName.append_attribute("name").set_value(std::string(name).c_str());
    definedName.text().set("#N/A");
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

    std::string templateStr = fmt::format(R"(<?xml version="1.0" encoding="UTF-8"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" name="{0}" sourceName="{1}">
  <pivotTables>
    <pivotTable tabId="{2}" name="{3}"/>
  </pivotTables>
  <data>
    <tabular pivotCacheId="{4}" showMissing="false">
      <items count="1">
        <i x="0" s="true"/>
      </items>
    </tabular>
  </data>
</slicerCacheDefinition>)",
                                          name,
                                          sourceName,
                                          sheetId,
                                          pivotTableName,
                                          pivotCacheId);

    m_archive.addEntry(filename, templateStr);
    m_contentTypes.addOverride("/" + filename, XLContentType::SlicerCache);

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(filename, DO_NOT_THROW);
    if (xmlData == nullptr) { m_data.emplace_back(this, filename, templateStr, XLContentType::SlicerCache); }

    // Add relationship to workbook.xml
    m_wbkRelationships.addRelationship(XLRelationshipType::SlicerCache, "/" + filename);
    std::string rId = m_wbkRelationships.relationshipByTarget("/" + filename).id();

    // Add extLst to workbook.xml to reference the slicer cache
    XMLNode wbkNode = m_workbook.xmlDocument().document_element();
    XMLNode extLst  = wbkNode.child("extLst");
    // Ensure mc namespaces are present
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
        ext = extLst.append_child("ext");
        ext.append_attribute("uri").set_value("{BBE1A952-AA13-448e-AADC-164F8A28A991}");
        ext.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
    }

    XMLNode slicerCaches = ext.child("x14:slicerCaches");
    if (slicerCaches.empty()) {
        slicerCaches = ext.append_child("x14:slicerCaches");
        slicerCaches.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
    }

    slicerCaches.append_child("x14:slicerCache").append_attribute("r:id").set_value(rId.c_str());

    // Add definedName to workbook.xml
    XMLNode definedNames = wbkNode.child("definedNames");
    if (definedNames.empty()) {
        XMLNode calcPr = wbkNode.child("calcPr");
        if (!calcPr.empty()) { definedNames = wbkNode.insert_child_before("definedNames", calcPr); }
        else {
            definedNames = wbkNode.append_child("definedNames");
        }
    }
    XMLNode definedName = definedNames.append_child("definedName");
    definedName.append_attribute("name").set_value(std::string(name).c_str());
    definedName.text().set("#N/A");

    return filename;
}

std::string XLDocument::createSlicer(std::string_view name, std::string_view cacheName, std::string_view caption)
{
    using namespace std::literals::string_literals;

    uint32_t    num      = 1;
    std::string filename = fmt::format("xl/slicers/slicer{}.xml", num);
    while (m_archive.hasEntry(filename)) {
        ++num;
        filename = fmt::format("xl/slicers/slicer{}.xml", num);
    }

    std::string templateStr = fmt::format(R"(<?xml version="1.0" encoding="UTF-8"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" mc:Ignorable="xr10">
  <slicer name="{0}" cache="{1}" caption="{2}" rowHeight="251883"/>
</slicers>)",
                                          name,
                                          cacheName,
                                          caption);

    m_archive.addEntry(filename, templateStr);
    m_contentTypes.addOverride("/" + filename, XLContentType::Slicer);

    constexpr bool DO_NOT_THROW = true;
    XLXmlData*     xmlData      = getXmlData(filename, DO_NOT_THROW);
    if (xmlData == nullptr) { m_data.emplace_back(this, filename, templateStr, XLContentType::Slicer); }

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

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

using namespace OpenXLSX;

namespace
{
    using namespace std::literals;

    // [Content_Types].xml
    constexpr std::string_view templateContentTypes =
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>)";

    // .rels
    constexpr std::string_view templateRootRels =
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>)";

    // xl/workbook.xml
    constexpr std::string_view templateWorkbook =
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/>
    <workbookPr defaultThemeVersion="164011"/>
    <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="14805" windowHeight="8010"/></bookViews>
    <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
    <calcPr calcId="122211"/>
</workbook>)";

    // xl/_rels/workbook.xml.rels
    constexpr std::string_view templateWorkbookRels =
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
    </Relationships>)"sv;

    // xl/worksheets/sheet1.xml
    constexpr std::string_view templateSheet =
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
    <dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
    <sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>)";

    // xl/styles.xml
    constexpr std::string_view templateStyles =
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>
    <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
    <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
    <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
    <dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>)";

    // docProps/app.xml
    constexpr std::string_view templateApp =
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>OpenXLSX</Application>
</Properties>)";

    // xl/theme/theme1.xml
    constexpr std::string_view templateTheme =
        R"(<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="5B9BD5"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>)"sv;
    // docProps/core.xml
    constexpr std::string_view templateCore =
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>OpenXLSX</dc:creator>
    <cp:lastModifiedBy>OpenXLSX</cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">2023-01-01T00:00:00Z</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">2023-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>)"sv;

    // xl/sharedStrings.xml
    constexpr std::string_view templateSharedStrings =
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sst>)"sv;
}    // namespace

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
        m_archive.addEntry(
            item.getXmlPath(),
            item.getRawData(XLXmlSavingDeclaration(m_xmlSavingDeclaration.version(), m_xmlSavingDeclaration.encoding(), xmlIsStandalone)));
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
 * @details Provides read access to standard OOXML document metadata (core and extended properties) such as creator, application version, or creation date.
 */
std::string XLDocument::property(XLProperty prop) const
{
    switch (prop) {
        case XLProperty::Application:
            return m_appProperties.property("Application");
        case XLProperty::AppVersion:
            return m_appProperties.property("AppVersion");
        case XLProperty::Category:
            return m_coreProperties.category();
        case XLProperty::Company:
            return m_appProperties.property("Company");
        case XLProperty::CreationDate:
            return m_coreProperties.property("dcterms:created");
        case XLProperty::Creator:
            return m_coreProperties.creator();
        case XLProperty::Description:
            return m_coreProperties.description();
        case XLProperty::DocSecurity:
            return m_appProperties.property("DocSecurity");
        case XLProperty::HyperlinkBase:
            return m_appProperties.property("HyperlinkBase");
        case XLProperty::HyperlinksChanged:
            return m_appProperties.property("HyperlinksChanged");
        case XLProperty::Keywords:
            return m_coreProperties.keywords();
        case XLProperty::LastModifiedBy:
            return m_coreProperties.lastModifiedBy();
        case XLProperty::LastPrinted:
            return m_coreProperties.lastPrinted();
        case XLProperty::LinksUpToDate:
            return m_appProperties.property("LinksUpToDate");
        case XLProperty::Manager:
            return m_appProperties.property("Manager");
        case XLProperty::ModificationDate:
            return m_coreProperties.property("dcterms:modified");
        case XLProperty::ScaleCrop:
            return m_appProperties.property("ScaleCrop");
        case XLProperty::SharedDoc:
            return m_appProperties.property("SharedDoc");
        case XLProperty::Subject:
            return m_coreProperties.subject();
        case XLProperty::Title:
            return m_coreProperties.title();
        default:
            return "";    // To silence compiler warning.
    }
}

/**
 * @brief extract an integer major version v1 and minor version v2 from a string
 * @details Extracts integer major and minor versions from a string, ensuring the formatting strictly complies with the OOXML AppVersion schema requirements.
 * [0-9]{1,2}\.[0-9]{1,4}  (Example: 14.123)
 * @param versionString the string to process
 * @param majorVersion by reference: store the major version here
 * @param minorVersion by reference: store the minor version here
 * @return true if string adheres to format & version numbers could be extracted
 * @return false in case of failure
 */
bool getAppVersion(std::string_view versionString, int& majorVersion, int& minorVersion)
{
    // ===== const expressions for hardcoded version limits
    constexpr int minMajorV = 0, maxMajorV = 99;      // allowed value range for major version number
    constexpr int minMinorV = 0, maxMinorV = 9999;    //          "          for minor   "

    const size_t begin  = versionString.find_first_not_of(" \t");
    const size_t dotPos = versionString.find_first_of('.');

    // early failure if string is only blanks or does not contain a dot
    if (begin == std::string_view::npos or dotPos == std::string_view::npos) return false;

    const size_t end = versionString.find_last_not_of(" \t");
    if (begin != std::string_view::npos and dotPos != std::string_view::npos) {
        const std::string strMajorVersion = std::string(versionString.substr(begin, dotPos - begin));
        const std::string strMinorVersion = std::string(versionString.substr(dotPos + 1, end - dotPos));
        try {
            size_t pos;
            majorVersion = std::stoi(strMajorVersion, &pos);
            if (pos != strMajorVersion.length()) throw 1;
            minorVersion = std::stoi(strMinorVersion, &pos);
            if (pos != strMinorVersion.length()) throw 1;
        }
        catch (...) {
            return false;    // conversion failed or did not convert the full string
        }
    }
    if (majorVersion < minMajorV or majorVersion > maxMajorV or minorVersion < minMinorV ||
        minorVersion > maxMinorV)    // final range check
        return false;

    return true;
}

/**
 * @details Mutates document metadata, explicitly validating format-specific constraints (e.g., boolean strings or version formatting 'XX.XXXX') to prevent schema validation failures in Excel.
 */
void XLDocument::setProperty(XLProperty prop, std::string_view value)
{
    const std::string valStr(value);
    switch (prop) {
        case XLProperty::Application:
            m_appProperties.setProperty("Application", valStr);
            break;
        case XLProperty::AppVersion: {
            int minorVersion, majorVersion;
            // ===== Check for the format "XX.XXXX", with X being a number.
            if (!getAppVersion(value, majorVersion, minorVersion)) throw XLPropertyError("Invalid property value: " + valStr);
            m_appProperties.setProperty("AppVersion", fmt::format("{}.{}", majorVersion, minorVersion));
            break;
        }

        case XLProperty::Category:
            m_coreProperties.setCategory(valStr);
            break;
        case XLProperty::Company:
            m_appProperties.setProperty("Company", valStr);
            break;
        case XLProperty::CreationDate:
            m_coreProperties.setProperty("dcterms:created", valStr);
            break;
        case XLProperty::Creator:
            m_coreProperties.setCreator(valStr);
            break;
        case XLProperty::Description:
            m_coreProperties.setDescription(valStr);
            break;
        case XLProperty::DocSecurity:
            if (value == "0" or value == "1" or value == "2" or value == "4" or value == "8")
                m_appProperties.setProperty("DocSecurity", valStr);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::HyperlinkBase:
            m_appProperties.setProperty("HyperlinkBase", valStr);
            break;
        case XLProperty::HyperlinksChanged:
            if (value == "true" or value == "false")
                m_appProperties.setProperty("HyperlinksChanged", valStr);
            else
                throw XLPropertyError("Invalid property value");

            break;

        case XLProperty::Keywords:
            m_coreProperties.setKeywords(valStr);
            break;
        case XLProperty::LastModifiedBy:
            m_coreProperties.setLastModifiedBy(valStr);
            break;
        case XLProperty::LastPrinted:
            m_coreProperties.setLastPrinted(valStr);
            break;
        case XLProperty::LinksUpToDate:
            if (value == "true" or value == "false")
                m_appProperties.setProperty("LinksUpToDate", valStr);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::Manager:
            m_appProperties.setProperty("Manager", valStr);
            break;
        case XLProperty::ModificationDate:
            m_coreProperties.setProperty("dcterms:modified", valStr);
            break;
        case XLProperty::ScaleCrop:
            if (value == "true" or value == "false")
                m_appProperties.setProperty("ScaleCrop", valStr);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::SharedDoc:
            if (value == "true" or value == "false")
                m_appProperties.setProperty("SharedDoc", valStr);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::Subject:
            m_coreProperties.setSubject(valStr);
            break;
        case XLProperty::Title:
            m_coreProperties.setTitle(valStr);
            break;
    }
}

/**
 * @details Clears a specific metadata property to clean up unneeded, obsolete, or sensitive document traces.
 */
void XLDocument::deleteProperty(XLProperty theProperty) { setProperty(theProperty, ""); }

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
bool XLDocument::hasSheetRelationships(uint16_t sheetXmlNo) const
{
    using namespace std::literals::string_literals;
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
XLRelationships XLDocument::sheetRelationships(uint16_t sheetXmlNo)
{
    using namespace std::literals::string_literals;
    std::string relsFilename = "xl/worksheets/_rels/sheet"s + std::to_string(sheetXmlNo) + ".xml.rels"s;

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

XLChart XLDocument::createChart()
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

    return XLChart(xmlData);
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
constexpr bool THROW_ON_INVALID = true;
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
        case XLCommandType::AddChartsheet:
            // TODO: To be implemented
            break;
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
            const auto sheetPath  = "/xl/worksheets/sheet" + std::to_string(internalID) + ".xml";
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

namespace OpenXLSX
{
    /**
     * @brief Return a hexadecimal digit as character that is the equivalent of value
     */
    [[nodiscard]] constexpr char hexDigit(unsigned int value)
    {
        if (value > 0xf) return 0;
        if (value < 0xa) return static_cast<char>(value + '0');
        return static_cast<char>((value - 0xa) + 'a');
    }

    /**
     * @details Converts raw byte arrays into hex-encoded strings, required for constructing valid OOXML password hashes.
     */
    std::string BinaryAsHexString(gsl::span<const std::byte> data)
    {
        if (data.empty()) return "";
        std::string strAssemble(data.size() * 2, 0);
        for (size_t i = 0; i < data.size(); ++i) {
            auto valueByte         = static_cast<uint8_t>(data[i]);
            strAssemble[i * 2]     = hexDigit(valueByte >> 4);
            strAssemble[i * 2 + 1] = hexDigit(valueByte & 0x0f);
        }
        return strAssemble;
    }

    /**
     * @details Implements the legacy Excel hash algorithm to verify or secure sheet protection features.
     */
    uint16_t ExcelPasswordHash(std::string_view password)
    {
        uint16_t       wPasswordHash = 0;
        const uint16_t cchPassword   = gsl::narrow_cast<uint16_t>(password.length());

        for (uint16_t pos = 0; pos < cchPassword; ++pos) {
            uint32_t byteHash = static_cast<uint32_t>(password[pos]) << ((pos + 1) % 15);
            byteHash          = (byteHash >> 15) | (byteHash & 0x7fff);
            wPasswordHash ^= static_cast<uint16_t>(byteHash);
        }
        wPasswordHash ^= cchPassword ^ 0xce4b;
        return wPasswordHash;
    }

    /**
     * @details Bridges the legacy integer hash to the required XML string format, simplifying integration with the DOM attributes.
     */
    std::string ExcelPasswordHashAsString(std::string_view password)
    {
        const uint16_t                 pw = ExcelPasswordHash(password);
        const std::array<std::byte, 2> hashData{static_cast<std::byte>(pw >> 8), static_cast<std::byte>(pw & 0xff)};
        return BinaryAsHexString(gsl::make_span(hashData.data(), 2));
    }

    /**
     * @details Split a path into its constituent entries (directories and file).
     * Rationale: Standard path decomposition used for relative path calculation.
     * Uses std::string_view for zero-copy parsing.
     */
    std::vector<std::string> disassemblePath(std::string_view path, bool eliminateDots = true)
    {
        std::vector<std::string> result;
        size_t                   startpos = (!path.empty() && path.front() == '/' ? 1 : 0);
        size_t                   pos;
        do {
            pos = path.find('/', startpos);
            if (pos == startpos) throw XLInternalError("path must not contain two subsequent forward slashes");

            std::string_view dirEntry = path.substr(startpos, pos - startpos);
            if (!dirEntry.empty()) {
                if (eliminateDots) {
                    if (dirEntry == ".") {}
                    else if (dirEntry == "..") {
                        if (!result.empty())
                            result.pop_back();
                        else
                            throw XLInternalError("no remaining directory to exit with ..");
                    }
                    else
                        result.emplace_back(dirEntry);
                }
                else
                    result.emplace_back(dirEntry);
            }
            startpos = pos + 1;
        }
        while (pos != std::string_view::npos);

        return result;
    }

    /**
     * @details Calculate the relative path from B to A.
     * Rationale: Resolves internal OOXML relationship paths, which are often stored
     * relative to the source part's directory.
     */
    std::string getPathARelativeToPathB(std::string_view pathA, std::string_view pathB)
    {
        size_t startpos = 0;
        while (startpos < pathA.length() && startpos < pathB.length() && pathA[startpos] == pathB[startpos]) ++startpos;
        while (startpos > 0 and pathA[startpos - 1] != '/') --startpos;
        if (startpos == 0) throw XLInternalError("getPathARelativeToPathB: pathA and pathB have no common beginning");

        std::vector<std::string> dirEntriesB = disassemblePath(pathB.substr(startpos));
        if (!dirEntriesB.empty() && (!pathB.empty() && pathB.back() != '/')) dirEntriesB.pop_back();

        std::string result;
        for (size_t i = 0; i < dirEntriesB.size(); ++i) result += "../";
        result += std::string(pathA.substr(startpos));

        return result;
    }

    /**
     * @details Standardize path by resolving '.' and '..' segments.
     * Rationale: Normalizes paths to ensure consistent internal ZIP archive lookups.
     */
    std::string eliminateDotAndDotDotFromPath(std::string_view path)
    {
        if (path.empty()) return "";
        std::vector<std::string> dirEntries = disassemblePath(path);

        std::string result = (!path.empty() && path.front() == '/') ? "/" : "";
        if (!dirEntries.empty()) {
            result += dirEntries[0];
            for (size_t i = 1; i < dirEntries.size(); ++i) {
                result += "/";
                result += dirEntries[i];
            }
        }

        if (((result.length() > 1 or result.front() != '/') && (!path.empty() && path.back() == '/'))) result += "/";
        return result;
    }

}    // namespace OpenXLSX

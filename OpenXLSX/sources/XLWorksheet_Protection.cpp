#include <fmt/format.h>
#include <string>

#include "XLUtilities.hpp"
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

using namespace OpenXLSX;

bool XLWorksheet::protect(const XLSheetProtectionOptions& options, std::string_view password)
{
    XMLNode sheetNode = xmlDocument().document_element();
    auto    node      = sheetNode.child("sheetProtection");
    if (!node.empty()) sheetNode.remove_child("sheetProtection");

    XMLNode protNode = sheetNode.insert_child_before("sheetProtection", sheetNode.child("protectedRanges"));
    if (protNode.empty()) {    // If protectedRanges is missing, try inserting before sheetData or something valid... but m_nodeOrder handles appending.
        protNode = sheetNode.append_child("sheetProtection"); // Temporarily, we will use appendAndSetNodeAttribute to fix order
    }
    
    // We will build attributes dynamically, but simpler is to use the existing `appendAndSetNodeAttribute` utility
    // to guarantee correct node placement according to `m_nodeOrder`.
    
    sheetNode.remove_child("sheetProtection"); // ensure it's clean
    
    // The OOXML attributes:
    // If option is true, meaning "allowed/true", we set it according to defaults. 
    // Wait, the OOXML specification uses '1' or 'true' to indicate the restriction is ON for some, and OFF for others.
    // Let's be precise:
    // sheet="1" means sheet protection is ON.
    // objects="1" means objects protection is ON.
    // scenarios="1" means scenarios protection is ON.
    // formatCells="0" means formatting cells is NOT allowed (restriction is ON).
    
    // Default of attributes in OOXML if omitted:
    // sheet=false, objects=false, scenarios=false, 
    // formatCells=true (allowed), formatColumns=true, formatRows=true, 
    // insertColumns=true, insertRows=true, insertHyperlinks=true,
    // deleteColumns=true, deleteRows=true, sort=true, autoFilter=true, pivotTables=true
    // selectLockedCells=false (actually the default is false if sheet is protected? No, MS Excel defaults to true for selection if protected).
    
    // Actually, in OpenXLSX: allowFormatCells(false) sets formatCells="true".
    // Wait, let's look at the current implementation:
    // allowInsertColumns(false) -> "true"
    // So formatCells="true" means RESTRICTED (not allowed).
    
    // Wait! Let's check `allowInsertColumns(bool set)`:
    // `(!set ? "true" : "false")` -> if set=true (allowed), the attribute is "false" (not restricted).
    // So the XML attribute means "is_restricted".
    
    // To keep it clean, let's just use appendAndSetNodeAttribute.
    auto setAttr = [&](const char* name, bool restricted) {
        appendAndSetNodeAttribute(sheetNode, "sheetProtection", name, restricted ? "1" : "0", XLKeepAttributes, m_nodeOrder);
    };

    // For sheet, objects, scenarios: true means protected (restricted)
    setAttr("sheet", options.sheet);
    setAttr("objects", options.objects);
    setAttr("scenarios", options.scenarios);

    // For others, true means allowed, so restricted is !options.xxx
    setAttr("formatCells", !options.formatCells);
    setAttr("formatColumns", !options.formatColumns);
    setAttr("formatRows", !options.formatRows);
    setAttr("insertColumns", !options.insertColumns);
    setAttr("insertRows", !options.insertRows);
    setAttr("insertHyperlinks", !options.insertHyperlinks);
    setAttr("deleteColumns", !options.deleteColumns);
    setAttr("deleteRows", !options.deleteRows);
    setAttr("sort", !options.sort);
    setAttr("autoFilter", !options.autoFilter);
    setAttr("pivotTables", !options.pivotTables);

    // selectLockedCells / selectUnlockedCells are also "allow" options
    setAttr("selectLockedCells", !options.selectLockedCells);
    setAttr("selectUnlockedCells", !options.selectUnlockedCells);

    if (!password.empty()) {
        std::string hash = ExcelPasswordHashAsString(password);
        appendAndSetNodeAttribute(sheetNode, "sheetProtection", "password", hash, XLKeepAttributes, m_nodeOrder);
    }

    return true;
}

bool XLWorksheet::protectSheet(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "sheet", (set ? "true" : "false"), XLKeepAttributes, m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::protectObjects(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "objects", (set ? "true" : "false"), XLKeepAttributes, m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::protectScenarios(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "scenarios", (set ? "true" : "false"), XLKeepAttributes, m_nodeOrder)
               .empty() == false;
}

bool XLWorksheet::allowInsertColumns(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode,
                                     "sheetProtection",
                                     "insertColumns",
                                     (!set ? "true" : "false"),
                                     XLKeepAttributes,
                                     m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowInsertRows(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "insertRows", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowDeleteColumns(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode,
                                     "sheetProtection",
                                     "deleteColumns",
                                     (!set ? "true" : "false"),
                                     XLKeepAttributes,
                                     m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowDeleteRows(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "deleteRows", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowFormatCells(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "formatCells", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowFormatColumns(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode,
                                     "sheetProtection",
                                     "formatColumns",
                                     (!set ? "true" : "false"),
                                     XLKeepAttributes,
                                     m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowFormatRows(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "formatRows", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowInsertHyperlinks(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode,
                                     "sheetProtection",
                                     "insertHyperlinks",
                                     (!set ? "true" : "false"),
                                     XLKeepAttributes,
                                     m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowSort(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "sort", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowAutoFilter(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "autoFilter", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowPivotTables(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "pivotTables", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowSelectLockedCells(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode,
                                     "sheetProtection",
                                     "selectLockedCells",
                                     (!set ? "true" : "false"),
                                     XLKeepAttributes,
                                     m_nodeOrder)
               .empty() == false;
}
bool XLWorksheet::allowSelectUnlockedCells(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode,
                                     "sheetProtection",
                                     "selectUnlockedCells",
                                     (!set ? "true" : "false"),
                                     XLKeepAttributes,
                                     m_nodeOrder)
               .empty() == false;
}

bool XLWorksheet::setPasswordHash(std::string hash)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "password", hash, XLKeepAttributes, m_nodeOrder).empty() == false;
}

bool XLWorksheet::setPassword(std::string password)
{ return setPasswordHash(password.length() ? ExcelPasswordHashAsString(password) : std::string("")); }
bool XLWorksheet::clearPassword() { return setPasswordHash(""); }
bool XLWorksheet::clearSheetProtection() { return xmlDocument().document_element().remove_child("sheetProtection"); }

bool XLWorksheet::sheetProtected() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return false;
    return node.attribute("sheet").as_bool(false);
}
bool XLWorksheet::objectsProtected() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return false;
    return node.attribute("objects").as_bool(false);
}
bool XLWorksheet::scenariosProtected() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return false;
    return node.attribute("scenarios").as_bool(false);
}

bool XLWorksheet::insertColumnsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("insertColumns").as_bool(true);
}
bool XLWorksheet::insertRowsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("insertRows").as_bool(true);
}
bool XLWorksheet::deleteColumnsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("deleteColumns").as_bool(true);
}
bool XLWorksheet::deleteRowsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("deleteRows").as_bool(true);
}
bool XLWorksheet::formatCellsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("formatCells").as_bool(true);
}
bool XLWorksheet::formatColumnsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("formatColumns").as_bool(true);
}
bool XLWorksheet::formatRowsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("formatRows").as_bool(true);
}
bool XLWorksheet::insertHyperlinksAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("insertHyperlinks").as_bool(true);
}
bool XLWorksheet::sortAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("sort").as_bool(true);
}
bool XLWorksheet::autoFilterAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("autoFilter").as_bool(true);
}
bool XLWorksheet::pivotTablesAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("pivotTables").as_bool(true);
}
bool XLWorksheet::selectLockedCellsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("selectLockedCells").as_bool(false);
}
bool XLWorksheet::selectUnlockedCellsAllowed() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return true;
    return not node.attribute("selectUnlockedCells").as_bool(false);
}

std::string XLWorksheet::passwordHash() const
{
    auto node = xmlDocument().document_element().child("sheetProtection");
    if (node.empty()) return "";
    return node.attribute("password").value();
}
bool XLWorksheet::passwordIsSet() const { return !passwordHash().empty(); }

std::string XLWorksheet::sheetProtectionSummary() const
{
    return fmt::format("sheet: {}, objects: {}, scenarios: {}, insertColumns: {}, insertRows: {}, deleteColumns: {}, deleteRows: {}, "
                       "formatCells: {}, formatColumns: {}, formatRows: {}, insertHyperlinks: {}, sort: {}, autoFilter: {}, pivotTables: "
                       "{}, selectLockedCells: {}, selectUnlockedCells: {}, passwordHash: \"{}\"",
                       sheetProtected() ? "true" : "false",
                       objectsProtected() ? "true" : "false",
                       scenariosProtected() ? "true" : "false",
                       insertColumnsAllowed() ? "allowed" : "protected",
                       insertRowsAllowed() ? "allowed" : "protected",
                       deleteColumnsAllowed() ? "allowed" : "protected",
                       deleteRowsAllowed() ? "allowed" : "protected",
                       formatCellsAllowed() ? "allowed" : "protected",
                       formatColumnsAllowed() ? "allowed" : "protected",
                       formatRowsAllowed() ? "allowed" : "protected",
                       insertHyperlinksAllowed() ? "allowed" : "protected",
                       sortAllowed() ? "allowed" : "protected",
                       autoFilterAllowed() ? "allowed" : "protected",
                       pivotTablesAllowed() ? "allowed" : "protected",
                       selectLockedCellsAllowed() ? "allowed" : "protected",
                       selectUnlockedCellsAllowed() ? "allowed" : "protected",
                       passwordHash());
}

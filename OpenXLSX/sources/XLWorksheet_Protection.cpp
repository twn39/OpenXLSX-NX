#include <string>
#include <fmt/format.h>

#include "XLWorksheet.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

bool XLWorksheet::protectSheet(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "sheet", (set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::protectObjects(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "objects", (set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::protectScenarios(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "scenarios", (set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}

bool XLWorksheet::allowInsertColumns(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "insertColumns", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowInsertRows(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "insertRows", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowDeleteColumns(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "deleteColumns", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowDeleteRows(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "deleteRows", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowFormatCells(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "formatCells", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowFormatColumns(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "formatColumns", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowFormatRows(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "formatRows", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowInsertHyperlinks(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "insertHyperlinks", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowSort(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "sort", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowAutoFilter(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "autoFilter", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowPivotTables(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "pivotTables", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowSelectLockedCells(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "selectLockedCells", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowSelectUnlockedCells(bool set)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "selectUnlockedCells", (!set ? "true" : "false"), XLKeepAttributes, m_nodeOrder).empty() == false;
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
    return fmt::format("sheet: {}, objects: {}, scenarios: {}, insertColumns: {}, insertRows: {}, deleteColumns: {}, deleteRows: {}, formatCells: {}, formatColumns: {}, formatRows: {}, insertHyperlinks: {}, sort: {}, autoFilter: {}, pivotTables: {}, selectLockedCells: {}, selectUnlockedCells: {}, passwordHash: \"{}\"",
                       sheetProtected() ? "true" : "false", objectsProtected() ? "true" : "false", scenariosProtected() ? "true" : "false", insertColumnsAllowed() ? "allowed" : "protected", insertRowsAllowed() ? "allowed" : "protected", deleteColumnsAllowed() ? "allowed" : "protected", deleteRowsAllowed() ? "allowed" : "protected", formatCellsAllowed() ? "allowed" : "protected", formatColumnsAllowed() ? "allowed" : "protected", formatRowsAllowed() ? "allowed" : "protected", insertHyperlinksAllowed() ? "allowed" : "protected", sortAllowed() ? "allowed" : "protected", autoFilterAllowed() ? "allowed" : "protected", pivotTablesAllowed() ? "allowed" : "protected", selectLockedCellsAllowed() ? "allowed" : "protected", selectUnlockedCellsAllowed() ? "allowed" : "protected", passwordHash());
}


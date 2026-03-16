// ===== External Includes ===== //
#include <pugixml.hpp>
#include <string_view>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLTables.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

namespace {
    const std::vector<std::string_view> TableNodeOrder = {
        "autoFilter",
        "sortState",
        "tableColumns",
        "tableStyleInfo",
        "extLst"
    };
}


/**
 * @details The constructor creates an instance of the superclass, XLXmlFile
 */
XLTables::XLTables(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
    if (xmlData && xmlData->getXmlType() != XLContentType::Table) throw XLInternalError("XLTables constructor: Invalid XML data.");
    
    auto root = xmlDocument().document_element();
    if (root.empty() || std::string_view(root.name()) != "table") {
        xmlDocument().remove_children();
        root = xmlDocument().append_child("table");
        root.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        
        // Use a unique ID from the document
        uint32_t tableId = parentDoc().nextTableId();
        root.append_attribute("id").set_value(std::to_string(tableId).c_str());
        
        std::string defaultName = "Table" + std::to_string(tableId);
        root.append_attribute("name").set_value(defaultName.c_str());
        root.append_attribute("displayName").set_value(defaultName.c_str());
    }

    // Ensure tableStyleInfo has default attributes in correct order
    auto info = appendAndGetNode(root, "tableStyleInfo", TableNodeOrder);
    if (info.attribute("showFirstColumn").empty()) info.append_attribute("showFirstColumn").set_value("false");
    if (info.attribute("showLastColumn").empty()) info.append_attribute("showLastColumn").set_value("false");
    if (info.attribute("showRowStripes").empty()) info.append_attribute("showRowStripes").set_value("true");
    if (info.attribute("showColumnStripes").empty()) info.append_attribute("showColumnStripes").set_value("false");
}

/**
 * @details Get the name of the table
 */
std::string XLTables::name() const
{
    return xmlDocument().document_element().attribute("name").value();
}

/**
 * @details Set the name of the table. Table names cannot have spaces.
 */
void XLTables::setName(std::string_view name)
{
    if (name.find(' ') != std::string_view::npos) {
        throw XLInputError("Table names cannot contain spaces.");
    }
    auto docNode = xmlDocument().document_element();
    appendAndSetAttribute(docNode, "name", std::string(name));
    appendAndSetAttribute(docNode, "displayName", std::string(name));
}

/**
 * @details Get the display name of the table
 */
std::string XLTables::displayName() const
{
    return xmlDocument().document_element().attribute("displayName").value();
}

/**
 * @details Set the display name of the table
 */
void XLTables::setDisplayName(std::string_view name)
{
    auto docNode = xmlDocument().document_element();
    appendAndSetAttribute(docNode, "displayName", std::string(name));
}

/**
 * @details Get the range reference of the table
 */
std::string XLTables::rangeReference() const
{
    return xmlDocument().document_element().attribute("ref").value();
}

/**
 * @details Set the range reference of the table
 */
void XLTables::setRangeReference(std::string_view ref)
{
    auto docNode = xmlDocument().document_element();
    appendAndSetAttribute(docNode, "ref", std::string(ref));
    
    // Ensure autoFilter exists and matches ref
    auto autoFilter = appendAndGetNode(docNode, "autoFilter", TableNodeOrder);
    appendAndSetAttribute(autoFilter, "ref", std::string(ref));
}

/**
 * @details Get the table style name
 */
std::string XLTables::styleName() const
{
    return xmlDocument().document_element().child("tableStyleInfo").attribute("name").value();
}

/**
 * @details Set the table style name
 */
void XLTables::setStyleName(std::string_view styleName)
{
    auto docNode = xmlDocument().document_element();
    auto info = appendAndGetNode(docNode, "tableStyleInfo", TableNodeOrder);
    appendAndSetAttribute(info, "name", std::string(styleName));
}

/**
 * @details Check if row stripes are shown
 */
bool XLTables::showRowStripes() const
{
    return xmlDocument().document_element().child("tableStyleInfo").attribute("showRowStripes").as_bool();
}

/**
 * @details Set whether row stripes are shown
 */
void XLTables::setShowRowStripes(bool show)
{
    auto docNode = xmlDocument().document_element();
    auto info = appendAndGetNode(docNode, "tableStyleInfo", TableNodeOrder);
    appendAndSetAttribute(info, "showRowStripes", show ? "true" : "false");
}

/**
 * @details Check if column stripes are shown
 */
bool XLTables::showColumnStripes() const
{
    return xmlDocument().document_element().child("tableStyleInfo").attribute("showColumnStripes").as_bool();
}

/**
 * @details Set whether column stripes are shown
 */
void XLTables::setShowColumnStripes(bool show)
{
    auto docNode = xmlDocument().document_element();
    auto info = appendAndGetNode(docNode, "tableStyleInfo", TableNodeOrder);
    appendAndSetAttribute(info, "showColumnStripes", show ? "true" : "false");
}

/**
 * @details Check if the first column is highlighted
 */
bool XLTables::showFirstColumn() const
{
    return xmlDocument().document_element().child("tableStyleInfo").attribute("showFirstColumn").as_bool();
}

/**
 * @details Set whether the first column is highlighted
 */
void XLTables::setShowFirstColumn(bool show)
{
    auto docNode = xmlDocument().document_element();
    auto info = appendAndGetNode(docNode, "tableStyleInfo", TableNodeOrder);
    appendAndSetAttribute(info, "showFirstColumn", show ? "true" : "false");
}

/**
 * @details Check if the last column is highlighted
 */
bool XLTables::showLastColumn() const
{
    return xmlDocument().document_element().child("tableStyleInfo").attribute("showLastColumn").as_bool();
}

/**
 * @details Set whether the last column is highlighted
 */
void XLTables::setShowLastColumn(bool show)
{
    auto docNode = xmlDocument().document_element();
    auto info = appendAndGetNode(docNode, "tableStyleInfo", TableNodeOrder);
    appendAndSetAttribute(info, "showLastColumn", show ? "true" : "false");
}

/**
 * @details Append a new column to the table
 */
void XLTables::appendColumn(std::string_view name)
{
    auto docNode = xmlDocument().document_element();
    auto columns = appendAndGetNode(docNode, "tableColumns", TableNodeOrder);
    
    uint32_t count = columns.attribute("count").as_uint(0);
    count++;
    appendAndSetAttribute(columns, "count", std::to_string(count));
    
    auto column = columns.append_child("tableColumn");
    appendAndSetAttribute(column, "id", std::to_string(count));
    appendAndSetAttribute(column, "name", std::string(name));
}

/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLTables::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print(ostr); }

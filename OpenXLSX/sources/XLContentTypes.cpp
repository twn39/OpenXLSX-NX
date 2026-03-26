// ===== External Includes ===== //
#include <cstring>
#include <gsl/pointers>
#include <pugixml.hpp>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "XLContentTypes.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

namespace
{    // anonymous namespace for local functions

    /**
     * @note 2024-08-31: In line with a change to hardcoded XML relationship domains in XLRelationships.cpp, replaced the repeating
     *          hardcoded strings here with the above declared constants, preparing a potential need for a similar "dumb" fallback
     * implementation
     */
    static XLContentType GetContentTypeFromString(std::string_view typeString)
    {
        if (typeString == "application/vnd.ms-excel.sheet.macroEnabled.main+xml") return XLContentType::WorkbookMacroEnabled;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml") return XLContentType::Workbook;
        if (typeString == "application/vnd.openxmlformats-package.relationships+xml") return XLContentType::Relationships;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml") return XLContentType::Worksheet;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml") return XLContentType::Chartsheet;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml") return XLContentType::ExternalLink;
        if (typeString == "application/vnd.openxmlformats-officedocument.theme+xml") return XLContentType::Theme;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml") return XLContentType::Styles;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml") return XLContentType::SharedStrings;
        if (typeString == "application/vnd.openxmlformats-officedocument.drawing+xml") return XLContentType::Drawing;
        if (typeString == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml") return XLContentType::Chart;
        if (typeString == "application/vnd.ms-office.chartstyle+xml") return XLContentType::ChartStyle;
        if (typeString == "application/vnd.ms-office.chartcolorstyle+xml") return XLContentType::ChartColorStyle;
        if (typeString == "application/vnd.ms-excel.controlproperties+xml") return XLContentType::ControlProperties;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml") return XLContentType::CalculationChain;
        if (typeString == "application/vnd.ms-office.vbaProject") return XLContentType::VBAProject;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml") return XLContentType::PivotTable;
        if (typeString == "application/vnd.ms-excel.slicer+xml") return XLContentType::Slicer;
        if (typeString == "application/vnd.ms-excel.slicerCache+xml") return XLContentType::SlicerCache;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml") return XLContentType::PivotCacheDefinition;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml") return XLContentType::PivotCacheRecords;
        if (typeString == "application/vnd.openxmlformats-package.core-properties+xml") return XLContentType::CoreProperties;
        if (typeString == "application/vnd.openxmlformats-officedocument.extended-properties+xml") return XLContentType::ExtendedProperties;
        if (typeString == "application/vnd.openxmlformats-officedocument.custom-properties+xml") return XLContentType::CustomProperties;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml") return XLContentType::Comments;
        if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml") return XLContentType::Table;
        if (typeString == "application/vnd.openxmlformats-officedocument.vmlDrawing") return XLContentType::VMLDrawing;

        return XLContentType::Unknown;
    }

    static std::string GetStringFromType(XLContentType type)
    {
        switch (type) {
            case XLContentType::WorkbookMacroEnabled: return "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
            case XLContentType::Workbook: return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
            case XLContentType::Relationships: return "application/vnd.openxmlformats-package.relationships+xml";
            case XLContentType::Worksheet: return "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
            case XLContentType::Chartsheet: return "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
            case XLContentType::ExternalLink: return "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
            case XLContentType::Theme: return "application/vnd.openxmlformats-officedocument.theme+xml";
            case XLContentType::Styles: return "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
            case XLContentType::SharedStrings: return "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
            case XLContentType::Drawing: return "application/vnd.openxmlformats-officedocument.drawing+xml";
            case XLContentType::Chart: return "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
            case XLContentType::ChartStyle: return "application/vnd.ms-office.chartstyle+xml";
            case XLContentType::ChartColorStyle: return "application/vnd.ms-office.chartcolorstyle+xml";
            case XLContentType::ControlProperties: return "application/vnd.ms-excel.controlproperties+xml";
            case XLContentType::CalculationChain: return "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
            case XLContentType::VBAProject: return "application/vnd.ms-office.vbaProject";
            case XLContentType::PivotTable: return "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
            case XLContentType::Slicer: return "application/vnd.ms-excel.slicer+xml";
            case XLContentType::SlicerCache: return "application/vnd.ms-excel.slicerCache+xml";
            case XLContentType::PivotCacheDefinition: return "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
            case XLContentType::PivotCacheRecords: return "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
            case XLContentType::CoreProperties: return "application/vnd.openxmlformats-package.core-properties+xml";
            case XLContentType::ExtendedProperties: return "application/vnd.openxmlformats-officedocument.extended-properties+xml";
            case XLContentType::CustomProperties: return "application/vnd.openxmlformats-officedocument.custom-properties+xml";
            case XLContentType::Comments: return "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
            case XLContentType::Table: return "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
            case XLContentType::VMLDrawing: return "application/vnd.openxmlformats-officedocument.vmlDrawing";
            default: throw XLInternalError("Unknown ContentType");
        }
    }
}    // anonymous namespace

XLContentItem::XLContentItem() : m_contentNode(std::make_unique<XMLNode>()) {}

XLContentItem::XLContentItem(const XMLNode& node) : m_contentNode(std::make_unique<XMLNode>(node)) {}

XLContentItem::XLContentItem(const XLContentItem& other) : m_contentNode(std::make_unique<XMLNode>(*other.m_contentNode)) {}

XLContentItem::XLContentItem(XLContentItem&& other) noexcept = default;

XLContentItem& XLContentItem::operator=(const XLContentItem& other)
{
    if (&other != this) *m_contentNode = *other.m_contentNode;
    return *this;
}

XLContentItem& XLContentItem::operator=(XLContentItem&& other) noexcept = default;

XLContentType XLContentItem::type() const { return GetContentTypeFromString(m_contentNode->attribute("ContentType").value()); }

std::string XLContentItem::path() const { return m_contentNode->attribute("PartName").value(); }


XLContentTypes::XLContentTypes() = default;

XLContentTypes::XLContentTypes(gsl::not_null<XLXmlData*> xmlData) : XLXmlFile(xmlData) {}

XLContentTypes::XLContentTypes(const XLContentTypes& other) = default;

XLContentTypes::XLContentTypes(XLContentTypes&& other) noexcept = default;

XLContentTypes& XLContentTypes::operator=(const XLContentTypes& other) = default;

XLContentTypes& XLContentTypes::operator=(XLContentTypes&& other) noexcept = default;

bool XLContentTypes::hasDefault(std::string_view extension) const
{ return !xmlDocument().document_element().find_child_by_attribute("Default", "Extension", std::string(extension).c_str()).empty(); }

bool XLContentTypes::hasOverride(std::string_view path) const
{
    std::string internalPath;
    if (!path.empty() && path.front() != '/') {
        internalPath.reserve(path.size() + 1);
        internalPath = "/";
        internalPath.append(path);
    } else {
        internalPath = std::string(path);
    }
    return !xmlDocument().document_element().find_child_by_attribute("Override", "PartName", internalPath.c_str()).empty();
}

void XLContentTypes::addDefault(std::string_view extension, std::string_view contentType)
{
    // Check if it already exists
    const std::string extStr = std::string(extension);
    if (!xmlDocument().document_element().find_child_by_attribute("Default", "Extension", extStr.c_str()).empty()) return;

    XMLNode root = xmlDocument().document_element();
    // Default nodes should come before Override nodes.
    // Find the first Override node to insert before it, or just append if none exists.
    const XMLNode firstOverride = root.child("Override");
    XMLNode node{};
    if (firstOverride.empty())
        node = root.append_child("Default");
    else
        node = root.insert_child_before("Default", firstOverride);

    node.append_attribute("Extension").set_value(extStr.c_str());
    node.append_attribute("ContentType").set_value(std::string(contentType).c_str());
}

/**
 * @note 2024-07-22: added more intelligent whitespace support
 */
void XLContentTypes::addOverride(std::string_view path, XLContentType type)
{
    const std::string typeString   = GetStringFromType(type);
    std::string       internalPath;
    if (!path.empty() && path.front() != '/') {
        internalPath.reserve(path.size() + 1);
        internalPath = "/";
        internalPath.append(path);
    } else {
        internalPath = std::string(path);
    }

    const XMLNode lastOverride = xmlDocument().document_element().last_child_of_type(pugi::node_element);    // see if there's a last element
    XMLNode node{};                                                                                    // scope declaration

    // Create new node in the [Content_Types].xml file
    if (lastOverride.empty())
        node = xmlDocument().document_element().append_child("Override");
    else {    // if last element found
        // ===== Insert node after previous override
        node = xmlDocument().document_element().insert_child_after("Override", lastOverride);
    }
    node.append_attribute("PartName").set_value(internalPath.c_str());
    node.append_attribute("ContentType").set_value(typeString.c_str());
}

void XLContentTypes::updateOverride(std::string_view path, XLContentType type)
{
    const std::string typeString   = GetStringFromType(type);
    std::string       internalPath;
    if (!path.empty() && path.front() != '/') {
        internalPath.reserve(path.size() + 1);
        internalPath = "/";
        internalPath.append(path);
    } else {
        internalPath = std::string(path);
    }

    XMLNode node = xmlDocument().document_element().find_child_by_attribute("Override", "PartName", internalPath.c_str());
    if (!node.empty()) {
        node.attribute("ContentType").set_value(typeString.c_str());
    } else {
        addOverride(path, type);
    }
}

void XLContentTypes::deleteOverride(std::string_view path)
{ xmlDocument().document_element().remove_child(xmlDocument().document_element().find_child_by_attribute("PartName", std::string(path).c_str())); }

void XLContentTypes::deleteOverride(const XLContentItem& item) { deleteOverride(item.path()); }

XLContentItem XLContentTypes::contentItem(std::string_view path)
{ return XLContentItem(xmlDocument().document_element().find_child_by_attribute("PartName", std::string(path).c_str())); }

std::vector<XLContentItem> XLContentTypes::getContentItems()
{
    std::vector<XLContentItem> result;
    XMLNode                    item = xmlDocument().document_element().first_child_of_type(pugi::node_element);
    while (not item.empty()) {
        if (strcmp(item.name(), "Override") == 0) result.emplace_back(item);
        item = item.next_sibling_of_type(pugi::node_element);
    }

    return result;
}

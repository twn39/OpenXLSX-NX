// ===== External Includes ===== //
#include <charconv>
#include <cstdint>    // uint32_t
#include <cstring>    // strlen
#include <gsl/gsl>
#include <memory>     // std::make_unique
#include <pugixml.hpp>
#include <random>       // std::mt19937, std::random_device
#include <stdexcept>    // std::invalid_argument
#include <string>       // std::stoi, std::literals::string_literals
#include <vector>       // std::vector

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLRelationships.hpp"

#include "XLException.hpp"

using namespace OpenXLSX;

namespace
{                             // anonymous namespace: do not export these symbols
    thread_local bool RandomIDs{false};    // default: use sequential IDs
    thread_local bool RandomizerInitialized{
        false};    // will be initialized by GetNewRelsID, if XLRand32 / XLRand64 are used elsewhere, user should invoke XLInitRandom
}    // anonymous namespace

namespace OpenXLSX
{    // anonymous namespace: do not export these symbols
    /**
     * @details Set the module-local status variable RandomIDs to true
     */
    void UseRandomIDs() { RandomIDs = true; }

    /**
     * @details Set the module-local status variable RandomIDs to false
     */
    void UseSequentialIDs() { RandomIDs = false; }

    /**
     * @details Use the std::mt19937 default return on operator()
     */
    thread_local std::mt19937 Rand32(0);

    /**
     * @details Combine values from two subsequent invocations of Rand32()
     */
    uint64_t Rand64()
    {
        return (static_cast<uint64_t>(Rand32()) << 32) + Rand32();
    }    // 2024-08-18 BUGFIX: left-shift does *not* do integer promotion to long on the left operand

    /**
     * @details Initialize the module's random functions
     */
    void InitRandom(bool pseudoRandom)
    {
        uint64_t rdSeed;
        if (pseudoRandom)
            rdSeed = 3744821255L;    // in pseudo-random mode, always use the same seed
        else {
            std::random_device rd;
            rdSeed = rd();
        }
        Rand32.seed(gsl::narrow_cast<std::mt19937::result_type>(rdSeed));
        RandomizerInitialized = true;
    }
}    // namespace OpenXLSX

namespace
{
    const std::string relationshipDomainOpenXml2006          = "http://schemas.openxmlformats.org/officeDocument/2006";
    const std::string relationshipDomainOpenXml2006CoreProps = "http://schemas.openxmlformats.org/package/2006";
    const std::string relationshipDomainMicrosoft2006        = "http://schemas.microsoft.com/office/2006";
    const std::string relationshipDomainMicrosoft2011        = "http://schemas.microsoft.com/office/2011";

    /**
     * @note 2024-08-31: Included a "dumb" fallback solution in relationship tests to support
     *          previously unknown relationship domains, e.g. type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
     */
    XLRelationshipType GetRelationshipTypeFromString(std::string_view typeString)
    {
        auto matches = [typeString](std::string_view domain, std::string_view path) -> bool {
            if (typeString.length() == domain.length() + path.length() &&
                typeString.substr(0, domain.length()) == domain &&
                typeString.substr(domain.length()) == path) {
                return true;
            }
            if (typeString.length() > 15) { // length of "/relationships/"
                auto pos = typeString.find("/relationships/");
                if (pos != std::string_view::npos) return typeString.substr(pos) == path;
            }
            return false;
        };

        if (matches(relationshipDomainOpenXml2006, "/relationships/extended-properties")) return XLRelationshipType::ExtendedProperties;
        if (matches(relationshipDomainOpenXml2006, "/relationships/custom-properties")) return XLRelationshipType::CustomProperties;
        if (matches(relationshipDomainOpenXml2006, "/relationships/officeDocument")) return XLRelationshipType::Workbook;
        if (matches(relationshipDomainOpenXml2006, "/relationships/worksheet")) return XLRelationshipType::Worksheet;
        if (matches(relationshipDomainOpenXml2006, "/relationships/styles")) return XLRelationshipType::Styles;
        if (matches(relationshipDomainOpenXml2006, "/relationships/sharedStrings")) return XLRelationshipType::SharedStrings;
        if (matches(relationshipDomainOpenXml2006, "/relationships/calcChain")) return XLRelationshipType::CalculationChain;
        if (matches(relationshipDomainOpenXml2006, "/relationships/externalLink")) return XLRelationshipType::ExternalLink;
        if (matches(relationshipDomainOpenXml2006, "/relationships/theme")) return XLRelationshipType::Theme;
        if (matches(relationshipDomainOpenXml2006, "/relationships/chartsheet")) return XLRelationshipType::Chartsheet;
        if (matches(relationshipDomainOpenXml2006, "/relationships/drawing")) return XLRelationshipType::Drawing;
        if (matches(relationshipDomainOpenXml2006, "/relationships/image")) return XLRelationshipType::Image;
        if (matches(relationshipDomainOpenXml2006, "/relationships/chart")) return XLRelationshipType::Chart;
        if (matches(relationshipDomainOpenXml2006, "/relationships/externalLinkPath")) return XLRelationshipType::ExternalLinkPath;
        if (matches(relationshipDomainOpenXml2006, "/relationships/printerSettings")) return XLRelationshipType::PrinterSettings;
        if (matches(relationshipDomainOpenXml2006, "/relationships/vmlDrawing")) return XLRelationshipType::VMLDrawing;
        if (matches(relationshipDomainOpenXml2006, "/relationships/ctrlProp")) return XLRelationshipType::ControlProperties;
        if (matches(relationshipDomainOpenXml2006CoreProps, "/relationships/metadata/core-properties")) return XLRelationshipType::CoreProperties;
        if (matches(relationshipDomainMicrosoft2006, "/relationships/vbaProject")) return XLRelationshipType::VBAProject;
        if (matches(relationshipDomainOpenXml2006, "/relationships/pivotTable")) return XLRelationshipType::PivotTable;
        if (matches(relationshipDomainOpenXml2006, "/relationships/pivotCacheDefinition")) return XLRelationshipType::PivotCacheDefinition;
        if (matches(relationshipDomainOpenXml2006, "/relationships/pivotCacheRecords")) return XLRelationshipType::PivotCacheRecords;
        if (matches(relationshipDomainMicrosoft2011, "/relationships/chartStyle")) return XLRelationshipType::ChartStyle;
        if (matches(relationshipDomainMicrosoft2011, "/relationships/chartColorStyle")) return XLRelationshipType::ChartColorStyle;
        if (matches(relationshipDomainOpenXml2006, "/relationships/comments")) return XLRelationshipType::Comments;
        if (matches(relationshipDomainOpenXml2006, "/relationships/table")) return XLRelationshipType::Table;

        return XLRelationshipType::Unknown;
    }

}    //    namespace

namespace OpenXLSX_XLRelationships
{    // make GetStringFromType accessible throughout the project (for use by XLAppProperties)
    using namespace OpenXLSX;
    std::string GetStringFromType(XLRelationshipType type)
    {
        switch (type) {
            case XLRelationshipType::ExtendedProperties:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
            case XLRelationshipType::CustomProperties:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
            case XLRelationshipType::Workbook:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            case XLRelationshipType::Worksheet:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
            case XLRelationshipType::Styles:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
            case XLRelationshipType::SharedStrings:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
            case XLRelationshipType::CalculationChain:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
            case XLRelationshipType::ExternalLink:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
            case XLRelationshipType::Theme:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
            case XLRelationshipType::Chartsheet:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
            case XLRelationshipType::Drawing:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
            case XLRelationshipType::Image:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
            case XLRelationshipType::Chart:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
            case XLRelationshipType::ExternalLinkPath:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
            case XLRelationshipType::PrinterSettings:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
            case XLRelationshipType::VMLDrawing:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
            case XLRelationshipType::ControlProperties:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
            case XLRelationshipType::CoreProperties:
                return "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
            case XLRelationshipType::VBAProject:
                return "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
            case XLRelationshipType::PivotTable:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
            case XLRelationshipType::PivotCacheDefinition:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
            case XLRelationshipType::PivotCacheRecords:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
            case XLRelationshipType::ChartStyle:
                return "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
            case XLRelationshipType::ChartColorStyle:
                return "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
            case XLRelationshipType::Comments:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
            case XLRelationshipType::Table:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
            case XLRelationshipType::Hyperlink:
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
            default:
                throw XLInternalError("RelationshipType not recognized!");
        }
    }
}    //    namespace OpenXLSX_XLRelationships

namespace
{    //    re-open anonymous namespace
    /**
     * @brief Get a new, unique relationship ID
     * @param relationshipsNode the XML node that contains the document relationships
     * @return A 64 bit integer with a relationship ID
     */
    uint64_t GetNewRelsID(XMLNode relationshipsNode)
    {
        if (RandomIDs) {
            if (!RandomizerInitialized) InitRandom();
            return Rand64();
        }
        
        XMLNode  relationship = relationshipsNode.first_child_of_type(pugi::node_element);
        uint64_t newId        = 1;
        while (!relationship.empty()) {
            std::string_view idVal = relationship.attribute("Id").value();
            if (idVal.length() > 3) { // skip "rId"
                uint64_t id;
                auto [ptr, ec] = std::from_chars(idVal.data() + 3, idVal.data() + idVal.size(), id);
                if (ec == std::errc()) {
                    if (id >= newId) newId = id + 1;
                }
            }
            relationship = relationship.next_sibling_of_type(pugi::node_element);
        }
        return newId;
    }

    /**
     * @brief Get a string representation of GetNewRelsID
     */
    std::string GetNewRelsIDString(XMLNode relationshipsNode)
    {
        uint64_t newId = GetNewRelsID(relationshipsNode);
        if (RandomIDs) return "R" + BinaryAsHexString(gsl::make_span(reinterpret_cast<const std::byte*>(&newId), sizeof(newId)));
        return "rId" + std::to_string(newId);
    }
}    // anonymous namespace

XLRelationshipItem::XLRelationshipItem(const XMLNode& node) : m_relationshipNode(node) {}

XLRelationshipType XLRelationshipItem::type() const { return GetRelationshipTypeFromString(m_relationshipNode.attribute("Type").value()); }

std::string XLRelationshipItem::target() const { return m_relationshipNode.attribute("Target").value(); }

std::string XLRelationshipItem::id() const { return m_relationshipNode.attribute("Id").value(); }

bool XLRelationshipItem::empty() const { return m_relationshipNode.empty(); }

/**
 * @details Creates a XLRelationships object. The pathTo will be verified and stored 
 * to resolve relationship targets as absolute paths within the XLSX archive.
 */
XLRelationships::XLRelationships(gsl::not_null<XLXmlData*> xmlData, std::string pathTo) : XLXmlFile(xmlData)
{
    constexpr std::string_view relFolder = "_rels/";    // all relationships are stored in a (sub-)folder named "_rels/"
    constexpr size_t           relFolderLen = relFolder.length();

    const bool   addFirstSlash = (pathTo[0] != '/');    // if first character of pathTo is NOT a slash, then addFirstSlash = true
    const size_t pathToEndsAt  = pathTo.find_last_of('/');
    if ((pathToEndsAt != std::string::npos) and (pathToEndsAt + 1 >= relFolderLen) &&
        (pathTo.substr(pathToEndsAt + 1 - relFolderLen, relFolderLen) == relFolder))    // nominal
        m_path = (addFirstSlash ? "/" : "") + pathTo.substr(0, pathToEndsAt - relFolderLen + 1);
    else {
        using namespace std::literals::string_literals;
        throw XLException("XLRelationships constructor: pathTo \""s + pathTo + "\" does not point to a file in a folder named \"" +
                          std::string(relFolder) + "\""s);
    }

    XMLDocument& doc = xmlDocument();
    if (doc.document_element().empty())    // handle a bad (no document element) relationships XML file
        doc.load_string("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"\n"
                        "</Relationships>",
                        pugi_parse_settings);
}

XLRelationships::~XLRelationships() = default;

XLRelationshipItem XLRelationships::relationshipById(std::string_view id) const
{ return XLRelationshipItem(xmlDocument().document_element().find_child_by_attribute("Id", std::string(id).c_str())); }

/**
 * @details Returns the XLRelationshipItem with the requested target.
 * Performance: Normalizes paths before comparison to ensure consistency between relative and absolute references.
 */
XLRelationshipItem XLRelationships::relationshipByTarget(std::string_view target, bool throwIfNotFound) const
{
    // Rationale: Convert to absolute path and normalize to resolve . and .. entries for consistent comparison.
    const std::string absoluteTarget = eliminateDotAndDotDotFromPath(!target.empty() && target.front() == '/' ? target : m_path + std::string(target));

    XMLNode relationshipNode = xmlDocument().document_element().first_child_of_type(pugi::node_element);
    while (!relationshipNode.empty()) {
        std::string relationTarget = relationshipNode.attribute("Target").value();
        if (relationTarget.empty() || relationTarget.front() != '/') relationTarget = m_path + relationTarget;
        relationTarget = eliminateDotAndDotDotFromPath(relationTarget);

        if (absoluteTarget == relationTarget) return XLRelationshipItem(relationshipNode);
        relationshipNode = relationshipNode.next_sibling_of_type(pugi::node_element);
    }

    if (throwIfNotFound) {
        using namespace std::literals::string_literals;
        throw XLException("XLRelationships::"s + __func__ + ": relationship with target \""s + std::string(target) + "\" (absolute: \""s +
                          absoluteTarget + "\") does not exist!"s);
    }
    return XLRelationshipItem();
}

std::vector<XLRelationshipItem> XLRelationships::relationships() const
{
    std::vector<XLRelationshipItem> result;
    XMLNode item = xmlDocument().document_element().first_child_of_type(pugi::node_element);
    while (!item.empty()) {
        result.emplace_back(item);
        item = item.next_sibling_of_type(pugi::node_element);
    }
    return result;
}

void XLRelationships::deleteRelationship(std::string_view relID)
{ xmlDocument().document_element().remove_child(xmlDocument().document_element().find_child_by_attribute("Id", std::string(relID).c_str())); }

void XLRelationships::deleteRelationship(const XLRelationshipItem& item) { deleteRelationship(item.id()); }

/**
 * @details Adds a new relationship. Handles whitespace preservation to maintain consistent XML formatting.
 */
XLRelationshipItem XLRelationships::addRelationship(XLRelationshipType type, std::string_view target, bool isExternal)
{
    const std::string typeString = OpenXLSX_XLRelationships::GetStringFromType(type);
    const std::string id = GetNewRelsIDString(xmlDocument().document_element());

    XMLNode lastRelationship = xmlDocument().document_element().last_child_of_type(pugi::node_element);
    XMLNode node{};

    if (lastRelationship.empty())
        node = xmlDocument().document_element().prepend_child("Relationship");
    else {
        node = xmlDocument().document_element().insert_child_after("Relationship", lastRelationship);

        // Preserve formatting by copying preceding whitespace
        XMLNode copyWhitespaceFrom = lastRelationship;
        XMLNode insertBefore       = node;
        while (copyWhitespaceFrom.previous_sibling().type() == pugi::node_pcdata) {
            copyWhitespaceFrom = copyWhitespaceFrom.previous_sibling();
            insertBefore = xmlDocument().document_element().insert_child_before(pugi::node_pcdata, insertBefore);
            insertBefore.set_value(copyWhitespaceFrom.value());
        }
    }
    node.append_attribute("Id").set_value(id.c_str());
    node.append_attribute("Type").set_value(typeString.c_str());
    node.append_attribute("Target").set_value(std::string(target).c_str());

    if (isExternal || type == XLRelationshipType::ExternalLinkPath) { node.append_attribute("TargetMode").set_value("External"); }

    return XLRelationshipItem(node);
}

bool XLRelationships::targetExists(std::string_view target) const
{
    constexpr bool DO_NOT_THROW = false;
    return !relationshipByTarget(target, DO_NOT_THROW).empty();
}

bool XLRelationships::idExists(std::string_view id) const
{ return !xmlDocument().document_element().find_child_by_attribute("Id", std::string(id).c_str()).empty(); }

/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLRelationships::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print(ostr); }

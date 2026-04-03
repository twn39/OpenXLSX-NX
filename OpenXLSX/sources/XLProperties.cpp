// ===== External Includes ===== //
#include <fmt/format.h>
#include <pugixml.hpp>
#include <string>
#include <string_view>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLProperties.hpp"
#include "XLRelationships.hpp"    // GetStringFromType
#include "XLUtilities.hpp"

using namespace OpenXLSX;

namespace    // anonymous namespace for module local functions
{
    /**
     * @brief insert node with nodeNameToInsert before insertBeforeThis and copy all the consecutive whitespace nodes preceeding
     * insertBeforeThis
     * @param nodeNameToInsert create a new element node with this name
     * @param insertBeforeThis self-explanatory, isn't it? :)
     * @returns the inserted XMLNode
     * @throws XLInternalError if insertBeforeThis has no parent node
     */
    XMLNode prettyInsertXmlNodeBefore(std::string_view nodeNameToInsert, XMLNode insertBeforeThis)
    {
        XMLNode parentNode = insertBeforeThis.parent();
        if (parentNode.empty()) throw XLInternalError("prettyInsertXmlNodeBefore can not insert before a node that has no parent");

        // ===== Find potential whitespaces preceeding insertBeforeThis
        XMLNode whitespaces = insertBeforeThis;
        while (whitespaces.previous_sibling().type() == pugi::node_pcdata) whitespaces = whitespaces.previous_sibling();

        // ===== Insert the nodeNameToInsert before insertBeforeThis, "hijacking" the preceeding whitespace nodes
        XMLNode insertedNode = parentNode.insert_child_before(std::string(nodeNameToInsert).c_str(), insertBeforeThis);
        // ===== Copy all hijacked whitespaces to also preceed (again) the node insertBeforeThis
        while (whitespaces.type() == pugi::node_pcdata) {
            parentNode.insert_copy_before(whitespaces, insertBeforeThis);
            whitespaces = whitespaces.next_sibling();
        }
        return insertedNode;
    }

    /**
     * @brief fetch the <HeadingPairs><vt:vector> node from docNode, create if not existing
     * @param docNode the xml document from which to fetch
     * @returns an XMLNode for the <vt:vector> child of <HeadingPairs>
     * @throws XLInternalError in case either XML element does not exist and fails to create
     */
    XMLNode headingPairsNode(XMLNode docNode)
    {
        XMLNode headingPairs = docNode.child("HeadingPairs");
        if (headingPairs.empty()) {
            XMLNode titlesOfParts =
                docNode.child("TitlesOfParts");    // attempt to locate TitlesOfParts to insert HeadingPairs right before
            if (titlesOfParts.empty())
                headingPairs = docNode.append_child("HeadingPairs");
            else
                headingPairs = prettyInsertXmlNodeBefore("HeadingPairs", titlesOfParts);
            if (headingPairs.empty()) throw XLInternalError("headingPairsNode was unable to create XML element HeadingPairs");
        }
        XMLNode node = headingPairs.first_child_of_type(pugi::node_element);
        if (node.empty()) {    // heading pairs child "vt:vector" does not exist
            node = headingPairs.append_child("vt:vector", XLForceNamespace);
            if (node.empty()) throw XLInternalError("headingPairsNode was unable to create HeadingPairs child element <vt:vector>");
            node.append_attribute("size")     = 0;    // NOTE: size counts each heading pair as 2 - but for now, there are no entries
            node.append_attribute("baseType") = "variant";
        }
        // ===== Finally, return what should be a good "vt:vector" node with all the heading pairs
        return node;
    }

    XMLAttribute headingPairsSize(XMLNode docNode)
    { return docNode.child("HeadingPairs").first_child_of_type(pugi::node_element).attribute("size"); }

    /**
     * @brief fetch the <TitlesOfParts><vt:vector> XML node, create if not existing
     * @param docNode the xml document from which to fetch
     * @returns an XMLNode for the <vt:vector> child of <TitlesOfParts>
     * @throws XLInternalError in case either XML element does not exist and fails to create
     */
    XMLNode sheetNames(XMLNode docNode)
    {
        XMLNode titlesOfParts = docNode.child("TitlesOfParts");
        if (titlesOfParts.empty()) {
            // ===== Attempt to find a good position to insert TitlesOfParts
            XMLNode node = docNode.first_child_of_type(pugi::node_element);
            while (not node.empty()) {
                std::string_view nodeName = node.name();
                if (nodeName == "Application" || nodeName == "AppVersion" || nodeName == "DocSecurity" || nodeName == "ScaleCrop" ||
                    nodeName == "Template" || nodeName == "TotalTime" || nodeName == "HeadingPairs")
                {
                    node = node.next_sibling_of_type(pugi::node_element);
                }
                else {
                    break;
                }
            }

            if (node.empty())
                titlesOfParts = docNode.append_child("TitlesOfParts");
            else
                titlesOfParts = prettyInsertXmlNodeBefore("TitlesOfParts", node);
        }
        XMLNode vtVector = titlesOfParts.first_child_of_type(pugi::node_element);
        if (vtVector.empty()) {
            vtVector = titlesOfParts.prepend_child("vt:vector", XLForceNamespace);
            if (vtVector.empty()) throw XLInternalError("sheetNames was unable to create TitlesOfParts child element <vt:vector>");
            vtVector.append_attribute("size")     = 0;
            vtVector.append_attribute("baseType") = "lpstr";
        }
        return vtVector;
    }

    XMLAttribute sheetCount(XMLNode docNode) { return sheetNames(docNode).attribute("size"); }

    /**
     * @brief from the HeadingPairs node, look up name and return the next sibling (which should be the associated value node)
     * @param docNode the XML document giving access to the <HeadingPairs> node
     * @param name the value pair whose attribute to return
     * @returns XMLNode where the value associated with name is stored (the next element sibling from the node matching name)
     * @returns XMLNode{} (empty) if name is not found or has no next element sibling
     */
    XMLNode getHeadingPairsValue(XMLNode docNode, std::string_view name)
    {
        XMLNode item = headingPairsNode(docNode).first_child_of_type(pugi::node_element);
        while (not item.empty() && std::string_view(item.first_child_of_type(pugi::node_element).child_value()) != name)
            item = item.next_sibling_of_type(pugi::node_element)
                       .next_sibling_of_type(pugi::node_element);    // advance two elements to skip count node
        if (not item.empty())                                        // if name was found
            item = item.next_sibling_of_type(pugi::node_element);    // advance once more to the value node

        // ===== Return the first element child of the <vt:variant> node containing the heading pair value
        //       (returns an empty node if item or first_child_of_type is empty)
        return item.first_child_of_type(pugi::node_element);
    }
}    // anonymous namespace

/**
 * @details
 */
void XLProperties::createFromTemplate()
{
    if (m_xmlData == nullptr) throw XLInternalError("XLProperties m_xmlData is nullptr");

    std::string       xmlns = OpenXLSX_XLRelationships::GetStringFromType(XLRelationshipType::CoreProperties);
    const std::string rels  = "/relationships/";
    size_t            pos   = xmlns.find(rels);
    if (pos != std::string::npos)
        xmlns.replace(pos, rels.size(), "/");
    else
        xmlns = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

    XMLNode props                            = xmlDocument().prepend_child("cp:coreProperties");
    props.append_attribute("xmlns:cp")       = xmlns.c_str();
    props.append_attribute("xmlns:dc")       = "http://purl.org/dc/elements/1.1/";
    props.append_attribute("xmlns:dcterms")  = "http://purl.org/dc/terms/";
    props.append_attribute("xmlns:dcmitype") = "http://purl.org/dc/dcmitype/";
    props.append_attribute("xmlns:xsi")      = "http://www.w3.org/2001/XMLSchema-instance";

    // Set default values
    setTitle("");
    setCreator("OpenXLSX");
    setLastModifiedBy("OpenXLSX");

    XLDateTime now = XLDateTime::now();
    setCreated(now);
    setModified(now);
}

/**
 * @details
 */
XLProperties::XLProperties(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
    if (xmlDocument().document_element().empty()) createFromTemplate();
}

/**
 * @details
 */
XLProperties::~XLProperties() = default;

/**
 * @details
 */
void XLProperties::setProperty(std::string_view name, std::string_view value)
{
    if (m_xmlData == nullptr) return;
    XMLNode node = xmlDocument().document_element().child(std::string(name).c_str());
    if (node.empty()) node = xmlDocument().document_element().append_child(std::string(name).c_str());

    node.text().set(std::string(value).c_str());
}

/**
 * @details
 */
void XLProperties::setProperty(std::string_view name, int value) { setProperty(name, fmt::format("{}", value)); }

/**
 * @details
 */
void XLProperties::setProperty(std::string_view name, double value) { setProperty(name, fmt::format("{}", value)); }

/**
 * @details
 */
std::string XLProperties::property(std::string_view name) const
{
    if (m_xmlData == nullptr) return "";
    XMLNode property = xmlDocument().document_element().child(std::string(name).c_str());
    if (property.empty()) return "";

    return property.text().get();
}

/**
 * @details
 */
void XLProperties::deleteProperty(std::string_view name)
{
    if (m_xmlData == nullptr) return;
    if (const XMLNode property = xmlDocument().document_element().child(std::string(name).c_str()); not property.empty())
        xmlDocument().document_element().remove_child(property);
}

// Standard Core Properties Implementations
std::string XLProperties::title() const { return property("dc:title"); }
void        XLProperties::setTitle(std::string_view title) { setProperty("dc:title", title); }

std::string XLProperties::subject() const { return property("dc:subject"); }
void        XLProperties::setSubject(std::string_view subject) { setProperty("dc:subject", subject); }

std::string XLProperties::creator() const { return property("dc:creator"); }
void        XLProperties::setCreator(std::string_view creator) { setProperty("dc:creator", creator); }

std::string XLProperties::keywords() const { return property("cp:keywords"); }
void        XLProperties::setKeywords(std::string_view keywords) { setProperty("cp:keywords", keywords); }

std::string XLProperties::description() const { return property("dc:description"); }
void        XLProperties::setDescription(std::string_view description) { setProperty("dc:description", description); }

std::string XLProperties::lastModifiedBy() const { return property("cp:lastModifiedBy"); }
void        XLProperties::setLastModifiedBy(std::string_view author) { setProperty("cp:lastModifiedBy", author); }

std::string XLProperties::revision() const { return property("cp:revision"); }
void        XLProperties::setRevision(std::string_view revision) { setProperty("cp:revision", revision); }

std::string XLProperties::lastPrinted() const { return property("cp:lastPrinted"); }
void        XLProperties::setLastPrinted(std::string_view date) { setProperty("cp:lastPrinted", date); }

XLDateTime XLProperties::created() const
{
    std::string date = property("dcterms:created");
    if (date.empty()) return XLDateTime();
    return XLDateTime::fromString(date, "%Y-%m-%dT%H:%M:%SZ");
}
void XLProperties::setCreated(const XLDateTime& date)
{
    auto node = xmlDocument().document_element().child("dcterms:created");
    if (node.empty()) node = xmlDocument().document_element().append_child("dcterms:created");
    appendAndSetAttribute(node, "xsi:type", "dcterms:W3CDTF");
    node.text().set(date.toString("%Y-%m-%dT%H:%M:%SZ").c_str());
}

XLDateTime XLProperties::modified() const
{
    std::string date = property("dcterms:modified");
    if (date.empty()) return XLDateTime();
    return XLDateTime::fromString(date, "%Y-%m-%dT%H:%M:%SZ");
}
void XLProperties::setModified(const XLDateTime& date)
{
    auto node = xmlDocument().document_element().child("dcterms:modified");
    if (node.empty()) node = xmlDocument().document_element().append_child("dcterms:modified");
    appendAndSetAttribute(node, "xsi:type", "dcterms:W3CDTF");
    node.text().set(date.toString("%Y-%m-%dT%H:%M:%SZ").c_str());
}

std::string XLProperties::category() const { return property("cp:category"); }
void        XLProperties::setCategory(std::string_view category) { setProperty("cp:category", category); }

std::string XLProperties::contentStatus() const { return property("cp:contentStatus"); }
void        XLProperties::setContentStatus(std::string_view status) { setProperty("cp:contentStatus", status); }

/**
 * @details
 */
void XLAppProperties::createFromTemplate(XMLDocument const& workbookXml)
{
    if (m_xmlData == nullptr) throw XLInternalError("XLAppProperties m_xmlData is nullptr");

    std::map<uint32_t, std::string> sheetsOrderedById;

    auto sheet = workbookXml.document_element().child("sheets").first_child_of_type(pugi::node_element);
    while (not sheet.empty()) {
        std::string sheetName = sheet.attribute("name").as_string();
        uint32_t    sheetId   = sheet.attribute("sheetId").as_uint();
        sheetsOrderedById.insert(std::pair<uint32_t, std::string>(sheetId, sheetName));
        sheet = sheet.next_sibling_of_type();
    }

    uint32_t worksheetCount = 0;
    for (const auto& [key, value] : sheetsOrderedById) {
        if (key != ++worksheetCount)
            throw XLInputError("xl/workbook.xml is missing sheet with sheetId=\"" + std::to_string(worksheetCount) + "\"");
    }

    std::string       xmlns = OpenXLSX_XLRelationships::GetStringFromType(XLRelationshipType::ExtendedProperties);
    const std::string rels  = "/relationships/";
    size_t            pos   = xmlns.find(rels);
    if (pos != std::string::npos)
        xmlns.replace(pos, rels.size(), "/");
    else
        xmlns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    XMLNode props                      = xmlDocument().prepend_child("Properties");
    props.append_attribute("xmlns")    = xmlns.c_str();
    props.append_attribute("xmlns:vt") = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    props.append_child("Application").text().set("OpenXLSX");
    props.append_child("DocSecurity").text().set(0);
    props.append_child("ScaleCrop").text().set(false);

    XMLNode headingPairs               = props.append_child("HeadingPairs");
    XMLNode vecHP                      = headingPairs.append_child("vt:vector", XLForceNamespace);
    vecHP.append_attribute("size")     = 2;
    vecHP.append_attribute("baseType") = "variant";
    vecHP.append_child("vt:variant", XLForceNamespace).append_child("vt:lpstr", XLForceNamespace).text().set("Worksheets");
    vecHP.append_child("vt:variant", XLForceNamespace).append_child("vt:i4", XLForceNamespace).text().set(worksheetCount);

    XMLNode sheetsVector                      = props.append_child("TitlesOfParts").append_child("vt:vector", XLForceNamespace);
    sheetsVector.append_attribute("size")     = worksheetCount;
    sheetsVector.append_attribute("baseType") = "lpstr";
    for (const auto& [key, value] : sheetsOrderedById) sheetsVector.append_child("vt:lpstr", XLForceNamespace).text().set(value.c_str());

    props.append_child("Company");
    props.append_child("LinksUpToDate").text().set(false);
    props.append_child("SharedDoc").text().set(false);
    props.append_child("HyperlinksChanged").text().set(false);
    props.append_child("AppVersion").text().set("1.7");
}

/**
 * @details
 */
XLAppProperties::XLAppProperties(XLXmlData* xmlData, XMLDocument const& workbookXml) : XLXmlFile(xmlData)
{
    if (xmlDocument().document_element().empty()) createFromTemplate(workbookXml);
}

/**
 * @details
 */
XLAppProperties::XLAppProperties(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

/**
 * @details
 */
XLAppProperties::~XLAppProperties() = default;

/**
 * @details
 */
void XLAppProperties::incrementSheetCount(int16_t increment)
{
    int32_t newCount = sheetCount(xmlDocument().document_element()).as_int() + increment;
    if (newCount < 1) throw XLInternalError("must not decrement sheet count below 1");
    sheetCount(xmlDocument().document_element()).set_value(newCount);
    XMLNode headingPairWorksheetsValue = getHeadingPairsValue(xmlDocument().document_element(), "Worksheets");
    if (headingPairWorksheetsValue.empty())    // seems heading pair does not exist
        addHeadingPair("Worksheets", newCount);
    else
        headingPairWorksheetsValue.text().set(newCount);
}

/**
 * @details
 */
void XLAppProperties::alignWorksheets(std::vector<std::string> const& workbookSheetNames)
{
    XMLNode titlesOfParts = sheetNames(xmlDocument().document_element());
    XMLNode sheetName     = titlesOfParts.first_child_of_type(pugi::node_element);
    for (const std::string& workbookSheetName : workbookSheetNames) {
        if (sheetName.empty()) sheetName = titlesOfParts.append_child("vt:lpstr", XLForceNamespace);
        sheetName.text().set(workbookSheetName.c_str());
        sheetName = sheetName.next_sibling_of_type(pugi::node_element);
    }

    if (not sheetName.empty()) {
        while (sheetName.previous_sibling().type() == pugi::node_pcdata) titlesOfParts.remove_child(sheetName.previous_sibling());
        XMLNode lastSheetName = titlesOfParts.last_child_of_type(pugi::node_element);
        if (sheetName != lastSheetName) {
            while (sheetName.next_sibling() != lastSheetName) titlesOfParts.remove_child(sheetName.next_sibling());
            titlesOfParts.remove_child(lastSheetName);
        }
        titlesOfParts.remove_child(sheetName);
    }

    sheetCount(xmlDocument().document_element()).set_value(workbookSheetNames.size());
    incrementSheetCount(0);
}

/**
 * @details
 */
void XLAppProperties::addSheetName(std::string_view title)
{
    if (m_xmlData == nullptr) return;
    XMLNode theNode = sheetNames(xmlDocument().document_element()).append_child("vt:lpstr", XLForceNamespace);
    theNode.text().set(std::string(title).c_str());
    incrementSheetCount(+1);
}

/**
 * @details
 */
void XLAppProperties::deleteSheetName(std::string_view title)
{
    if (m_xmlData == nullptr) return;
    XMLNode theNode = sheetNames(xmlDocument().document_element()).first_child_of_type(pugi::node_element);
    while (not theNode.empty()) {
        if (std::string_view(theNode.child_value()) == title) {
            sheetNames(xmlDocument().document_element()).remove_child(theNode);
            incrementSheetCount(-1);
            return;
        }
        theNode = theNode.next_sibling_of_type(pugi::node_element);
    }
}

/**
 * @details
 */
void XLAppProperties::setSheetName(std::string_view oldTitle, std::string_view newTitle)
{
    if (m_xmlData == nullptr) return;
    XMLNode theNode = sheetNames(xmlDocument().document_element()).first_child_of_type(pugi::node_element);
    while (not theNode.empty()) {
        if (std::string_view(theNode.child_value()) == oldTitle) {
            theNode.text().set(std::string(newTitle).c_str());
            return;
        }
        theNode = theNode.next_sibling_of_type(pugi::node_element);
    }
}

/**
 * @details
 */
void XLAppProperties::addHeadingPair(std::string_view name, int value)
{
    if (m_xmlData == nullptr) return;

    XMLNode HeadingPairsNode = headingPairsNode(xmlDocument().document_element());
    XMLNode item             = HeadingPairsNode.first_child_of_type(pugi::node_element);
    while (not item.empty() && std::string_view(item.first_child_of_type(pugi::node_element).child_value()) != name)
        item = item.next_sibling_of_type(pugi::node_element).next_sibling_of_type(pugi::node_element);

    XMLNode pairCategory = item;
    XMLNode pairValue{};
    if (not pairCategory.empty())
        pairValue = pairCategory.next_sibling_of_type(pugi::node_element).first_child_of_type(pugi::node_element);
    else {
        item = HeadingPairsNode.last_child_of_type(pugi::node_element);
        if (not item.empty())
            pairCategory = HeadingPairsNode.insert_child_after("vt:variant", item, XLForceNamespace);
        else
            pairCategory = HeadingPairsNode.append_child("vt:variant", XLForceNamespace);
        XMLNode categoryName = pairCategory.append_child("vt:lpstr", XLForceNamespace);
        categoryName.text().set(std::string(name).c_str());
        XMLNode pairValueParent = HeadingPairsNode.insert_child_after("vt:variant", pairCategory, XLForceNamespace);
        pairValue               = pairValueParent.append_child("vt:i4", XLForceNamespace);
    }

    if (not pairValue.empty())
        pairValue.text().set(fmt::format("{}", value).c_str());
    else {
        throw XLInternalError(fmt::format("XLAppProperties::addHeadingPair: found no matching pair count value to name {}", name));
    }
    headingPairsSize(xmlDocument().document_element()).set_value(HeadingPairsNode.child_count_of_type());
}

/**
 * @details
 */
void XLAppProperties::deleteHeadingPair(std::string_view name)
{
    if (m_xmlData == nullptr) return;

    XMLNode HeadingPairsNode = headingPairsNode(xmlDocument().document_element());
    XMLNode item             = HeadingPairsNode.first_child_of_type(pugi::node_element);
    while (not item.empty() && std::string_view(item.first_child_of_type(pugi::node_element).child_value()) != name)
        item = item.next_sibling_of_type(pugi::node_element).next_sibling_of_type(pugi::node_element);

    if (not item.empty()) {
        const XMLNode count = item.next_sibling_of_type(pugi::node_element);
        if (not count.empty()) {
            while (item.next_sibling() != count) HeadingPairsNode.remove_child(item.next_sibling());
            while ((not count.next_sibling().empty()) and (count.next_sibling().type() != pugi::node_element))
                HeadingPairsNode.remove_child(count.next_sibling());
            HeadingPairsNode.remove_child(count);
        }
        HeadingPairsNode.remove_child(item);
        headingPairsSize(xmlDocument().document_element()).set_value(HeadingPairsNode.child_count_of_type());
    }
}

/**
 * @details
 */
void XLAppProperties::setHeadingPair(std::string_view name, int newValue)
{
    if (m_xmlData == nullptr) return;

    XMLNode pairValue = getHeadingPairsValue(xmlDocument().document_element(), name);
    if (not pairValue.empty() && std::string_view(pairValue.name()) == "vt:i4")
        pairValue.text().set(fmt::format("{}", newValue).c_str());
    else
        throw XLInternalError(fmt::format("XLAppProperties::setHeadingPair: found no matching pair count value to name {}", name));
}

/**
 * @details
 */
void XLAppProperties::setProperty(std::string_view name, std::string_view value)
{
    if (m_xmlData == nullptr) return;
    auto property = xmlDocument().document_element().child(std::string(name).c_str());
    if (property.empty()) property = xmlDocument().document_element().append_child(std::string(name).c_str());
    property.text().set(std::string(value).c_str());
}

/**
 * @details
 */
std::string XLAppProperties::property(std::string_view name) const
{
    if (m_xmlData == nullptr) return "";
    XMLNode property = xmlDocument().document_element().child(std::string(name).c_str());
    if (property.empty()) return "";
    return property.text().get();
}

/**
 * @details
 */
void XLAppProperties::deleteProperty(std::string_view name)
{
    if (m_xmlData == nullptr) return;
    const auto property = xmlDocument().document_element().child(std::string(name).c_str());
    if (property.empty()) return;
    xmlDocument().document_element().remove_child(property);
}

/**
 * @details
 */
void XLAppProperties::appendSheetName(std::string_view sheetName)
{
    if (m_xmlData == nullptr) return;
    auto theNode = sheetNames(xmlDocument().document_element()).append_child("vt:lpstr", XLForceNamespace);
    theNode.text().set(std::string(sheetName).c_str());
    incrementSheetCount(+1);
}

/**
 * @details
 */
void XLAppProperties::prependSheetName(std::string_view sheetName)
{
    if (m_xmlData == nullptr) return;
    auto theNode = sheetNames(xmlDocument().document_element()).prepend_child("vt:lpstr", XLForceNamespace);
    theNode.text().set(std::string(sheetName).c_str());
    incrementSheetCount(+1);
}

/**
 * @details
 */
void XLAppProperties::insertSheetName(std::string_view sheetName, unsigned int index)
{
    if (m_xmlData == nullptr) return;

    if (index <= 1) {
        prependSheetName(sheetName);
        return;
    }

    if (index > sheetCount(xmlDocument().document_element()).as_uint()) {
        const XMLNode lastSheet = sheetNames(xmlDocument().document_element()).last_child_of_type(pugi::node_element);
        if (not lastSheet.empty()) {
            XMLNode theNode = sheetNames(xmlDocument().document_element()).insert_child_after("vt:lpstr", lastSheet, XLForceNamespace);
            theNode.text().set(std::string(sheetName).c_str());
            incrementSheetCount(+1);
        }
        else
            appendSheetName(sheetName);
        return;
    }

    XMLNode  curNode = sheetNames(xmlDocument().document_element()).first_child_of_type(pugi::node_element);
    unsigned idx     = 1;
    while (not curNode.empty()) {
        if (idx == index) break;
        curNode = curNode.next_sibling_of_type(pugi::node_element);
        ++idx;
    }

    if (curNode.empty()) {
        appendSheetName(sheetName);
        return;
    }

    XMLNode theNode = sheetNames(xmlDocument().document_element()).insert_child_before("vt:lpstr", curNode, XLForceNamespace);
    theNode.text().set(std::string(sheetName).c_str());

    incrementSheetCount(+1);
}

namespace
{
    XMLNode findOrCreateProperty(XMLNode root, std::string_view name)
    {
        for (auto p : root.children("property")) {
            if (std::string_view(p.attribute("name").value()) == name) return XMLNode(p);
        }

        int maxPid = 1;
        for (auto p : root.children("property")) {
            int pid = p.attribute("pid").as_int();
            if (pid > maxPid) maxPid = pid;
        }

        auto prop                      = root.append_child("property");
        prop.append_attribute("fmtid") = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
        prop.append_attribute("pid")   = maxPid + 1;
        prop.append_attribute("name")  = std::string(name).c_str();
        return prop;
    }

    void setPropertyValueNode(XMLNode prop, std::string_view type, std::string_view value)
    {
        XMLNode valueNode = prop.first_child_of_type(pugi::node_element);
        if (valueNode.empty())
            valueNode = prop.append_child(std::string(type).c_str());
        else if (std::string_view(valueNode.name()) != type) {
            prop.remove_child(valueNode);
            valueNode = prop.append_child(std::string(type).c_str());
        }
        valueNode.text().set(std::string(value).c_str());
    }
}    // namespace

/**
 * @details
 */
void XLCustomProperties::createFromTemplate()
{
    if (m_xmlData == nullptr) throw XLInternalError("XLCustomProperties m_xmlData is nullptr");

    XMLNode props                      = xmlDocument().prepend_child("Properties");
    props.append_attribute("xmlns")    = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    props.append_attribute("xmlns:vt") = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
}

/**
 * @details
 */
XLCustomProperties::XLCustomProperties(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
    if (xmlDocument().document_element().empty()) createFromTemplate();
}

/**
 * @details
 */
XLCustomProperties::~XLCustomProperties() = default;

/**
 * @details
 */
void XLCustomProperties::setProperty(std::string_view name, std::string_view value)
{
    if (m_xmlData == nullptr) return;
    auto prop = findOrCreateProperty(xmlDocument().document_element(), name);
    setPropertyValueNode(prop, "vt:lpwstr", value);
}

/**
 * @details
 */
void XLCustomProperties::setProperty(std::string_view name, const char* value) { setProperty(name, std::string_view(value)); }

/**
 * @details
 */
void XLCustomProperties::setProperty(std::string_view name, int value)
{
    if (m_xmlData == nullptr) return;
    auto prop = findOrCreateProperty(xmlDocument().document_element(), name);
    setPropertyValueNode(prop, "vt:i4", fmt::format("{}", value));
}

/**
 * @details
 */
void XLCustomProperties::setProperty(std::string_view name, double value)
{
    if (m_xmlData == nullptr) return;
    auto prop = findOrCreateProperty(xmlDocument().document_element(), name);
    setPropertyValueNode(prop, "vt:r8", fmt::format("{}", value));
}

/**
 * @details
 */
void XLCustomProperties::setProperty(std::string_view name, bool value)
{
    if (m_xmlData == nullptr) return;
    auto prop = findOrCreateProperty(xmlDocument().document_element(), name);
    setPropertyValueNode(prop, "vt:bool", value ? "true" : "false");
}

/**
 * @details
 */
void XLCustomProperties::setProperty(std::string_view name, const XLDateTime& value)
{
    if (m_xmlData == nullptr) return;
    auto prop = findOrCreateProperty(xmlDocument().document_element(), name);
    setPropertyValueNode(prop, "vt:filetime", value.toString("%Y-%m-%dT%H:%M:%SZ"));
}

/**
 * @details
 */
std::string XLCustomProperties::property(std::string_view name) const
{
    if (m_xmlData == nullptr) return "";
    for (auto p : xmlDocument().document_element().children("property")) {
        if (std::string_view(p.attribute("name").value()) == name) {
            XMLNode valueNode = XMLNode(p).first_child_of_type(pugi::node_element);
            return valueNode.text().get();
        }
    }
    return "";
}

/**
 * @details
 */
void XLCustomProperties::deleteProperty(std::string_view name)
{
    if (m_xmlData == nullptr) return;
    for (auto p : xmlDocument().document_element().children("property")) {
        if (std::string_view(p.attribute("name").value()) == name) {
            xmlDocument().document_element().remove_child(p);
            return;
        }
    }
}

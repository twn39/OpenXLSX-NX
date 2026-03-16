// ===== External Includes ===== //
#include <algorithm>
#include <fmt/format.h>
#include <gsl/gsl>
#include <iterator>
#include <pugixml.hpp>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLSheet.hpp"
#include "XLUtilities.hpp"
#include "XLWorkbook.hpp"

using namespace OpenXLSX;

namespace
{
    /**
     * @brief Get the <sheets> node from the workbook XML
     * @param doc The workbook XML document
     * @return The <sheets> node
     */
    XMLNode sheetsNode(const XMLDocument& doc) { return doc.document_element().child("sheets"); }

    /**
     * @brief Node order for XLWorkbook XML. Strictly follows OOXML sequence with common extensions.
     * Reference: ECMA-376 Part 1, Section 18.2.27
     * Note: mc:AlternateContent often appears before workbookProtection in modern Excel files.
     */
    const std::vector<std::string_view> XLWorkbookNodeOrder = {
        "fileVersion",    "fileSharing",         "workbookPr",     "mc:AlternateContent",
        "xr:revisionPtr", "workbookProtection",  "bookViews",      "sheets",
        "functionGroups", "externalReferences",  "definedNames",   "calcPr",
        "oleSize",        "customWorkbookViews", "pivotCaches",    "smartTagPr",
        "smartTagTypes",  "webPublishing",       "fileRecoveryPr", "webPublishObjects",
        "extLst"};
}    // namespace


XLDefinedName::XLDefinedName(const XMLNode& node) : m_node(node) {}

std::string XLDefinedName::name() const { return m_node.attribute("name").value(); }
void        XLDefinedName::setName(std::string_view name) { m_node.attribute("name").set_value(std::string(name).c_str()); }

std::string XLDefinedName::refersTo() const { return m_node.text().get(); }
void        XLDefinedName::setRefersTo(std::string_view formula) { m_node.text().set(std::string(formula).c_str()); }

std::optional<uint32_t> XLDefinedName::localSheetId() const
{
    auto attr = m_node.attribute("localSheetId");
    if (attr.empty()) return std::nullopt;
    return attr.as_uint();
}

void XLDefinedName::setLocalSheetId(uint32_t id)
{
    auto attr = m_node.attribute("localSheetId");
    if (attr.empty()) attr = m_node.append_attribute("localSheetId");
    attr.set_value(id);
}

bool XLDefinedName::hidden() const { return m_node.attribute("hidden").as_bool(); }
void XLDefinedName::setHidden(bool hidden)
{
    auto attr = m_node.attribute("hidden");
    if (attr.empty()) attr = m_node.append_attribute("hidden");
    attr.set_value(hidden);
}

std::string XLDefinedName::comment() const { return m_node.attribute("comment").value(); }
void        XLDefinedName::setComment(std::string_view comment)
{
    auto attr = m_node.attribute("comment");
    if (attr.empty()) attr = m_node.append_attribute("comment");
    attr.set_value(std::string(comment).c_str());
}


XLDefinedNames::XLDefinedNames(const XMLNode& node) : m_node(node) {}

XLDefinedName XLDefinedNames::append(std::string_view name, std::string_view formula, std::optional<uint32_t> localSheetId)
{
    XMLNode newNode = m_node.append_child("definedName");
    newNode.append_attribute("name").set_value(std::string(name).c_str());
    if (localSheetId) newNode.append_attribute("localSheetId").set_value(*localSheetId);
    newNode.text().set(std::string(formula).c_str());
    return XLDefinedName(newNode);
}

void XLDefinedNames::remove(std::string_view name, std::optional<uint32_t> localSheetId)
{
    for (auto node : m_node.children("definedName")) {
        if (std::string_view(node.attribute("name").value()) == name) {
            auto attr = node.attribute("localSheetId");
            if (localSheetId) {
                if (!attr.empty() && attr.as_uint() == *localSheetId) {
                    m_node.remove_child(node);
                    return;
                }
            }
            else if (attr.empty()) {
                m_node.remove_child(node);
                return;
            }
        }
    }
}

XLDefinedName XLDefinedNames::get(std::string_view name, std::optional<uint32_t> localSheetId) const
{
    for (auto node : m_node.children("definedName")) {
        if (std::string_view(node.attribute("name").value()) == name) {
            auto attr = node.attribute("localSheetId");
            if (localSheetId) {
                if (!attr.empty() && attr.as_uint() == *localSheetId) return XLDefinedName(node);
            }
            else if (attr.empty()) {
                return XLDefinedName(node);
            }
        }
    }
    return XLDefinedName();
}

std::vector<XLDefinedName> XLDefinedNames::all() const
{
    std::vector<XLDefinedName> result;
    for (auto node : m_node.children("definedName")) { result.emplace_back(node); }
    return result;
}

bool XLDefinedNames::exists(std::string_view name, std::optional<uint32_t> localSheetId) const { return get(name, localSheetId).valid(); }

size_t XLDefinedNames::count() const
{
    size_t c = 0;
    for (auto node [[maybe_unused]] : m_node.children("definedName")) ++c;
    return c;
}


XLWorkbook::XLWorkbook(XLXmlData* xmlData) : XLXmlFile(xmlData) {}
XLWorkbook::~XLWorkbook() = default;

XLSheet XLWorkbook::sheet(std::string_view sheetName)
{
    auto snNode    = sheetsNode(xmlDocument());
    auto sheetNode = snNode.find_child_by_attribute("name", std::string(sheetName).c_str());
    if (sheetNode.empty()) throw XLInputError(fmt::format("Sheet \"{}\" does not exist", sheetName));

    const std::string xmlID = sheetNode.attribute("r:id").value();

    XLQuery pathQuery(XLQueryType::QuerySheetRelsTarget);
    pathQuery.setParam("sheetID", xmlID);
    auto xmlPath = parentDoc().execQuery(pathQuery).result<std::string>();

    if (xmlPath.substr(0, 4) == "/xl/") xmlPath = xmlPath.substr(4);

    XLQuery xmlQuery(XLQueryType::QueryXmlData);
    xmlQuery.setParam("xmlPath", "xl/" + xmlPath);
    return XLSheet(parentDoc().execQuery(xmlQuery).result<XLXmlData*>());
}

XLSheet XLWorkbook::sheet(uint16_t index)
{
    Expects(index >= 1);
    if (index > sheetCount()) throw XLInputError(fmt::format("Sheet index {} is out of bounds", index));

    uint16_t curIndex = 0;
    for (XMLNode node = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element); not node.empty();
         node         = node.next_sibling_of_type(pugi::node_element))
    {
        if (++curIndex == index) return sheet(node.attribute("name").value());
    }

    throw XLInternalError(fmt::format("Sheet index {} is out of bounds", index));
}

XLWorksheet  XLWorkbook::worksheet(std::string_view sheetName) { return sheet(sheetName).get<XLWorksheet>(); }
XLWorksheet  XLWorkbook::worksheet(uint16_t index) { return sheet(index).get<XLWorksheet>(); }
XLChartsheet XLWorkbook::chartsheet(std::string_view sheetName) { return sheet(sheetName).get<XLChartsheet>(); }
XLChartsheet XLWorkbook::chartsheet(uint16_t index) { return sheet(index).get<XLChartsheet>(); }

XLDefinedNames XLWorkbook::definedNames()
{
    auto    rootNode = xmlDocument().document_element();
    XMLNode dnNode   = rootNode.child("definedNames");
    if (dnNode.empty()) { dnNode = appendAndGetNode(rootNode, "definedNames", XLWorkbookNodeOrder); }
    return XLDefinedNames(dnNode);
}

void XLWorkbook::deleteNamedRanges() { xmlDocument().document_element().child("definedNames").remove_children(); }

void XLWorkbook::deleteSheet(std::string_view sheetName)
{
    auto sheetNode = sheetsNode(xmlDocument()).find_child_by_attribute("name", std::string(sheetName).c_str());
    if (sheetNode.empty()) { throw XLException(fmt::format("XLWorkbook::deleteSheet: workbook has no sheet with name \"{}\"", sheetName)); }

    std::string sheetID = sheetNode.attribute("r:id").value();

    XLQuery sheetTypeQuery(XLQueryType::QuerySheetType);
    sheetTypeQuery.setParam("sheetID", sheetID);
    const auto sheetType = parentDoc().execQuery(sheetTypeQuery).result<XLContentType>();

    const auto worksheetCount =
        std::count_if(sheetsNode(xmlDocument()).children().begin(), sheetsNode(xmlDocument()).children().end(), [&](const XMLNode& item) {
            if (item.type() != pugi::node_element) return false;
            XLQuery query(XLQueryType::QuerySheetType);
            query.setParam("sheetID", std::string(item.attribute("r:id").value()));
            return parentDoc().execQuery(query).result<XLContentType>() == XLContentType::Worksheet;
        });

    if (worksheetCount == 1 and sheetType == XLContentType::Worksheet)
        throw XLInputError("Invalid operation. There must be at least one worksheet in the workbook.");

    parentDoc().execCommand(
        XLCommand(XLCommandType::DeleteSheet).setParam("sheetID", sheetID).setParam("sheetName", std::string(sheetName)));

    // Delete non-element nodes after the sheet (whitespaces)
    XMLNode nonElementNode = sheetNode.next_sibling();
    while (not nonElementNode.empty() and nonElementNode.type() != pugi::node_element) {
        sheetsNode(xmlDocument()).remove_child(nonElementNode);
        nonElementNode = nonElementNode.next_sibling();
    }
    sheetsNode(xmlDocument()).remove_child(sheetNode);

    if (sheetIsActive(sheetID))
        xmlDocument().document_element().child("bookViews").first_child_of_type(pugi::node_element).remove_attribute("activeTab");
}

void XLWorkbook::addWorksheet(std::string_view sheetName)
{
    if (xmlDocument().document_element().child("sheets").find_child_by_attribute("name", std::string(sheetName).c_str()))
        throw XLInputError(fmt::format("Sheet named \"{}\" already exists.", sheetName));

    auto internalID = createInternalSheetID();

    parentDoc().execCommand(XLCommand(XLCommandType::AddWorksheet)
                                .setParam("sheetName", std::string(sheetName))
                                .setParam("sheetPath", fmt::format("/xl/worksheets/sheet{}.xml", internalID)));
    prepareSheetMetadata(sheetName, internalID);
}

void XLWorkbook::cloneSheet(std::string_view existingName, std::string_view newName)
{
    parentDoc().execCommand(
        XLCommand(XLCommandType::CloneSheet).setParam("sheetID", sheetID(existingName)).setParam("cloneName", std::string(newName)));
}

uint16_t XLWorkbook::createInternalSheetID()
{
    XMLNode  sheet           = xmlDocument().document_element().child("sheets").first_child_of_type(pugi::node_element);
    uint32_t maxSheetIdFound = 0;
    while (not sheet.empty()) {
        uint32_t thisSheetId = sheet.attribute("sheetId").as_uint();
        if (thisSheetId > maxSheetIdFound) maxSheetIdFound = thisSheetId;
        sheet = sheet.next_sibling_of_type(pugi::node_element);
    }
    return static_cast<uint16_t>(maxSheetIdFound + 1);
}

std::string XLWorkbook::sheetID(std::string_view sheetName)
{
    auto node = xmlDocument().document_element().child("sheets").find_child_by_attribute("name", std::string(sheetName).c_str());
    if (node.empty()) return "";
    return node.attribute("r:id").value();
}

std::string XLWorkbook::sheetName(std::string_view sheetID) const
{
    auto node = xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", std::string(sheetID).c_str());
    if (node.empty()) return "";
    return node.attribute("name").value();
}

std::string XLWorkbook::sheetVisibility(std::string_view sheetID) const
{
    auto node = xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", std::string(sheetID).c_str());
    if (node.empty()) return "";
    return node.attribute("state").value();
}

void XLWorkbook::prepareSheetMetadata(std::string_view sheetName, uint16_t internalID)
{
    auto        node                 = sheetsNode(xmlDocument()).append_child("sheet");
    std::string sheetPath            = fmt::format("/xl/worksheets/sheet{}.xml", internalID);
    node.append_attribute("name")    = std::string(sheetName).c_str();
    node.append_attribute("sheetId") = internalID;

    XLQuery query(XLQueryType::QuerySheetRelsID);
    query.setParam("sheetPath", sheetPath);
    node.append_attribute("r:id") = parentDoc().execQuery(query).result<std::string>().c_str();
}

void XLWorkbook::setSheetName(std::string_view sheetRID, std::string_view newName)
{
    auto sheetNode = xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", std::string(sheetRID).c_str());
    if (sheetNode.empty()) return;

    auto sheetNameAttr = sheetNode.attribute("name");
    updateSheetReferences(sheetNameAttr.value(), newName);
    sheetNameAttr.set_value(std::string(newName).c_str());
}

void XLWorkbook::setSheetVisibility(std::string_view sheetRID, std::string_view state)
{
    int     visibleSheets = 0;
    XMLNode item          = xmlDocument().document_element().child("sheets").first_child_of_type(pugi::node_element);
    while (not item.empty()) {
        if (std::string_view(item.attribute("r:id").value()) != sheetRID) {
            if (isVisible(item)) ++visibleSheets;
        }
        item = item.next_sibling_of_type(pugi::node_element);
    }

    bool hideSheet = not isVisibleState(state);
    if (hideSheet and visibleSheets == 0) throw XLSheetError("At least one sheet must be visible.");

    auto sheetNode      = xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", std::string(sheetRID).c_str());
    auto stateAttribute = sheetNode.attribute("state");
    if (stateAttribute.empty()) stateAttribute = sheetNode.prepend_attribute("state");
    stateAttribute.set_value(std::string(state).c_str());

    std::string name  = sheetNode.attribute("name").value();
    auto        index = indexOfSheet(name) - 1;

    XMLNode workbookView       = xmlDocument().document_element().child("bookViews").first_child_of_type(pugi::node_element);
    auto    activeTabAttribute = workbookView.attribute("activeTab");
    if (activeTabAttribute.empty()) {
        activeTabAttribute = workbookView.append_attribute("activeTab");
        activeTabAttribute.set_value(0);
    }
    const auto activeTabIndex = activeTabAttribute.as_uint();

    if (hideSheet and activeTabIndex == index) {
        XMLNode sheetItem = xmlDocument().document_element().child("sheets").first_child_of_type(pugi::node_element);
        while (not sheetItem.empty()) {
            if (isVisible(sheetItem)) {
                activeTabAttribute.set_value(indexOfSheet(sheetItem.attribute("name").value()) - 1);
                break;
            }
            sheetItem = sheetItem.next_sibling_of_type(pugi::node_element);
        }
    }
}

void XLWorkbook::setSheetIndex(std::string_view sheetName, unsigned int index)
{
    const XMLNode  workbookView     = xmlDocument().document_element().child("bookViews").first_child_of_type(pugi::node_element);
    const uint32_t activeSheetIndex = workbookView.attribute("activeTab").as_uint(0);

    XMLNode      sheetToMove{};
    unsigned int sheetToMoveIndex = 0;
    XMLNode      existingSheet{};
    std::string  activeSheet_rId{};

    unsigned int curIdx   = 0;
    XMLNode      curSheet = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);

    std::vector<std::string> originalSheetRIds;

    while (not curSheet.empty()) {
        originalSheetRIds.push_back(curSheet.attribute("r:id").value());
        if (sheetToMove.empty() and (std::string_view(curSheet.attribute("name").value()) == sheetName)) {
            sheetToMoveIndex = curIdx;
            sheetToMove      = curSheet;
        }
        if (curIdx == index - 1) { existingSheet = curSheet; }
        curSheet = curSheet.next_sibling_of_type(pugi::node_element);
        ++curIdx;
    }

    if (sheetToMove.empty()) throw XLInputError(fmt::format("XLWorkbook::setSheetIndex: no worksheet exists with name {}", sheetName));

    uint32_t targetIdx = std::min(index - 1, curIdx - 1);
    if (targetIdx == sheetToMoveIndex) return;

    activeSheet_rId = originalSheetRIds[activeSheetIndex];

    // Modify the node in the XML file
    if (index > curIdx)
        sheetsNode(xmlDocument()).append_move(sheetToMove);
    else {
        if (sheetToMoveIndex < targetIdx)
            sheetsNode(xmlDocument()).insert_move_after(sheetToMove, existingSheet);
        else
            sheetsNode(xmlDocument()).insert_move_before(sheetToMove, existingSheet);
    }

    // Update defined names with worksheet scopes
    std::vector<uint32_t>    oldToNew(originalSheetRIds.size());
    std::vector<std::string> newSheetRIds;
    curSheet = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
    while (!curSheet.empty()) {
        newSheetRIds.push_back(curSheet.attribute("r:id").value());
        curSheet = curSheet.next_sibling_of_type(pugi::node_element);
    }

    for (size_t oldI = 0; oldI < originalSheetRIds.size(); ++oldI) {
        for (size_t newI = 0; newI < newSheetRIds.size(); ++newI) {
            if (originalSheetRIds[oldI] == newSheetRIds[newI]) {
                oldToNew[oldI] = static_cast<uint32_t>(newI);
                break;
            }
        }
    }

    XMLNode definedName = xmlDocument().document_element().child("definedNames").first_child_of_type(pugi::node_element);
    while (not definedName.empty()) {
        auto attr = definedName.attribute("localSheetId");
        if (!attr.empty()) {
            uint32_t oldId = attr.as_uint();
            if (oldId < oldToNew.size()) { attr.set_value(oldToNew[oldId]); }
        }
        definedName = definedName.next_sibling_of_type(pugi::node_element);
    }

    if (!activeSheet_rId.empty()) setSheetActive(activeSheet_rId);
}

unsigned int XLWorkbook::indexOfSheet(std::string_view sheetName) const
{
    unsigned int index = 1;
    for (XMLNode sheet = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element); not sheet.empty();
         sheet         = sheet.next_sibling_of_type(pugi::node_element))
    {
        if (sheetName == sheet.attribute("name").value()) return index;
        index++;
    }

    throw XLInputError(fmt::format("Sheet \"{}\" does not exist", sheetName));
}

XLSheetType XLWorkbook::typeOfSheet(std::string_view sheetName) const
{
    if (!sheetExists(sheetName)) throw XLInputError(fmt::format("Sheet with name \"{}\" doesn't exist.", sheetName));

    if (worksheetExists(sheetName)) return XLSheetType::Worksheet;
    return XLSheetType::Chartsheet;
}

XLSheetType XLWorkbook::typeOfSheet(unsigned int index) const
{
    unsigned int thisIndex = 1;
    for (XMLNode sheet = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element); not sheet.empty();
         sheet         = sheet.next_sibling_of_type(pugi::node_element))
    {
        if (thisIndex == index) return typeOfSheet(sheet.attribute("name").value());
        ++thisIndex;
    }

    throw XLInputError(fmt::format("XLWorkbook::typeOfSheet: index {} is out of range", index));
}

unsigned int XLWorkbook::sheetCount() const
{
    unsigned int count = 0;
    for (XMLNode node = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element); not node.empty();
         node         = node.next_sibling_of_type(pugi::node_element))
        ++count;
    return count;
}

unsigned int XLWorkbook::worksheetCount() const
{
    unsigned int count = 0;
    for (XMLNode item = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element); not item.empty();
         item         = item.next_sibling_of_type(pugi::node_element))
    {
        XLQuery query(XLQueryType::QuerySheetType);
        query.setParam("sheetID", std::string(item.attribute("r:id").value()));
        if (parentDoc().execQuery(query).result<XLContentType>() == XLContentType::Worksheet) ++count;
    }
    return count;
}

unsigned int XLWorkbook::chartsheetCount() const
{
    unsigned int count = 0;
    for (XMLNode item = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element); not item.empty();
         item         = item.next_sibling_of_type(pugi::node_element))
    {
        XLQuery query(XLQueryType::QuerySheetType);
        query.setParam("sheetID", std::string(item.attribute("r:id").value()));
        if (parentDoc().execQuery(query).result<XLContentType>() == XLContentType::Chartsheet) ++count;
    }
    return count;
}

std::vector<std::string> XLWorkbook::sheetNames() const
{
    std::vector<std::string> results;
    for (XMLNode item = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element); not item.empty();
         item         = item.next_sibling_of_type(pugi::node_element))
        results.emplace_back(item.attribute("name").value());
    return results;
}

std::vector<std::string> XLWorkbook::worksheetNames() const
{
    std::vector<std::string> results;
    for (XMLNode item = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element); not item.empty();
         item         = item.next_sibling_of_type(pugi::node_element))
    {
        XLQuery query(XLQueryType::QuerySheetType);
        query.setParam("sheetID", std::string(item.attribute("r:id").value()));
        if (parentDoc().execQuery(query).result<XLContentType>() == XLContentType::Worksheet)
            results.emplace_back(item.attribute("name").value());
    }
    return results;
}

std::vector<std::string> XLWorkbook::chartsheetNames() const
{
    std::vector<std::string> results;
    for (XMLNode item = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element); not item.empty();
         item         = item.next_sibling_of_type(pugi::node_element))
    {
        XLQuery query(XLQueryType::QuerySheetType);
        query.setParam("sheetID", std::string(item.attribute("r:id").value()));
        if (parentDoc().execQuery(query).result<XLContentType>() == XLContentType::Chartsheet)
            results.emplace_back(item.attribute("name").value());
    }
    return results;
}

bool XLWorkbook::sheetExists(std::string_view sheetName) const
{ return !sheetsNode(xmlDocument()).find_child_by_attribute("name", std::string(sheetName).c_str()).empty(); }

bool XLWorkbook::worksheetExists(std::string_view sheetName) const
{
    auto names = worksheetNames();
    return std::find(names.begin(), names.end(), std::string(sheetName)) != names.end();
}

bool XLWorkbook::chartsheetExists(std::string_view sheetName) const
{
    auto names = chartsheetNames();
    return std::find(names.begin(), names.end(), std::string(sheetName)) != names.end();
}

void XLWorkbook::updateSheetReferences(std::string_view oldName, std::string_view newName)
{
    std::string oldNameTemp(oldName);
    std::string newNameTemp(newName);

    if (oldName.find(' ') != std::string_view::npos) oldNameTemp = fmt::format("'{}'", oldName);
    if (newName.find(' ') != std::string_view::npos) newNameTemp = fmt::format("'{}'", newName);

    oldNameTemp += '!';
    newNameTemp += '!';

    XMLNode definedName = xmlDocument().document_element().child("definedNames").first_child_of_type(pugi::node_element);
    for (; not definedName.empty(); definedName = definedName.next_sibling_of_type(pugi::node_element)) {
        std::string formula = definedName.text().get();
        if (formula.find('[') == std::string::npos and formula.find(']') == std::string::npos) {
            size_t pos = 0;
            while ((pos = formula.find(oldNameTemp, pos)) != std::string::npos) {
                formula.replace(pos, oldNameTemp.length(), newNameTemp);
                pos += newNameTemp.length();
            }
            definedName.text().set(formula.c_str());
        }
    }
}

void XLWorkbook::updateWorksheetDimensions()
{
    for (const auto& name : worksheetNames()) { worksheet(name).updateDimension(); }
}

void XLWorkbook::setFullCalculationOnLoad()
{
    auto    root   = xmlDocument().document_element();
    XMLNode calcPr = root.child("calcPr");
    if (calcPr.empty()) { calcPr = appendAndGetNode(root, "calcPr", XLWorkbookNodeOrder); }

    auto setAttr = [&](const char* name, bool val) {
        auto attr = calcPr.attribute(name);
        if (attr.empty()) attr = calcPr.append_attribute(name);
        attr.set_value(val);
    };

    setAttr("forceFullCalc", true);
    setAttr("fullCalcOnLoad", true);
}

void XLWorkbook::protect(bool lockStructure, bool lockWindows, std::string_view password)
{
    auto    root        = xmlDocument().document_element();
    XMLNode protectNode = root.child("workbookProtection");
    if (protectNode.empty()) { protectNode = appendAndGetNode(root, "workbookProtection", XLWorkbookNodeOrder); }

    auto setAttr = [&](const char* name, std::string_view val) {
        auto attr = protectNode.attribute(std::string(name).c_str());
        if (attr.empty()) attr = protectNode.append_attribute(std::string(name).c_str());
        attr.set_value(std::string(val).c_str());
    };

    setAttr("lockStructure", lockStructure ? "true" : "false");
    setAttr("lockWindows", lockWindows ? "true" : "false");

    if (!password.empty()) { setAttr("workbookPassword", ExcelPasswordHashAsString(std::string(password))); }
    else {
        protectNode.remove_attribute("workbookPassword");
    }
}

void XLWorkbook::unprotect() { xmlDocument().document_element().remove_child("workbookProtection"); }

bool XLWorkbook::isProtected() const { return !xmlDocument().document_element().child("workbookProtection").empty(); }

void XLWorkbook::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print(ostr); }

bool XLWorkbook::sheetIsActive(std::string_view sheetRID) const
{
    const XMLNode      workbookView       = xmlDocument().document_element().child("bookViews").first_child_of_type(pugi::node_element);
    const XMLAttribute activeTabAttribute = workbookView.attribute("activeTab");
    if (activeTabAttribute.empty()) return false;

    int32_t activeTabIndex = activeTabAttribute.as_int();
    int32_t index          = 0;
    XMLNode item           = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
    while (not item.empty()) {
        if (std::string_view(item.attribute("r:id").value()) == sheetRID) return index == activeTabIndex;
        ++index;
        item = item.next_sibling_of_type(pugi::node_element);
    }

    return false;
}

bool XLWorkbook::setSheetActive(std::string_view sheetRID)
{
    XMLNode            workbookView       = xmlDocument().document_element().child("bookViews").first_child_of_type(pugi::node_element);
    const XMLAttribute activeTabAttribute = workbookView.attribute("activeTab");
    int32_t            activeTabIndex     = activeTabAttribute.empty() ? -1 : activeTabAttribute.as_int();

    int32_t index = 0;
    XMLNode item  = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
    while (not item.empty() and (std::string_view(item.attribute("r:id").value()) != sheetRID)) {
        ++index;
        item = item.next_sibling_of_type(pugi::node_element);
    }

    if (item.empty() or !isVisible(item)) return false;

    if ((activeTabIndex != -1) and (index != activeTabIndex)) sheet(static_cast<uint16_t>(activeTabIndex + 1)).setSelected(false);

    auto attr = workbookView.attribute("activeTab");
    if (attr.empty()) attr = workbookView.append_attribute("activeTab");
    attr.set_value(index);

    return true;
}

bool XLWorkbook::isVisibleState(std::string_view state) const { return (state != "hidden" and state != "veryHidden"); }

bool XLWorkbook::isVisible(XMLNode const& sheetNode) const
{
    if (sheetNode.empty()) return false;
    auto attr = sheetNode.attribute("state");
    if (attr.empty()) return true;
    return isVisibleState(attr.value());
}

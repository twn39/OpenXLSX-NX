// ===== External Includes ===== //
#include <cctype>
#include <charconv>
#include <cstdint>
#include <fmt/format.h>
#include <pugixml.hpp>
#include <string>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLDrawing.hpp"
#include "XLRelationships.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{
    constexpr std::string_view ShapeNodeName     = "v:shape";
    constexpr std::string_view ShapeTypeNodeName = "v:shapetype";

    constexpr uint32_t EMU_PER_PIXEL      = 9525;
    constexpr uint32_t DEFAULT_IMAGE_SIZE = 100;

    // utility functions

    bool wouldBeDuplicateShapeType(XMLNode const& rootNode, XMLNode const& shapeTypeNode)
    {
        const std::string id   = shapeTypeNode.attribute("id").value();
        XMLNode           node = rootNode.first_child_of_type(pugi::node_element);
        using namespace std::literals::string_literals;
        while (node != shapeTypeNode && node.raw_name() == std::string(ShapeTypeNodeName)) {
            if (node.attribute("id").value() == id) return true;
            node = node.next_sibling_of_type(pugi::node_element);
        }
        return false;
    }

    XMLNode moveNode(XMLNode& rootNode, XMLNode& node, XMLNode const& insertAfter, bool withWhitespaces = true)
    {
        if (rootNode.empty() or node.empty()) return XMLNode{};
        if (node == insertAfter or node == insertAfter.next_sibling()) return node;

        XMLNode newNode{};
        if (insertAfter.empty())
            newNode = rootNode.prepend_copy(node);
        else
            newNode = rootNode.insert_copy_after(node, insertAfter);

        if (withWhitespaces) {
            copyLeadingWhitespaces(rootNode, node, newNode);
            while (node.previous_sibling().type() == pugi::node_pcdata) rootNode.remove_child(node.previous_sibling());
        }
        rootNode.remove_child(node);

        return newNode;
    }

    XLShapeTextVAlign XLShapeTextVAlignFromString(std::string_view vAlignString)
    {
        if (vAlignString == "Center") return XLShapeTextVAlign::Center;
        if (vAlignString == "Top") return XLShapeTextVAlign::Top;
        if (vAlignString == "Bottom") return XLShapeTextVAlign::Bottom;
        throw XLInternalError(std::string(__func__) + ": invalid XLShapeTextVAlign setting " + std::string(vAlignString));
    }
    std::string XLShapeTextVAlignToString(XLShapeTextVAlign vAlign)
    {
        switch (vAlign) {
            case XLShapeTextVAlign::Center:
                return "Center";
            case XLShapeTextVAlign::Top:
                return "Top";
            case XLShapeTextVAlign::Bottom:
                return "Bottom";
            case XLShapeTextVAlign::Invalid:
                [[fallthrough]];
            default:
                return "(invalid)";
        }
    }
    XLShapeTextHAlign XLShapeTextHAlignFromString(std::string_view hAlignString)
    {
        if (hAlignString == "Left") return XLShapeTextHAlign::Left;
        if (hAlignString == "Right") return XLShapeTextHAlign::Right;
        if (hAlignString == "Center") return XLShapeTextHAlign::Center;
        throw XLInternalError(std::string(__func__) + ": invalid XLShapeTextHAlign setting " + std::string(hAlignString));
    }
    std::string XLShapeTextHAlignToString(XLShapeTextHAlign hAlign)
    {
        switch (hAlign) {
            case XLShapeTextHAlign::Left:
                return "Left";
            case XLShapeTextHAlign::Right:
                return "Right";
            case XLShapeTextHAlign::Center:
                return "Center";
            case XLShapeTextHAlign::Invalid:
                [[fallthrough]];
            default:
                return "(invalid)";
        }
    }

}    // namespace OpenXLSX

XLShapeClientData::XLShapeClientData() : m_clientDataNode() {}

XLShapeClientData::XLShapeClientData(const XMLNode& node) : m_clientDataNode(node) {}

bool elementTextAsBool(XMLNode const& node)
{
    if (node.empty()) return false;
    if (node.text().empty()) return true;
    char c = node.text().get()[0];
    if (c == 't' or c == 'T') return true;
    return false;
}

std::string       XLShapeClientData::objectType() const { return appendAndGetAttribute(m_clientDataNode, "ObjectType", "Note").value(); }
bool              XLShapeClientData::moveWithCells() const { return elementTextAsBool(m_clientDataNode.child("x:MoveWithCells")); }
bool              XLShapeClientData::sizeWithCells() const { return elementTextAsBool(m_clientDataNode.child("x:SizeWithCells")); }
std::string       XLShapeClientData::anchor() const { return m_clientDataNode.child("x:Anchor").value(); }
bool              XLShapeClientData::autoFill() const { return elementTextAsBool(m_clientDataNode.child("x:AutoFill")); }
XLShapeTextVAlign XLShapeClientData::textVAlign() const
{ return XLShapeTextVAlignFromString(m_clientDataNode.child("x:TextVAlign").text().get()); }
XLShapeTextHAlign XLShapeClientData::textHAlign() const
{ return XLShapeTextHAlignFromString(m_clientDataNode.child("x:TextHAlign").text().get()); }
uint32_t XLShapeClientData::row() const
{ return appendAndGetNode(m_clientDataNode, "x:Row", m_nodeOrder, XLForceNamespace).text().as_uint(0); }
uint16_t XLShapeClientData::column() const
{ return static_cast<uint16_t>(appendAndGetNode(m_clientDataNode, "x:Column", m_nodeOrder, XLForceNamespace).text().as_uint(0)); }

bool XLShapeClientData::setObjectType(std::string_view newObjectType)
{ return appendAndSetAttribute(m_clientDataNode, "ObjectType", std::string(newObjectType)).empty() == false; }

bool XLShapeClientData::setMoveWithCells(bool set)
{ return appendAndGetNode(m_clientDataNode, "x:MoveWithCells", m_nodeOrder, XLForceNamespace).text().set(set ? "" : "False"); }

bool XLShapeClientData::setSizeWithCells(bool set)
{ return appendAndGetNode(m_clientDataNode, "x:SizeWithCells", m_nodeOrder, XLForceNamespace).text().set(set ? "" : "False"); }

bool XLShapeClientData::setAnchor(std::string_view newAnchor)
{ return appendAndGetNode(m_clientDataNode, "x:Anchor", m_nodeOrder, XLForceNamespace).text().set(std::string(newAnchor).c_str()); }

bool XLShapeClientData::setAutoFill(bool set)
{ return appendAndGetNode(m_clientDataNode, "x:AutoFill", m_nodeOrder, XLForceNamespace).text().set(set ? "True" : "False"); }

bool XLShapeClientData::setTextVAlign(XLShapeTextVAlign newTextVAlign)
{
    return appendAndGetNode(m_clientDataNode, "x:TextVAlign", m_nodeOrder, XLForceNamespace)
        .text()
        .set(XLShapeTextVAlignToString(newTextVAlign).c_str());
}

bool XLShapeClientData::setTextHAlign(XLShapeTextHAlign newTextHAlign)
{
    return appendAndGetNode(m_clientDataNode, "x:TextHAlign", m_nodeOrder, XLForceNamespace)
        .text()
        .set(XLShapeTextHAlignToString(newTextHAlign).c_str());
}

bool XLShapeClientData::setRow(uint32_t newRow)
{ return appendAndGetNode(m_clientDataNode, "x:Row", m_nodeOrder, XLForceNamespace).text().set(newRow); }

bool XLShapeClientData::setColumn(uint16_t newColumn)
{ return appendAndGetNode(m_clientDataNode, "x:Column", m_nodeOrder, XLForceNamespace).text().set(newColumn); }

XLShapeStyle::XLShapeStyle() : m_style(""), m_styleAttribute(XMLAttribute())
{
    setPosition("absolute");
    setMarginLeft(100);
    setMarginTop(0);
    setWidth(200);     // Increased from 120
    setHeight(150);    // Increased from 80
    hide();
}

XLShapeStyle::XLShapeStyle(const std::string& styleAttribute) : m_style(styleAttribute), m_styleAttribute(XMLAttribute()) {}

XLShapeStyle::XLShapeStyle(const XMLAttribute& styleAttribute) : m_style(""), m_styleAttribute(styleAttribute) {}

int16_t XLShapeStyle::attributeOrderIndex(std::string_view attributeName) const
{
    auto attributeIterator = std::find(m_nodeOrder.begin(), m_nodeOrder.end(), attributeName);
    if (attributeIterator == m_nodeOrder.end()) return -1;
    return static_cast<int16_t>(attributeIterator - m_nodeOrder.begin());
}

XLShapeStyleAttribute XLShapeStyle::getAttribute(std::string_view attributeName, std::string_view valIfNotFound) const
{
    if (attributeOrderIndex(attributeName) == -1) {
        using namespace std::literals::string_literals;
        throw XLInternalError("XLShapeStyle.getAttribute: attribute "s + std::string(attributeName) + " is not defined in class"s);
    }

    // if attribute is linked, re-read m_style each time in case the underlying XML has changed
    if (not m_styleAttribute.empty()) m_style = std::string(m_styleAttribute.value());

    size_t                lastPos = 0;
    XLShapeStyleAttribute result;
    result.name  = "";    // indicates "not found"
    result.value = "";    // default in case attribute name is found but has no value

    do {
        size_t      pos      = m_style.find(';', lastPos);
        std::string attrPair = m_style.substr(lastPos, pos - lastPos);
        if (!attrPair.empty()) {
            size_t separatorPos = attrPair.find(':');
            if (attributeName == attrPair.substr(0, separatorPos)) {    // found!
                result.name = std::string(attributeName);
                if (separatorPos != std::string::npos) result.value = attrPair.substr(separatorPos + 1);
                break;
            }
        }
        lastPos = pos + 1;
    }
    while (lastPos < m_style.length());
    if (lastPos >= m_style.length()) result.value = std::string(valIfNotFound);
    return result;
}

bool XLShapeStyle::setAttribute(std::string_view attributeName, std::string_view attributeValue)
{
    int16_t attrIndex = attributeOrderIndex(attributeName);
    if (attrIndex == -1) {
        using namespace std::literals::string_literals;
        throw XLInternalError("XLShapeStyle.setAttribute: attribute "s + std::string(attributeName) + " is not defined in class"s);
    }

    if (not m_styleAttribute.empty()) m_style = std::string(m_styleAttribute.value());

    size_t lastPos   = 0;
    size_t appendPos = 0;
    do {
        size_t pos = m_style.find(';', lastPos);

        std::string attrPair = m_style.substr(lastPos, pos - lastPos);
        if (!attrPair.empty()) {
            size_t  separatorPos  = attrPair.find(':');
            int16_t thisAttrIndex = attributeOrderIndex(attrPair.substr(0, separatorPos));
            if (thisAttrIndex >= attrIndex) {
                appendPos = (thisAttrIndex == attrIndex ? pos : lastPos);
                break;
            }
        }
        if (pos == std::string::npos)
            lastPos = pos;
        else
            lastPos = pos + 1;
    }
    while (lastPos < m_style.length());
    if (lastPos >= m_style.length() or appendPos > m_style.length()) {
        appendPos = m_style.length();

        if (lastPos > m_style.length()) lastPos = m_style.length();
    }

    using namespace std::literals::string_literals;
    m_style = m_style.substr(0, lastPos) + ((lastPos != 0 and appendPos == m_style.length()) ? ";"s : ""s) + std::string(attributeName) +
              ":"s + std::string(attributeValue) + (appendPos < m_style.length() ? ";"s : ""s) + m_style.substr(appendPos);

    if (not m_styleAttribute.empty()) m_styleAttribute.set_value(m_style.c_str());

    return true;
}

std::string XLShapeStyle::position() const { return getAttribute("position").value; }
uint16_t    XLShapeStyle::marginLeft() const
{
    uint16_t    val = 0;
    std::string str = getAttribute("margin-left").value;
    std::from_chars(str.data(), str.data() + str.size(), val);
    return val;
}
uint16_t XLShapeStyle::marginTop() const
{
    uint16_t    val = 0;
    std::string str = getAttribute("margin-top").value;
    std::from_chars(str.data(), str.data() + str.size(), val);
    return val;
}
uint16_t XLShapeStyle::width() const
{
    uint16_t    val = 0;
    std::string str = getAttribute("width").value;
    std::from_chars(str.data(), str.data() + str.size(), val);
    return val;
}
uint16_t XLShapeStyle::height() const
{
    uint16_t    val = 0;
    std::string str = getAttribute("height").value;
    std::from_chars(str.data(), str.data() + str.size(), val);
    return val;
}
std::string XLShapeStyle::msoWrapStyle() const { return getAttribute("mso-wrap-style").value; }
std::string XLShapeStyle::vTextAnchor() const { return getAttribute("v-text-anchor").value; }
bool        XLShapeStyle::hidden() const { return ("hidden" == getAttribute("visibility").value); }
bool        XLShapeStyle::visible() const { return !hidden(); }

bool XLShapeStyle::setPosition(std::string_view newPosition) { return setAttribute("position", newPosition); }
bool XLShapeStyle::setMarginLeft(uint16_t newMarginLeft) { return setAttribute("margin-left", fmt::format("{}pt", newMarginLeft)); }
bool XLShapeStyle::setMarginTop(uint16_t newMarginTop) { return setAttribute("margin-top", fmt::format("{}pt", newMarginTop)); }
bool XLShapeStyle::setWidth(uint16_t newWidth) { return setAttribute("width", fmt::format("{}pt", newWidth)); }
bool XLShapeStyle::setHeight(uint16_t newHeight) { return setAttribute("height", fmt::format("{}pt", newHeight)); }
bool XLShapeStyle::setMsoWrapStyle(std::string_view newMsoWrapStyle) { return setAttribute("mso-wrap-style", newMsoWrapStyle); }
bool XLShapeStyle::setVTextAnchor(std::string_view newVTextAnchor) { return setAttribute("v-text-anchor", newVTextAnchor); }
bool XLShapeStyle::hide() { return setAttribute("visibility", "hidden"); }
bool XLShapeStyle::show() { return setAttribute("visibility", "visible"); }

XLShape::XLShape() : m_shapeNode() {}

XLShape::XLShape(const XMLNode& node) : m_shapeNode(node) {}

std::string  XLShape::shapeId() const { return m_shapeNode.attribute("id").value(); }
std::string  XLShape::fillColor() const { return appendAndGetAttribute(m_shapeNode, "fillcolor", "#ffffc0").value(); }
bool         XLShape::stroked() const { return appendAndGetAttribute(m_shapeNode, "stroked", "t").as_bool(); }
std::string  XLShape::type() const { return appendAndGetAttribute(m_shapeNode, "type", "").value(); }
bool         XLShape::allowInCell() const { return appendAndGetAttribute(m_shapeNode, "o:allowincell", "f").as_bool(); }
XLShapeStyle XLShape::style() { return XLShapeStyle(appendAndGetAttribute(m_shapeNode, "style", "")); }

XLShapeClientData XLShape::clientData()
{ return XLShapeClientData(appendAndGetNode(m_shapeNode, "x:ClientData", m_nodeOrder, XLForceNamespace)); }

bool XLShape::setFillColor(std::string_view newFillColor)
{ return appendAndSetAttribute(m_shapeNode, "fillcolor", std::string(newFillColor)).empty() == false; }
bool XLShape::setStroked(bool set) { return appendAndSetAttribute(m_shapeNode, "stroked", (set ? "t" : "f")).empty() == false; }
bool XLShape::setType(std::string_view newType)
{ return appendAndSetAttribute(m_shapeNode, "type", std::string(newType)).empty() == false; }
bool XLShape::setAllowInCell(bool set) { return appendAndSetAttribute(m_shapeNode, "o:allowincell", (set ? "t" : "f")).empty() == false; }
bool XLShape::setStyle(std::string_view newStyle)
{ return appendAndSetAttribute(m_shapeNode, "style", std::string(newStyle)).empty() == false; }
bool XLShape::setStyle(XLShapeStyle const& newStyle) { return setStyle(newStyle.raw()); }

XLVmlDrawing::XLVmlDrawing(gsl::not_null<XLXmlData*> xmlData) : XLXmlFile(xmlData)
{
    if (xmlData->getXmlType() != XLContentType::VMLDrawing) throw XLInternalError("XLVmlDrawing constructor: Invalid XML data.");

    XMLDocument& doc = xmlDocument();
    if (doc.document_element().empty())    // handle a bad (no document element) drawing XML file
        doc.load_string("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
                        "<xml"
                        " xmlns:v=\"urn:schemas-microsoft-com:vml\""
                        " xmlns:o=\"urn:schemas-microsoft-com:office:office\""
                        " xmlns:x=\"urn:schemas-microsoft-com:office:excel\""
                        " xmlns:w10=\"urn:schemas-microsoft-com:office:word\""
                        ">"
                        "\n</xml>",
                        pugi_parse_settings);

    // ===== Re-sort the document: move all v:shapetype nodes to the beginning of the XML document element and eliminate duplicates
    // ===== Also: determine highest used shape id, regardless of basename (pattern [^0-9]*[0-9]*) and m_shapeCount
    using namespace std::literals::string_literals;

    XMLNode rootNode = doc.document_element();
    XMLNode node     = rootNode.first_child_of_type(pugi::node_element);
    XMLNode lastShapeTypeNode{};
    while (not node.empty()) {
        XMLNode nextNode =
            node.next_sibling_of_type(pugi::node_element);    // determine next node early because node may be invalidated by moveNode
        if (node.raw_name() == ShapeTypeNodeName) {
            if (wouldBeDuplicateShapeType(rootNode, node)) {                   // if shapetype attribute id already exists at begin of file
                while (node.previous_sibling().type() == pugi::node_pcdata)    // delete preceeding whitespaces
                    rootNode.remove_child(node.previous_sibling());            //  ...
                rootNode.remove_child(node);    // and the v:shapetype node, as it can not be referenced for lack of a unique id
            }
            else
                lastShapeTypeNode =
                    moveNode(rootNode, node, lastShapeTypeNode);    // move node to end of continuous list of shapetype nodes
        }
        else if (node.raw_name() == ShapeNodeName) {
            ++m_shapeCount;
            std::string strNodeId = node.attribute("id").value();
            size_t      pos       = strNodeId.length();
            uint32_t    nodeId    = 0;
            while (pos > 0 and std::isdigit(strNodeId[--pos]))    // determine any trailing nodeId
                nodeId = 10 * nodeId + static_cast<uint32_t>(strNodeId[pos] - '0');
            m_lastAssignedShapeId = std::max(m_lastAssignedShapeId, nodeId);
        }
        node = nextNode;
    }
    // Henceforth: assume that it is safe to consider shape nodes a continuous list (well - unless there are other node types as well)

    XMLNode shapeTypeNode{};
    if (not lastShapeTypeNode.empty()) {
        shapeTypeNode = rootNode.first_child_of_type(pugi::node_element);
        while (not shapeTypeNode.empty() and shapeTypeNode.raw_name() != ShapeTypeNodeName)
            shapeTypeNode = shapeTypeNode.next_sibling_of_type(pugi::node_element);
    }
    if (shapeTypeNode.empty()) {
        shapeTypeNode = rootNode.prepend_child(std::string(ShapeTypeNodeName).c_str(), XLForceNamespace);
        rootNode.prepend_child(pugi::node_pcdata).set_value("\n\t");
    }
    if (shapeTypeNode.first_child().empty())
        shapeTypeNode.append_child(pugi::node_pcdata).set_value("\n\t");    // insert indentation if node was empty

    // ===== Ensure that attributes exist for first shapetype node, default-initialize if needed:
    m_defaultShapeTypeId = appendAndGetAttribute(shapeTypeNode, "id", "_x0000_t202").value();
    appendAndGetAttribute(shapeTypeNode, "coordsize", "21600,21600");
    appendAndGetAttribute(shapeTypeNode, "o:spt", "202");
    appendAndGetAttribute(shapeTypeNode, "path", "m,l,21600l21600,21600l21600,xe");

    XMLNode strokeNode = shapeTypeNode.child("v:stroke");
    if (strokeNode.empty()) {
        strokeNode = shapeTypeNode.prepend_child("v:stroke", XLForceNamespace);
        shapeTypeNode.prepend_child(pugi::node_pcdata).set_value("\n\t\t");
    }
    appendAndGetAttribute(strokeNode, "joinstyle", "miter");

    XMLNode pathNode = shapeTypeNode.child("v:path");
    if (pathNode.empty()) {
        pathNode = shapeTypeNode.insert_child_after("v:path", strokeNode, XLForceNamespace);
        copyLeadingWhitespaces(shapeTypeNode, strokeNode, pathNode);
    }
    appendAndGetAttribute(pathNode, "gradientshapeok", "t");
    appendAndGetAttribute(pathNode, "o:connecttype", "rect");
}

XMLNode XLVmlDrawing::firstShapeNode() const
{
    using namespace std::literals::string_literals;

    XMLNode node = xmlDocument().document_element().first_child_of_type(pugi::node_element);
    while (not node.empty() and node.raw_name() != ShapeNodeName) node = node.next_sibling_of_type(pugi::node_element);
    return node;
}

XMLNode XLVmlDrawing::lastShapeNode() const
{
    using namespace std::literals::string_literals;

    XMLNode node = xmlDocument().document_element().last_child_of_type(pugi::node_element);
    while (not node.empty() and node.raw_name() != ShapeNodeName) node = node.previous_sibling_of_type(pugi::node_element);
    return node;
}

XMLNode XLVmlDrawing::shapeNode(uint32_t index) const
{
    using namespace std::literals::string_literals;

    XMLNode node{};
    if (index < m_shapeCount) {
        uint16_t i = 0;
        node       = firstShapeNode();
        while (i != index and not node.empty() and node.raw_name() == ShapeNodeName) {
            ++i;
            node = node.next_sibling_of_type(pugi::node_element);
        }
    }
    if (node.empty() or node.raw_name() != ShapeNodeName)
        throw XLException(fmt::format("XLVmlDrawing: shape index {} is out of bounds", index));

    return node;
}

XMLNode XLVmlDrawing::shapeNode(std::string_view cellRef) const
{
    const XLCellReference destRef{std::string(cellRef)};
    const uint32_t        destRow = destRef.row() - 1;
    const uint16_t        destCol = destRef.column() - 1;

    XMLNode node = firstShapeNode();
    while (not node.empty()) {
        if ((destRow == node.child("x:ClientData").child("x:Row").text().as_uint()) &&
            (destCol == node.child("x:ClientData").child("x:Column").text().as_uint()))
            break;

        do {
            node = node.next_sibling_of_type(pugi::node_element);
        }
        while (not node.empty() and node.name() != ShapeNodeName);
    }
    return node;
}

uint32_t XLVmlDrawing::shapeCount() const { return m_shapeCount; }

XLShape XLVmlDrawing::shape(uint32_t index) const { return XLShape(shapeNode(index)); }

bool XLVmlDrawing::deleteShape(uint32_t index)
{
    XMLNode rootNode = xmlDocument().document_element();
    XMLNode node     = shapeNode(index);
    --m_shapeCount;
    while (node.previous_sibling().type() == pugi::node_pcdata) rootNode.remove_child(node.previous_sibling());
    rootNode.remove_child(node);

    return true;
}

bool XLVmlDrawing::deleteShape(std::string_view cellRef)
{
    XMLNode rootNode = xmlDocument().document_element();
    XMLNode node     = shapeNode(cellRef);
    if (node.empty()) return false;

    --m_shapeCount;
    while (node.previous_sibling().type() == pugi::node_pcdata) rootNode.remove_child(node.previous_sibling());
    rootNode.remove_child(node);

    return true;
}

XLShape XLVmlDrawing::createShape([[maybe_unused]] const XLShape& shapeTemplate)
{
    XMLNode rootNode = xmlDocument().document_element();
    XMLNode node     = lastShapeNode();
    if (node.empty()) { node = rootNode.last_child_of_type(pugi::node_element); }
    if (not node.empty()) {
        node = rootNode.insert_child_after(std::string(ShapeNodeName).c_str(), node, XLForceNamespace);
        copyLeadingWhitespaces(rootNode, node.previous_sibling(), node);
    }
    else {
        node = rootNode.prepend_child(std::string(ShapeNodeName).c_str(), XLForceNamespace);
        rootNode.prepend_child(pugi::node_pcdata).set_value("\n\t");
    }

    using namespace std::literals::string_literals;
    node.prepend_attribute("id").set_value(fmt::format("shape_{}", m_lastAssignedShapeId++).c_str());
    node.append_attribute("type").set_value(("#"s + m_defaultShapeTypeId).c_str());

    m_shapeCount++;
    return XLShape(node);
}

void XLVmlDrawing::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print(ostr); }

XLDrawing::XLDrawing(gsl::not_null<XLXmlData*> xmlData) : XLXmlFile(xmlData), m_relationships(nullptr)
{
    if (xmlData->getXmlType() != XLContentType::Drawing) throw XLInternalError("XLDrawing constructor: Invalid XML data.");

    initXml();
}

XLDrawing::XLDrawing(const XLDrawing& other) : XLXmlFile(other), m_relationships(nullptr)
{
    if (other.m_relationships) { m_relationships = std::make_unique<XLRelationships>(*other.m_relationships); }
}

XLDrawing::XLDrawing(XLDrawing&& other) noexcept = default;

XLDrawing& XLDrawing::operator=(const XLDrawing& other)
{
    if (this != &other) {
        XLXmlFile::operator=(other);
        if (other.m_relationships) { m_relationships = std::make_unique<XLRelationships>(*other.m_relationships); }
        else {
            m_relationships.reset();
        }
    }
    return *this;
}

XLDrawing& XLDrawing::operator=(XLDrawing&& other) noexcept = default;

void XLDrawing::initXml()
{
    if (getXmlPath().empty()) return;
    XMLDocument& doc = xmlDocument();
    if (doc.document_element().empty()) {
        doc.load_string("<?xml version=\"1.0\" encoding=\"UTF-8\"standalone=\"yes\"?>\n"
                        "<xdr:wsDr"
                        " xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\""
                        " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
                        " xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\""
                        " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                        ">"
                        "</xdr:wsDr>",
                        pugi_parse_settings);
    }
}

void XLDrawing::addImage(std::string_view      rId,
                         std::string_view      name,
                         std::string_view      description,
                         uint32_t              row,
                         uint32_t              col,
                         uint32_t              width,
                         uint32_t              height,
                         const XLImageOptions& options)
{
    XMLNode rootNode = xmlDocument().document_element();

    XMLNode anchor;
    if (options.positioning == XLImagePositioning::TwoCell) { anchor = rootNode.append_child("xdr:twoCellAnchor"); }
    else if (options.positioning == XLImagePositioning::Absolute) {
        anchor = rootNode.append_child("xdr:absoluteAnchor");
    }
    else {
        anchor = rootNode.append_child("xdr:oneCellAnchor");
    }

    if (options.positioning == XLImagePositioning::Absolute) {
        XMLNode pos = anchor.append_child("xdr:pos");
        pos.append_attribute("x").set_value(static_cast<int64_t>(options.offsetX) * EMU_PER_PIXEL);
        pos.append_attribute("y").set_value(static_cast<int64_t>(options.offsetY) * EMU_PER_PIXEL);
    }
    else {
        // From node (top-left)
        XMLNode from = anchor.append_child("xdr:from");
        from.append_child("xdr:col").text().set(col);
        from.append_child("xdr:colOff").text().set(static_cast<int64_t>(options.offsetX) * EMU_PER_PIXEL);
        from.append_child("xdr:row").text().set(row);
        from.append_child("xdr:rowOff").text().set(static_cast<int64_t>(options.offsetY) * EMU_PER_PIXEL);

        if (options.positioning == XLImagePositioning::TwoCell) {
            uint32_t toRow = row + 1;    // Default fallback if not provided
            uint32_t toCol = col + 1;
            if (!options.bottomRightCell.empty()) {
                XLCellReference brRef(options.bottomRightCell);
                toRow = brRef.row() - 1;
                toCol = brRef.column() - 1;
            }

            XMLNode to = anchor.append_child("xdr:to");
            to.append_child("xdr:col").text().set(toCol);
            to.append_child("xdr:colOff").text().set(0);
            to.append_child("xdr:row").text().set(toRow);
            to.append_child("xdr:rowOff").text().set(0);
        }
    }

    const uint64_t emuWidth  = static_cast<uint64_t>(width) * EMU_PER_PIXEL;
    const uint64_t emuHeight = static_cast<uint64_t>(height) * EMU_PER_PIXEL;

    // Extents (size) - Required for oneCellAnchor and absoluteAnchor
    if (options.positioning != XLImagePositioning::TwoCell) {
        XMLNode ext = anchor.append_child("xdr:ext");
        ext.append_attribute("cx").set_value(fmt::format("{}", emuWidth).c_str());
        ext.append_attribute("cy").set_value(fmt::format("{}", emuHeight).c_str());
    }

    // Picture node
    XMLNode pic = anchor.append_child("xdr:pic");

    // Non-visual properties
    XMLNode nvPicPr    = pic.append_child("xdr:nvPicPr");
    XMLNode cNvPr      = nvPicPr.append_child("xdr:cNvPr");
    auto    childCount = static_cast<size_t>(std::distance(rootNode.children().begin(), rootNode.children().end()));
    cNvPr.append_attribute("id").set_value(fmt::format("{}", childCount + 1).c_str());
    cNvPr.append_attribute("name").set_value(std::string(name).c_str());
    cNvPr.append_attribute("descr").set_value(std::string(description).c_str());

    nvPicPr.append_child("xdr:cNvPicPr").append_child("a:picLocks").append_attribute("noChangeAspect").set_value("1");

    // Blip fill
    XMLNode blipFill = pic.append_child("xdr:blipFill");
    XMLNode blip     = blipFill.append_child("a:blip");
    blip.append_attribute("r:embed").set_value(std::string(rId).c_str());

    blipFill.append_child("a:stretch").append_child("a:fillRect");

    // Shape properties
    XMLNode spPr = pic.append_child("xdr:spPr");
    XMLNode xfrm = spPr.append_child("a:xfrm");
    xfrm.append_child("a:off").append_attribute("x").set_value("0");
    xfrm.child("a:off").append_attribute("y").set_value("0");
    xfrm.append_child("a:ext").append_attribute("cx").set_value(fmt::format("{}", emuWidth).c_str());
    xfrm.child("a:ext").append_attribute("cy").set_value(fmt::format("{}", emuHeight).c_str());

    XMLNode prstGeom = spPr.append_child("a:prstGeom");
    prstGeom.append_attribute("prst").set_value("rect");
    prstGeom.append_child("a:avLst");

    anchor.append_child("xdr:clientData");
}

void XLDrawing::addChartAnchor(std::string_view rId, std::string_view name, uint32_t row, uint32_t col, uint32_t width, uint32_t height)
{
    XMLNode rootNode = xmlDocument().document_element();

    // Create oneCellAnchor
    XMLNode anchor = rootNode.append_child("xdr:oneCellAnchor");

    // From node (top-left)
    XMLNode from = anchor.append_child("xdr:from");
    from.append_child("xdr:col").text().set(col);
    from.append_child("xdr:colOff").text().set(0);
    from.append_child("xdr:row").text().set(row);
    from.append_child("xdr:rowOff").text().set(0);

    // Extents (size)
    const uint64_t emuWidth  = static_cast<uint64_t>(width) * EMU_PER_PIXEL;
    const uint64_t emuHeight = static_cast<uint64_t>(height) * EMU_PER_PIXEL;

    XMLNode ext = anchor.append_child("xdr:ext");
    ext.append_attribute("cx").set_value(fmt::format("{}", emuWidth).c_str());
    ext.append_attribute("cy").set_value(fmt::format("{}", emuHeight).c_str());

    // graphicFrame node
    XMLNode frame = anchor.append_child("xdr:graphicFrame");

    // nvGraphicFramePr
    XMLNode nvGraphicFramePr = frame.append_child("xdr:nvGraphicFramePr");
    XMLNode cNvPr            = nvGraphicFramePr.append_child("xdr:cNvPr");
    auto    childCount       = static_cast<size_t>(std::distance(rootNode.children().begin(), rootNode.children().end()));
    cNvPr.append_attribute("id").set_value(fmt::format("{}", childCount + 1).c_str());
    cNvPr.append_attribute("name").set_value(std::string(name).c_str());
    nvGraphicFramePr.append_child("xdr:cNvGraphicFramePr");

    // xfrm
    XMLNode xfrm = frame.append_child("xdr:xfrm");
    XMLNode off  = xfrm.append_child("a:off");
    off.append_attribute("x").set_value("0");
    off.append_attribute("y").set_value("0");
    XMLNode extXfrm = xfrm.append_child("a:ext");
    extXfrm.append_attribute("cx").set_value("0");
    extXfrm.append_attribute("cy").set_value("0");

    // a:graphic
    XMLNode graphic     = frame.append_child("a:graphic");
    XMLNode graphicData = graphic.append_child("a:graphicData");
    graphicData.append_attribute("uri").set_value("http://schemas.openxmlformats.org/drawingml/2006/chart");

    XMLNode chartNode = graphicData.append_child("c:chart");
    chartNode.append_attribute("r:id").set_value(std::string(rId).c_str());

    anchor.append_child("xdr:clientData");
}

void XLDrawing::addScaledImage(std::string_view rId,
                               std::string_view name,
                               std::string_view description,
                               std::string_view data,
                               uint32_t         row,
                               uint32_t         col,
                               double           scalingFactor)
{
    auto [width, height] = getImageDimensions(std::string(data));
    if (width == 0 or height == 0) {
        width  = DEFAULT_IMAGE_SIZE;
        height = DEFAULT_IMAGE_SIZE;
    }

    const auto scaledWidth  = gsl::narrow_cast<uint32_t>(static_cast<double>(width) * scalingFactor);
    const auto scaledHeight = gsl::narrow_cast<uint32_t>(static_cast<double>(height) * scalingFactor);

    addImage(rId, name, description, row, col, scaledWidth, scaledHeight);
}

/**
 * @details fetches XLRelationships for the drawing - creates & assigns the class if empty
 */
XLRelationships& XLDrawing::relationships()
{
    if (!m_relationships or !m_relationships->valid()) {
        m_relationships = std::make_unique<XLRelationships>(parentDoc().drawingRelationships(getXmlPath()));
    }

    return *m_relationships;
}

XLDrawingItem::XLDrawingItem() : m_anchorNode(), m_parentDoc(nullptr) {}

XLDrawingItem::XLDrawingItem(const XMLNode& node, XLDocument* parentDoc) : m_anchorNode(node), m_parentDoc(parentDoc) {}

std::string XLDrawingItem::name() const
{ return m_anchorNode.child("xdr:pic").child("xdr:nvPicPr").child("xdr:cNvPr").attribute("name").value(); }

std::string XLDrawingItem::description() const
{ return m_anchorNode.child("xdr:pic").child("xdr:nvPicPr").child("xdr:cNvPr").attribute("descr").value(); }

uint32_t XLDrawingItem::row() const { 
    if (!m_anchorNode.child("xdr:from").empty()) { return m_anchorNode.child("xdr:from").child("xdr:row").text().as_uint(); }
    if (!m_anchorNode.child("xdr:row").empty()) { return m_anchorNode.child("xdr:row").text().as_uint(); }
    return 0;
}

uint32_t XLDrawingItem::col() const { 
    if (!m_anchorNode.child("xdr:from").empty()) { return m_anchorNode.child("xdr:from").child("xdr:col").text().as_uint(); }
    if (!m_anchorNode.child("xdr:col").empty()) { return m_anchorNode.child("xdr:col").text().as_uint(); }
    return 0;
}

uint32_t XLDrawingItem::width() const
{
    if (m_anchorNode.child("xdr:ext").empty()) return 0;
    uint64_t         emus = 0;
    std::string_view str  = m_anchorNode.child("xdr:ext").attribute("cx").value();
    std::from_chars(str.data(), str.data() + str.size(), emus);
    return static_cast<uint32_t>(emus / EMU_PER_PIXEL);
}

uint32_t XLDrawingItem::height() const
{
    if (m_anchorNode.child("xdr:ext").empty()) return 0;
    uint64_t         emus = 0;
    std::string_view str  = m_anchorNode.child("xdr:ext").attribute("cy").value();
    std::from_chars(str.data(), str.data() + str.size(), emus);
    return static_cast<uint32_t>(emus / EMU_PER_PIXEL);
}

std::string XLDrawingItem::relationshipId() const
{ return m_anchorNode.child("xdr:pic").child("xdr:blipFill").child("a:blip").attribute("r:embed").value(); }

std::vector<uint8_t> XLDrawingItem::imageBinary() const
{
    std::vector<uint8_t> buffer;
    if (!m_parentDoc) return buffer;
    
    std::string relId = relationshipId();
    if (relId.empty()) return buffer;

    std::string targetPath;
    for (const auto& entry : m_parentDoc->archive().entryNames()) {
        if (entry.find("_rels") != std::string::npos && entry.find("drawing") != std::string::npos) {
            std::string relsContent = m_parentDoc->archive().getEntry(entry);
            if (relsContent.find(relId) != std::string::npos) {
                pugi::xml_document relsDoc;
                relsDoc.load_string(relsContent.c_str());
                for (const auto& rel : relsDoc.document_element().children("Relationship")) {
                    if (std::string(rel.attribute("Id").value()) == relId) {
                        targetPath = rel.attribute("Target").value();
                        break;
                    }
                }
            }
        }
        if (!targetPath.empty()) break;
    }

    if (targetPath.empty()) return buffer;

    if (targetPath.find("../") == 0) {
        targetPath = "xl/" + targetPath.substr(3);
    } else if (targetPath.find("/") == 0) {
        targetPath = targetPath.substr(1);
    } else {
        targetPath = "xl/drawings/" + targetPath;
    }

    if (!m_parentDoc->archive().hasEntry(targetPath)) return buffer;

    std::string binData = m_parentDoc->archive().getEntry(targetPath);
    buffer.assign(binData.begin(), binData.end());
    return buffer;
}

uint32_t XLDrawing::imageCount() const
{
    uint32_t count = 0;
    for (const auto& child : xmlDocument().document_element().children()) {
        std::string nodeName = child.name();
        if (nodeName == "xdr:oneCellAnchor" or nodeName == "xdr:twoCellAnchor" or nodeName == "xdr:absoluteAnchor") {
            if (!child.child("xdr:pic").empty()) { count++; }
        }
    }
    return count;
}

XLDrawingItem XLDrawing::image(uint32_t index) const
{
    uint32_t count = 0;
    for (const auto& child : xmlDocument().document_element().children()) {
        std::string nodeName = child.name();
        if (nodeName == "xdr:oneCellAnchor" or nodeName == "xdr:twoCellAnchor" or nodeName == "xdr:absoluteAnchor") {
            if (!child.child("xdr:pic").empty()) {
                if (count == index) { return XLDrawingItem(child, const_cast<XLDocument*>(&parentDoc())); }
                count++;
            }
        }
    }
    throw XLException("XLDrawing::image: index out of bounds");
}

void XLDrawing::addChartAnchor(std::string_view rId, std::string_view name, uint32_t row, uint32_t col, XLDistance width, XLDistance height)
{
    XMLNode rootNode = xmlDocument().document_element();

    // Create oneCellAnchor
    XMLNode anchor = rootNode.append_child("xdr:oneCellAnchor");

    // From node (top-left)
    XMLNode from = anchor.append_child("xdr:from");
    from.append_child("xdr:col").text().set(col);
    from.append_child("xdr:colOff").text().set(0);
    from.append_child("xdr:row").text().set(row);
    from.append_child("xdr:rowOff").text().set(0);

    // Extents (size)
    XMLNode ext = anchor.append_child("xdr:ext");
    ext.append_attribute("cx").set_value(fmt::format("{}", width.getEMU()).c_str());
    ext.append_attribute("cy").set_value(fmt::format("{}", height.getEMU()).c_str());

    // Graphic Frame
    XMLNode frame = anchor.append_child("xdr:graphicFrame");
    frame.append_attribute("macro").set_value("");

    XMLNode nvPr = frame.append_child("xdr:nvGraphicFramePr");
    XMLNode cNvp = nvPr.append_child("xdr:cNvPr");
    cNvp.append_attribute("id").set_value("2");
    cNvp.append_attribute("name").set_value(std::string(name).c_str());
    nvPr.append_child("xdr:cNvGraphicFramePr");

    XMLNode xfrm = frame.append_child("xdr:xfrm");
    XMLNode off  = xfrm.append_child("a:off");
    off.append_attribute("x").set_value("0");
    off.append_attribute("y").set_value("0");
    XMLNode ext2 = xfrm.append_child("a:ext");
    ext2.append_attribute("cx").set_value(0);
    ext2.append_attribute("cy").set_value(0);

    XMLNode graphic = frame.append_child("a:graphic");
    XMLNode data    = graphic.append_child("a:graphicData");
    data.append_attribute("uri").set_value("http://schemas.openxmlformats.org/drawingml/2006/chart");

    XMLNode chartNode = data.append_child("c:chart");
    chartNode.append_attribute("xmlns:c").set_value("http://schemas.openxmlformats.org/drawingml/2006/chart");
    chartNode.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    chartNode.append_attribute("r:id").set_value(std::string(rId).c_str());

    anchor.append_child("xdr:clientData");
}

std::string vectorShapeTypeToString(XLVectorShapeType type)
{
    switch (type) {
        case XLVectorShapeType::Rectangle:
            return "rect";
        case XLVectorShapeType::Ellipse:
            return "ellipse";
        case XLVectorShapeType::Line:
            return "line";
        case XLVectorShapeType::Triangle:
            return "triangle";
        case XLVectorShapeType::RightTriangle:
            return "rtTriangle";
        case XLVectorShapeType::Arrow:
            return "rightArrow";
        case XLVectorShapeType::Diamond:
            return "diamond";
        case XLVectorShapeType::Parallelogram:
            return "parallelogram";
        case XLVectorShapeType::Hexagon:
            return "hexagon";
        case XLVectorShapeType::Star4:
            return "star4";
        case XLVectorShapeType::Star5:
            return "star5";
        case XLVectorShapeType::Star16:
            return "star16";
        case XLVectorShapeType::Star24:
            return "star24";
        case XLVectorShapeType::Heart:
            return "heart";
        case XLVectorShapeType::SmileyFace:
            return "smileyFace";
        case XLVectorShapeType::Cloud:
            return "cloud";
        case XLVectorShapeType::Donut:
            return "donut";
        case XLVectorShapeType::Ribbon:
            return "ribbon";
        case XLVectorShapeType::Sun:
            return "sun";
        case XLVectorShapeType::Moon:
            return "moon";
        case XLVectorShapeType::LightningBolt:
            return "lightningBolt";
        case XLVectorShapeType::FlowChartProcess:
            return "flowChartProcess";
        case XLVectorShapeType::FlowChartDecision:
            return "flowChartDecision";
        case XLVectorShapeType::FlowChartDocument:
            return "flowChartDocument";
        case XLVectorShapeType::FlowChartData:
            return "flowChartData";
        default:
            return "rect";
    }
}

void XLDrawing::addShape(uint32_t row, uint32_t col, const XLVectorShapeOptions& options)
{
    XMLNode rootNode = xmlDocument().document_element();
    
    XMLNode anchor;
    if (options.endRow.has_value() && options.endCol.has_value()) {
        anchor = rootNode.append_child("xdr:twoCellAnchor");
    } else {
        anchor = rootNode.append_child("xdr:oneCellAnchor");
    }

    XMLNode from = anchor.append_child("xdr:from");
    from.append_child("xdr:col").text().set(col);
    from.append_child("xdr:colOff").text().set(options.offsetX * EMU_PER_PIXEL);
    from.append_child("xdr:row").text().set(row);
    from.append_child("xdr:rowOff").text().set(options.offsetY * EMU_PER_PIXEL);

    const uint64_t emuWidth  = static_cast<uint64_t>(options.width) * EMU_PER_PIXEL;
    const uint64_t emuHeight = static_cast<uint64_t>(options.height) * EMU_PER_PIXEL;

    if (options.endRow.has_value() && options.endCol.has_value()) {
        XMLNode to = anchor.append_child("xdr:to");
        to.append_child("xdr:col").text().set(options.endCol.value());
        to.append_child("xdr:colOff").text().set(options.endOffsetX * EMU_PER_PIXEL);
        to.append_child("xdr:row").text().set(options.endRow.value());
        to.append_child("xdr:rowOff").text().set(options.endOffsetY * EMU_PER_PIXEL);
    } else {
        XMLNode ext = anchor.append_child("xdr:ext");
        ext.append_attribute("cx").set_value(fmt::format("{}", emuWidth).c_str());
        ext.append_attribute("cy").set_value(fmt::format("{}", emuHeight).c_str());
    }

    XMLNode sp = anchor.append_child("xdr:sp");
    sp.append_attribute("macro").set_value(options.macro.empty() ? "" : options.macro.c_str());
    sp.append_attribute("textlink").set_value("");

    XMLNode nvSpPr     = sp.append_child("xdr:nvSpPr");
    XMLNode cNvPr      = nvSpPr.append_child("xdr:cNvPr");
    auto    childCount = static_cast<size_t>(std::distance(rootNode.children().begin(), rootNode.children().end()));
    cNvPr.append_attribute("id").set_value(fmt::format("{}", childCount + 1).c_str());
    cNvPr.append_attribute("name").set_value(options.name.c_str());
    nvSpPr.append_child("xdr:cNvSpPr");

    XMLNode spPr = sp.append_child("xdr:spPr");
    XMLNode xfrm = spPr.append_child("a:xfrm");
    
    // Add rotation/flip properties
    if (options.rotation > 0) {
        xfrm.append_attribute("rot").set_value(fmt::format("{}", options.rotation * 60000).c_str());
    }
    if (options.flipH) {
        xfrm.append_attribute("flipH").set_value("1");
    }
    if (options.flipV) {
        xfrm.append_attribute("flipV").set_value("1");
    }

    xfrm.append_child("a:off").append_attribute("x").set_value("0");
    xfrm.child("a:off").append_attribute("y").set_value("0");
    xfrm.append_child("a:ext").append_attribute("cx").set_value(fmt::format("{}", emuWidth).c_str());
    xfrm.child("a:ext").append_attribute("cy").set_value(fmt::format("{}", emuHeight).c_str());

    XMLNode prstGeom = spPr.append_child("a:prstGeom");
    prstGeom.append_attribute("prst").set_value(vectorShapeTypeToString(options.type).c_str());
    prstGeom.append_child("a:avLst");

    if (!options.fillColor.empty()) {
        XMLNode solidFill = spPr.append_child("a:solidFill");
        solidFill.append_child("a:srgbClr").append_attribute("val").set_value(options.fillColor.c_str());
    }

    XMLNode ln = spPr.append_child("a:ln");
    ln.append_attribute("w").set_value(fmt::format("{}", static_cast<uint32_t>(options.lineWidth * EMU_PER_PIXEL)).c_str());
    if (!options.lineColor.empty()) {
        XMLNode solidFill = ln.append_child("a:solidFill");
        solidFill.append_child("a:srgbClr").append_attribute("val").set_value(options.lineColor.c_str());
    }
    else {
        ln.append_child("a:noFill");
    }
    
    // Advanced outline
    if (!options.lineDash.empty()) {
        ln.append_child("a:prstDash").append_attribute("val").set_value(options.lineDash.c_str());
    }
    if (!options.arrowStart.empty()) {
        ln.append_child("a:headEnd").append_attribute("type").set_value(options.arrowStart.c_str());
    }
    if (!options.arrowEnd.empty()) {
        ln.append_child("a:tailEnd").append_attribute("type").set_value(options.arrowEnd.c_str());
    }

    if (!options.text.empty() || !options.macro.empty() || options.richText.has_value()) {
        XMLNode txBody = sp.append_child("xdr:txBody");
        XMLNode bodyPr = txBody.append_child("a:bodyPr");
        
        // vertical alignment within shape
        if (!options.vertAlign.empty()) {
            bodyPr.append_attribute("anchor").set_value(options.vertAlign.c_str());
        } else if (!options.macro.empty()) {
            bodyPr.append_attribute("anchor").set_value("ctr");
        }

        if (!options.macro.empty()) {
            bodyPr.append_attribute("vertOverflow").set_value("clip");
            bodyPr.append_attribute("wrap").set_value("square");
        }
        
        txBody.append_child("a:lstStyle");
        XMLNode p = txBody.append_child("a:p");

        // Horizontal alignment
        if (!options.horzAlign.empty()) {
            p.append_child("a:pPr").append_attribute("algn").set_value(options.horzAlign.c_str());
        }

        if (options.richText.has_value()) {
            const auto& rt = options.richText.value();
            for (const auto& run : rt.runs()) {
                XMLNode rNode = p.append_child("a:r");
                XMLNode rPrNode = rNode.append_child("a:rPr");
                
                // For a:rPr we must set sz (in 100ths of a point) and lang
                uint32_t fontSize = run.fontSize().value_or(11) * 100;
                rPrNode.append_attribute("sz").set_value(fmt::format("{}", fontSize).c_str());
                rPrNode.append_attribute("lang").set_value("en-US");
                
                if (run.bold().value_or(false)) rPrNode.append_attribute("b").set_value("1");
                if (run.italic().value_or(false)) rPrNode.append_attribute("i").set_value("1");
                
                if (run.fontColor().has_value()) {
                    XMLNode solidFill = rPrNode.append_child("a:solidFill");
                    std::string hexColor = run.fontColor()->hex();
                    if (hexColor.length() == 8) {
                        hexColor = hexColor.substr(2); // Strip ARGB alpha channel -> RGB
                    }
                    solidFill.append_child("a:srgbClr").append_attribute("val").set_value(hexColor.c_str());
                }
                
                XMLNode tNode = rNode.append_child("a:t");
                tNode.text().set(run.text().c_str());
            }
        } else if (!options.text.empty()) {
            XMLNode r = p.append_child("a:r");
            XMLNode rPr = r.append_child("a:rPr");
            rPr.append_attribute("lang").set_value("en-US");
            rPr.append_attribute("sz").set_value("1100"); // 11pt default
            XMLNode t = r.append_child("a:t");
            t.text().set(options.text.c_str());
        }
        p.append_child("a:endParaRPr").append_attribute("lang").set_value("en-US");
    }

    anchor.append_child("xdr:clientData");
}

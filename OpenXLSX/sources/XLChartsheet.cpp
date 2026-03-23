#include "XLChartsheet.hpp"
#include "XLXmlData.hpp"
#include "XLDocument.hpp"
#include <charconv>
#include "XLUtilities.hpp"

using namespace OpenXLSX;

XLChartsheet::XLChartsheet(XLXmlData* xmlData) : XLSheetBase(xmlData) {}
XLChartsheet::~XLChartsheet() = default;

uint16_t XLChartsheet::sheetXmlNumber() const
{
    constexpr const char* searchPattern = "xl/chartsheets/sheet";
    std::string           xmlPath       = getXmlPath();
    size_t                pos           = xmlPath.find(searchPattern);
    if (pos == std::string::npos) return 0;
    pos += strlen(searchPattern);
    size_t pos2 = pos;
    while (std::isdigit(xmlPath[pos2])) ++pos2;
    if (pos2 == pos or xmlPath.substr(pos2) != ".xml") return 0;
    
    uint16_t num = 0;
    std::string_view numStr(xmlPath.data() + pos, pos2 - pos);
    std::from_chars(numStr.data(), numStr.data() + numStr.size(), num);
    return num;
}

XLRelationships& XLChartsheet::relationships()
{
    if (!m_relationships.valid()) {
        m_relationships = parentDoc().sheetRelationships(sheetXmlNumber(), true);
    }
    if (!m_relationships.valid()) throw XLException("XLChartsheet::relationships(): could not create relationships XML");
    return m_relationships;
}

XLDrawing& XLChartsheet::drawing()
{
    if (!m_drawing.valid()) {
        XMLNode docElement = xmlDocument().document_element();
        std::ignore        = relationships();

        // Find existing drawing relationship if any
        XLRelationshipItem drawingRelationship;
        for (auto rel : m_relationships.relationships()) {
            if (rel.type() == XLRelationshipType::Drawing) {
                drawingRelationship = rel;
                break;
            }
        }

        if (drawingRelationship.empty()) {
            // Create a new drawing file
            m_drawing = parentDoc().createDrawing();
            if (!m_drawing.valid()) throw XLException("XLChartsheet::drawing(): could not create drawing XML");

            std::string drawingPath         = m_drawing.getXmlPath();
            std::string drawingRelativePath = getPathARelativeToPathB(drawingPath, getXmlPath());

            drawingRelationship = m_relationships.addRelationship(XLRelationshipType::Drawing, drawingRelativePath);
            if (drawingRelationship.empty()) throw XLException("XLChartsheet::drawing(): could not add sheet relationship for Drawing");

            XMLNode drawingNode = docElement.child("drawing");
            if (drawingNode.empty()) {
                drawingNode = docElement.append_child("drawing");
            }
            if (drawingNode.attribute("r:id").empty()) {
                drawingNode.append_attribute("r:id").set_value(drawingRelationship.id().c_str());
            } else {
                drawingNode.attribute("r:id").set_value(drawingRelationship.id().c_str());
            }
        } else {
            // Load existing drawing
            std::string drawingPath = drawingRelationship.target();
            if (drawingPath.front() != '/') {
                std::string sheetDir = getXmlPath();
                sheetDir = sheetDir.substr(0, sheetDir.find_last_of('/'));
                drawingPath = sheetDir + "/" + drawingPath;
                drawingPath = eliminateDotAndDotDotFromPath(drawingPath);
            }
            if (drawingPath.front() == '/') drawingPath = drawingPath.substr(1);
            m_drawing = parentDoc().drawing(drawingPath);
        }
    }
    return m_drawing;
}

XLChart XLChartsheet::addChart(XLChartType type, std::string_view name)
{
    // 1. Create Chart File in Document
    XLChart chart = parentDoc().createChart(type);

    // 2. Get Drawing for the chartsheet
    XLDrawing& drw = drawing();
    std::string drawingPath       = drw.getXmlPath();
    std::string chartRelativePath = getPathARelativeToPathB(chart.getXmlPath(), drawingPath);

    // 3. Add Relationship in Drawing
    XLRelationshipItem chartRel;
    if (!drw.relationships().targetExists(chartRelativePath)) {
        chartRel = drw.relationships().addRelationship(XLRelationshipType::Chart, chartRelativePath);
    } else {
        chartRel = drw.relationships().relationshipByTarget(chartRelativePath);
    }

    // 4. Add Absolute Anchor in Drawing (specific to Chartsheet)
    XMLDocument& drwDoc = drw.xmlDocument();
    XMLNode wsDrNode = drwDoc.document_element();
    if (wsDrNode.empty()) {
        constexpr std::string_view drwTemplate = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>)";
        drwDoc.load_string(drwTemplate.data(), pugi_parse_settings);
        wsDrNode = drwDoc.document_element();
    }
    
    // Clear any existing anchors to ensure chartsheet only has one main chart
    for (auto child = wsDrNode.first_child(); child; ) {
        auto next = child.next_sibling();
        wsDrNode.remove_child(child);
        child = next;
    }

    XMLNode anchor = wsDrNode.append_child("xdr:absoluteAnchor");
    XMLNode posNode = anchor.append_child("xdr:pos");
    posNode.append_attribute("x").set_value("0");
    posNode.append_attribute("y").set_value("0");
    
    XMLNode extNode = anchor.append_child("xdr:ext");
    extNode.append_attribute("cx").set_value("0");
    extNode.append_attribute("cy").set_value("0");

    XMLNode graphicFrame = anchor.append_child("xdr:graphicFrame");
    graphicFrame.append_attribute("macro").set_value("");

    XMLNode nvGraphicFramePr = graphicFrame.append_child("xdr:nvGraphicFramePr");
    XMLNode cNvPr = nvGraphicFramePr.append_child("xdr:cNvPr");
    cNvPr.append_attribute("id").set_value(2);
    cNvPr.append_attribute("name").set_value(std::string(name).c_str());
    nvGraphicFramePr.append_child("xdr:cNvGraphicFramePr");

    XMLNode xfrm = graphicFrame.append_child("xdr:xfrm");
    XMLNode off = xfrm.append_child("a:off");
    off.append_attribute("x").set_value("0");
    off.append_attribute("y").set_value("0");
    XMLNode ext = xfrm.append_child("a:ext");
    ext.append_attribute("cx").set_value("0");
    ext.append_attribute("cy").set_value("0");

    XMLNode graphic = graphicFrame.append_child("a:graphic");
    XMLNode graphicData = graphic.append_child("a:graphicData");
    graphicData.append_attribute("uri").set_value("http://schemas.openxmlformats.org/drawingml/2006/chart");

    XMLNode cChart = graphicData.append_child("c:chart");
    cChart.append_attribute("xmlns:c").set_value("http://schemas.openxmlformats.org/drawingml/2006/chart");
    cChart.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    cChart.append_attribute("r:id").set_value(chartRel.id().c_str());

    XMLNode clientData = anchor.append_child("xdr:clientData");
    clientData.append_attribute("fLocksWithSheet").set_value("1");
    clientData.append_attribute("fPrintsWithSheet").set_value("1");

    return chart;
}

XLColor XLChartsheet::getColor_impl() const
{
    auto node = xmlDocument().document_element().child("sheetPr").child("tabColor");
    if (node.empty()) return XLColor();
    return XLColor(node.attribute("rgb").value());
}

void XLChartsheet::setColor_impl(const XLColor& color) { setTabColor(xmlDocument(), color); }

bool XLChartsheet::isSelected_impl() const { return tabIsSelected(xmlDocument()); }

void XLChartsheet::setSelected_impl(bool selected) { setTabSelected(xmlDocument(), selected); }

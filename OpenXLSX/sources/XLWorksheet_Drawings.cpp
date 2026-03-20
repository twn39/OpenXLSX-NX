#include <string>
#include <vector>

#include "XLWorksheet.hpp"
#include "XLDocument.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

bool XLWorksheet::hasRelationships() const { return parentDoc().hasSheetRelationships(sheetXmlNumber()); }
bool XLWorksheet::hasDrawing() const { return parentDoc().hasSheetDrawing(sheetXmlNumber()); }
bool XLWorksheet::hasVmlDrawing() const { return parentDoc().hasSheetVmlDrawing(sheetXmlNumber()); }
bool XLWorksheet::hasComments() const { return parentDoc().hasSheetComments(sheetXmlNumber()); }
bool XLWorksheet::hasTables() const { return parentDoc().hasSheetTables(sheetXmlNumber()); }

XLRelationships& XLWorksheet::relationships()
{
    if (!m_relationships.valid()) {
        m_relationships = parentDoc().sheetRelationships(sheetXmlNumber());
    }
    if (!m_relationships.valid()) throw XLException("XLWorksheet::relationships(): could not create relationships XML");
    return m_relationships;
}

XLDrawing& XLWorksheet::drawing()
{
    if (!m_drawing.valid()) {
        XMLNode docElement = xmlDocument().document_element();
        std::ignore        = relationships();

        uint16_t sheetXmlNo = sheetXmlNumber();
        m_drawing           = parentDoc().sheetDrawing(sheetXmlNo);
        if (!m_drawing.valid()) throw XLException("XLWorksheet::drawing(): could not create drawing XML");

        std::string drawingPath         = m_drawing.getXmlPath();
        std::string drawingRelativePath = getPathARelativeToPathB(drawingPath, getXmlPath());

        XLRelationshipItem drawingRelationship;
        if (!m_relationships.targetExists(drawingRelativePath))
            drawingRelationship = m_relationships.addRelationship(XLRelationshipType::Drawing, drawingRelativePath);
        else
            drawingRelationship = m_relationships.relationshipByTarget(drawingRelativePath);

        if (drawingRelationship.empty()) throw XLException("XLWorksheet::drawing(): could not add sheet relationship for Drawing");

        if (docElement.child("drawing").empty()) {
            XMLNode drawingNode = appendAndGetNode(docElement, "drawing", m_nodeOrder);
            appendAndSetAttribute(drawingNode, "r:id", drawingRelationship.id());
        }
    }

    return m_drawing;
}

void XLWorksheet::addImage(const std::string& name, const std::string& data, uint32_t row, uint32_t col, uint32_t width, uint32_t height)
{
    std::string internalPath = parentDoc().addImage(name, data);
    XLDrawing& drw = drawing();
    std::string drawingPath       = drw.getXmlPath();
    std::string imageRelativePath = getPathARelativeToPathB(internalPath, drawingPath);

    XLRelationshipItem imgRel;
    if (!drw.relationships().targetExists(imageRelativePath))
        imgRel = drw.relationships().addRelationship(XLRelationshipType::Image, imageRelativePath);
    else
        imgRel = drw.relationships().relationshipByTarget(imageRelativePath);

    drw.addImage(imgRel.id(), name, "", row, col, width, height);
}

void XLWorksheet::addScaledImage(const std::string& name, const std::string& data, uint32_t row, uint32_t col, double scalingFactor)
{
    std::string internalPath = parentDoc().addImage(name, data);
    XLDrawing& drw = drawing();
    std::string drawingPath       = drw.getXmlPath();
    std::string imageRelativePath = getPathARelativeToPathB(internalPath, drawingPath);

    XLRelationshipItem imgRel;
    if (!drw.relationships().targetExists(imageRelativePath))
        imgRel = drw.relationships().addRelationship(XLRelationshipType::Image, imageRelativePath);
    else
        imgRel = drw.relationships().relationshipByTarget(imageRelativePath);

    drw.addScaledImage(imgRel.id(), name, "", data, row, col, scalingFactor);
}

XLChart XLWorksheet::addChart([[maybe_unused]] XLChartType type, std::string_view name, uint32_t row, uint32_t col, uint32_t width, uint32_t height)
{
    // 1. Create Chart File in Document
    XLChart chart = parentDoc().createChart();

    // 2. Get Drawing for the sheet
    XLDrawing& drw = drawing();
    std::string drawingPath     = drw.getXmlPath();
    std::string chartRelativePath = getPathARelativeToPathB(chart.getXmlPath(), drawingPath);

    // 3. Add Relationship in drawing to the new chart file
    XLRelationshipItem chartRel;
    if (!drw.relationships().targetExists(chartRelativePath)) {
        chartRel = drw.relationships().addRelationship(XLRelationshipType::Chart, chartRelativePath);
    } else {
        chartRel = drw.relationships().relationshipByTarget(chartRelativePath);
    }

    // 4. Add Chart Anchor in Drawing
    drw.addChartAnchor(chartRel.id(), name, row, col, width, height);

    return chart;
}

std::vector<XLDrawingItem> XLWorksheet::images()
{
    std::vector<XLDrawingItem> result;
    if (!m_drawing.valid()) {
        std::string drawingRelativePath = "";
        for (auto& rel : relationships().relationships()) {
            if (rel.type() == XLRelationshipType::Drawing) {
                drawingRelativePath = rel.target();
                break;
            }
        }
        if (drawingRelativePath.empty()) return result;
        m_drawing = drawing();
    }

    uint32_t count = m_drawing.imageCount();
    for (uint32_t i = 0; i < count; ++i) { result.push_back(m_drawing.image(i)); }
    return result;
}

XLVmlDrawing& XLWorksheet::vmlDrawing()
{
    if (!m_vmlDrawing.valid()) {
        XMLNode docElement = xmlDocument().document_element();
        std::ignore        = relationships();

        uint16_t sheetXmlNo = sheetXmlNumber();
        m_vmlDrawing        = parentDoc().sheetVmlDrawing(sheetXmlNo);
        if (!m_vmlDrawing.valid()) throw XLException("XLWorksheet::vmlDrawing(): could not create drawing XML");
        std::string        drawingRelativePath = getPathARelativeToPathB(m_vmlDrawing.getXmlPath(), getXmlPath());
        XLRelationshipItem vmlDrawingRelationship;
        if (!m_relationships.targetExists(drawingRelativePath))
            vmlDrawingRelationship = m_relationships.addRelationship(XLRelationshipType::VMLDrawing, drawingRelativePath);
        else
            vmlDrawingRelationship = m_relationships.relationshipByTarget(drawingRelativePath);
        if (vmlDrawingRelationship.empty())
            throw XLException("XLWorksheet::vmlDrawing(): could not add determine sheet relationship for VML Drawing");
        XMLNode legacyDrawing = appendAndGetNode(docElement, "legacyDrawing", m_nodeOrder);
        if (legacyDrawing.empty()) throw XLException("XLWorksheet::vmlDrawing(): could not add <legacyDrawing> element to worksheet XML");
        appendAndSetAttribute(legacyDrawing, "r:id", vmlDrawingRelationship.id());
    }

    return m_vmlDrawing;
}

XLComments& XLWorksheet::comments()
{
    if (!m_comments.valid()) {
        std::ignore = relationships();
        std::ignore = vmlDrawing();

        uint16_t sheetXmlNo = sheetXmlNumber();
        m_comments = parentDoc().sheetComments(sheetXmlNo);
        if (!m_comments.valid()) throw XLException("XLWorksheet::comments(): could not create comments XML");
        m_comments.setVmlDrawing(m_vmlDrawing);
        std::string commentsRelativePath = getPathARelativeToPathB(m_comments.getXmlPath(), getXmlPath());
        if (!m_relationships.targetExists(commentsRelativePath))
            m_relationships.addRelationship(XLRelationshipType::Comments, commentsRelativePath);
    }

    return m_comments;
}

XLTableCollection& XLWorksheet::tables()
{
    if (!m_tables.valid()) {
        m_tables = XLTableCollection(xmlDocument().document_element(), this);
    }

    return m_tables;
}

void XLWorksheet::addHyperlink(std::string_view cellRef, std::string_view url, std::string_view tooltip)
{
    removeHyperlink(cellRef);
    std::ignore = relationships();
    const auto rel = m_relationships.addRelationship(XLRelationshipType::Hyperlink, std::string(url), true);
    XMLNode docElement     = xmlDocument().document_element();
    XMLNode hyperlinksNode = docElement.child("hyperlinks");
    if (hyperlinksNode.empty()) { hyperlinksNode = appendAndGetNode(docElement, "hyperlinks", m_nodeOrder); }
    XMLNode hyperlinkNode = hyperlinksNode.append_child("hyperlink");
    hyperlinkNode.append_attribute("ref").set_value(std::string(cellRef).c_str());
    hyperlinkNode.append_attribute("r:id").set_value(rel.id().c_str());
    if (!tooltip.empty()) { hyperlinkNode.append_attribute("tooltip").set_value(std::string(tooltip).c_str()); }
}

void XLWorksheet::addInternalHyperlink(std::string_view cellRef, std::string_view location, std::string_view tooltip)
{
    removeHyperlink(cellRef);
    XMLNode docElement     = xmlDocument().document_element();
    XMLNode hyperlinksNode = docElement.child("hyperlinks");
    if (hyperlinksNode.empty()) { hyperlinksNode = appendAndGetNode(docElement, "hyperlinks", m_nodeOrder); }
    XMLNode hyperlinkNode = hyperlinksNode.append_child("hyperlink");
    hyperlinkNode.append_attribute("ref").set_value(std::string(cellRef).c_str());
    hyperlinkNode.append_attribute("location").set_value(std::string(location).c_str());
    hyperlinkNode.append_attribute("display").set_value(std::string(location).c_str());
    if (!tooltip.empty()) { hyperlinkNode.append_attribute("tooltip").set_value(std::string(tooltip).c_str()); }
}

bool XLWorksheet::hasHyperlink(std::string_view cellRef) const
{
    XMLNode docElement     = xmlDocument().document_element();
    XMLNode hyperlinksNode = docElement.child("hyperlinks");
    if (hyperlinksNode.empty()) return false;
    for (auto link : hyperlinksNode.children("hyperlink")) {
        if (std::string_view(link.attribute("ref").value()) == cellRef) return true;
    }
    return false;
}

std::string XLWorksheet::getHyperlink(std::string_view cellRef) const
{
    XMLNode docElement     = xmlDocument().document_element();
    XMLNode hyperlinksNode = docElement.child("hyperlinks");
    if (hyperlinksNode.empty()) return "";
    for (auto link : hyperlinksNode.children("hyperlink")) {
        if (std::string_view(link.attribute("ref").value()) == cellRef) {
            if (link.attribute("location")) return link.attribute("location").value();
            if (link.attribute("r:id")) {
                const auto rId = link.attribute("r:id").value();
                return const_cast<XLWorksheet*>(this)->relationships().relationshipById(rId).target();
            }
        }
    }
    return "";
}

void XLWorksheet::removeHyperlink(std::string_view cellRef)
{
    XMLNode docElement     = xmlDocument().document_element();
    XMLNode hyperlinksNode = docElement.child("hyperlinks");
    if (hyperlinksNode.empty()) return;
    for (auto link : hyperlinksNode.children("hyperlink")) {
        if (std::string_view(link.attribute("ref").value()) == cellRef) {
            hyperlinksNode.remove_child(link);
            break;
        }
    }
    if (hyperlinksNode.first_child().empty()) { docElement.remove_child(hyperlinksNode); }
}


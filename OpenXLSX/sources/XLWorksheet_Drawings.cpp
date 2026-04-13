#include <string>
#include <vector>

#include "XLDocument.hpp"
#include "XLUtilities.hpp"
#include "XLWorksheet.hpp"

using namespace OpenXLSX;

bool XLWorksheet::hasRelationships() const { return parentDoc().hasSheetRelationships(sheetXmlNumber()); }
bool XLWorksheet::hasDrawing() const { return parentDoc().hasSheetDrawing(sheetXmlNumber()); }
bool XLWorksheet::hasVmlDrawing() const { return parentDoc().hasSheetVmlDrawing(sheetXmlNumber()); }
bool XLWorksheet::hasComments() const { return parentDoc().hasSheetComments(sheetXmlNumber()); }
bool XLWorksheet::hasThreadedComments() const { return parentDoc().hasSheetThreadedComments(sheetXmlNumber()); }
bool XLWorksheet::hasTables() const { return parentDoc().hasSheetTables(sheetXmlNumber()); }

XLRelationships& XLWorksheet::relationships()
{
    if (!m_relationships.valid()) { m_relationships = parentDoc().sheetRelationships(sheetXmlNumber()); }
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

void XLWorksheet::addImage(const std::string&    name,
                           const std::string&    data,
                           uint32_t              row,
                           uint32_t              col,
                           uint32_t              width,
                           uint32_t              height,
                           const XLImageOptions& options)
{
    std::string internalPath      = parentDoc().addImage(name, data);
    XLDrawing&  drw               = drawing();
    std::string drawingPath       = drw.getXmlPath();
    std::string imageRelativePath = getPathARelativeToPathB(internalPath, drawingPath);

    XLRelationshipItem imgRel;
    if (!drw.relationships().targetExists(imageRelativePath))
        imgRel = drw.relationships().addRelationship(XLRelationshipType::Image, imageRelativePath);
    else
        imgRel = drw.relationships().relationshipByTarget(imageRelativePath);

    drw.addImage(imgRel.id(), name, "", row, col, width, height, options);
}

void XLWorksheet::addScaledImage(const std::string& name, const std::string& data, uint32_t row, uint32_t col, double scalingFactor)
{
    std::string internalPath      = parentDoc().addImage(name, data);
    XLDrawing&  drw               = drawing();
    std::string drawingPath       = drw.getXmlPath();
    std::string imageRelativePath = getPathARelativeToPathB(internalPath, drawingPath);

    XLRelationshipItem imgRel;
    if (!drw.relationships().targetExists(imageRelativePath))
        imgRel = drw.relationships().addRelationship(XLRelationshipType::Image, imageRelativePath);
    else
        imgRel = drw.relationships().relationshipByTarget(imageRelativePath);

    drw.addScaledImage(imgRel.id(), name, "", data, row, col, scalingFactor);
}

XLChart XLWorksheet::addChart(XLChartType type, std::string_view name, uint32_t row, uint32_t col, uint32_t width, uint32_t height)
{
    // 1. Create Chart File in Document
    XLChart chart = parentDoc().createChart(type);

    // 2. Get Drawing for the sheet
    XLDrawing&  drw               = drawing();
    std::string drawingPath       = drw.getXmlPath();
    std::string chartRelativePath = getPathARelativeToPathB(chart.getXmlPath(), drawingPath);

    // 3. Add Relationship in drawing to the new chart file
    XLRelationshipItem chartRel;
    if (!drw.relationships().targetExists(chartRelativePath)) {
        chartRel = drw.relationships().addRelationship(XLRelationshipType::Chart, chartRelativePath);
    }
    else {
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
        m_comments          = parentDoc().sheetComments(sheetXmlNo);
        if (!m_comments.valid()) throw XLException("XLWorksheet::comments(): could not create comments XML");
        m_comments.setVmlDrawing(m_vmlDrawing);
        std::string commentsRelativePath = getPathARelativeToPathB(m_comments.getXmlPath(), getXmlPath());
        if (!m_relationships.targetExists(commentsRelativePath))
            m_relationships.addRelationship(XLRelationshipType::Comments, commentsRelativePath);
    }

    return m_comments;
}

XLThreadedComments& XLWorksheet::threadedComments()
{
    if (!m_threadedComments.valid()) {
        std::ignore        = relationships();
        uint16_t sheetXmlNo = sheetXmlNumber();
        m_threadedComments  = parentDoc().sheetThreadedComments(sheetXmlNo);
        if (!m_threadedComments.valid()) throw XLException("XLWorksheet::threadedComments(): could not create threadedComments XML");

        std::string commentsPath         = m_threadedComments.getXmlPath();
        std::string commentsRelativePath = getPathARelativeToPathB(commentsPath, getXmlPath());

        if (!m_relationships.targetExists(commentsRelativePath))
            m_relationships.addRelationship(XLRelationshipType::ThreadedComments, commentsRelativePath);
    }
    return m_threadedComments;
}

XLTableCollection& XLWorksheet::tables()
{
    if (!m_tables.valid()) { m_tables = XLTableCollection(xmlDocument().document_element(), this); }

    return m_tables;
}

void XLWorksheet::addHyperlink(std::string_view cellRef, std::string_view url, std::string_view tooltip)
{
    removeHyperlink(cellRef);
    std::ignore               = relationships();
    const auto rel            = m_relationships.addRelationship(XLRelationshipType::Hyperlink, std::string(url), true);
    XMLNode    docElement     = xmlDocument().document_element();
    XMLNode    hyperlinksNode = docElement.child("hyperlinks");
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

XLChart XLWorksheet::addChart(XLChartType type, const XLChartAnchor& anchor)
{
    // 1. Create Chart File in Document
    XLChart chart = parentDoc().createChart(type);

    // 2. Get Drawing for the sheet
    XLDrawing&  drw               = drawing();
    std::string drawingPath       = drw.getXmlPath();
    std::string chartRelativePath = getPathARelativeToPathB(chart.getXmlPath(), drawingPath);

    // 3. Add Relationship in drawing to the new chart file
    XLRelationshipItem chartRel;
    if (!drw.relationships().targetExists(chartRelativePath)) {
        chartRel = drw.relationships().addRelationship(XLRelationshipType::Chart, chartRelativePath);
    }
    else {
        chartRel = drw.relationships().relationshipByTarget(chartRelativePath);
    }

    // 4. Add Chart Anchor in Drawing
    drw.addChartAnchor(chartRel.id(), anchor.name, anchor.row, anchor.col, anchor.width, anchor.height);

    return chart;
}

void XLWorksheet::addNote(std::string_view cellRef, std::string_view text, std::string_view author)
{
    std::string finalAuthor = author.empty() ? parentDoc().defaultAuthor() : std::string(author);
    comments().set(std::string(cellRef), std::string(text), finalAuthor);
}

XLThreadedComment XLWorksheet::addComment(std::string_view cellRef, std::string_view text, std::string_view author)
{
    std::string finalAuthor = author.empty() ? parentDoc().defaultAuthor() : std::string(author);
    std::string personId = parentDoc().persons().addPerson(finalAuthor);
    auto tc = threadedComments().addComment(std::string(cellRef), personId, std::string(text));
    
    // Add legacy note for fallback and to generate the hidden VML box required by Excel.
    // Excel strictly requires the legacy author name to be "tc=" + the threaded comment ID!
    std::string legacyAuthor = "tc=" + tc.id();
    addNote(cellRef, text, legacyAuthor);

    // Inject extLst to signal Excel that threaded comments exist here
    XMLNode root = xmlDocument().document_element();
    XMLNode extLst = root.child("extLst");
    if (!extLst) {
        extLst = root.append_child("extLst");
    }
    bool hasThreadedExt = false;
    for (XMLNode ext = extLst.child("ext"); ext; ext = ext.next_sibling("ext")) {
        if (std::string(ext.attribute("uri").value()) == "{C5A4FD07-D465-4E85-B267-3A3644040A51}") {
            hasThreadedExt = true;
            break;
        }
    }
    if (!hasThreadedExt) {
        XMLNode ext = extLst.append_child("ext");
        ext.append_attribute("uri").set_value("{C5A4FD07-D465-4E85-B267-3A3644040A51}");
        ext.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
        ext.append_child("x14:threadedComments");
    }
    
    // Set the worksheet reference for fluent replies
    tc.setWorksheet(this); return tc;
}

XLThreadedComment XLWorksheet::addReply(const std::string& parentId, const std::string& text, const std::string& author)
{
    std::string finalAuthor = author.empty() ? parentDoc().defaultAuthor() : author;
    std::string personId = parentDoc().persons().addPerson(finalAuthor);
    auto tc = threadedComments().addReply(parentId, personId, text);
    
    // Attempt to sync the legacy comment fallback string
    std::string refStr = tc.ref();
    if (!refStr.empty()) {
        std::string fallbackText = comments().get(refStr);
        fallbackText += "\nReply: " + text;
        
        // Retain the existing 'tc=' author from the parent
        std::string parentAuthor = "tc=" + parentId;
        addNote(refStr, fallbackText, parentAuthor);
    }

    tc.setWorksheet(this); return tc;
}

void XLWorksheet::deleteComment(std::string_view cellRef)
{
    if (hasThreadedComments()) {
        threadedComments().deleteComment(std::string(cellRef));
    }
    deleteNote(cellRef); // remove the legacy hidden box
}

void XLWorksheet::deleteNote(std::string_view cellRef)
{
    if (hasComments()) {
        comments().deleteComment(std::string(cellRef));
    }
}

void XLWorksheet::insertImage(const std::string& cellReference, const std::string& imagePath)
{
    XLImageOptions opts;
    insertImage(cellReference, imagePath, opts);
}

#include "XLEMUConverter.hpp"
#include "XLImageParser.hpp"
#include <fstream>
#include <sstream>

void XLWorksheet::insertImage(const std::string& cellReference, const std::string& imagePath, const XLImageOptions& options)
{
    XLImageSize size = XLImageParser::parseDimensions(imagePath);
    if (!size.valid) throw XLInputError("Invalid or unsupported image file");

    // Read file data into string
    std::ifstream file(imagePath, std::ios::binary);
    if (!file) throw XLInputError("Could not open image file");
    std::ostringstream ss;
    ss << file.rdbuf();
    std::string data = ss.str();

    // Generate internal name
    std::string name = "image_openxlsx." + size.extension;

    // Parse coordinate
    XLCellReference ref{std::string(cellReference)};
    uint32_t        row = ref.row() - 1;    // 0-indexed for addImage
    uint16_t        col = ref.column() - 1;

    // Scaled dimensions
    uint32_t scaledWidth  = static_cast<uint32_t>(size.width * options.scaleX);
    uint32_t scaledHeight = static_cast<uint32_t>(size.height * options.scaleY);

    // Call addImage with options to correctly generate TwoCell or Absolute anchors
    addImage(name, data, row, col, scaledWidth, scaledHeight, options);
}

void XLWorksheet::addShape(std::string_view cellReference, const XLVectorShapeOptions& options)
{
    XLCellReference ref{std::string(cellReference)};
    uint32_t        row = ref.row() - 1;    // 0-indexed for drawing
    uint16_t        col = ref.column() - 1;

    XLDrawing& drw = drawing();
    drw.addShape(row, col, options);
}

void XLWorksheet::addTableSlicer(std::string_view       cellReference,
                                 const XLTable&         table,
                                 std::string_view       columnName,
                                 const XLSlicerOptions& options)
{
    uint32_t tableId = table.id();
    auto     col     = table.column(columnName);
    uint32_t colId   = col.id();

    std::string name      = options.name.empty() ? std::string(columnName) : options.name;
    std::string caption   = options.caption.empty() ? std::string(columnName) : options.caption;
    std::string cacheName = "Slicer_" + name;

    // 1. Create Slicer Cache and Slicer files
    parentDoc().createTableSlicerCache(tableId, colId, cacheName, columnName);
    std::string slicerFilename = parentDoc().createSlicer(name, cacheName, caption);

    // 2. Add Slicer relationship to worksheet
    relationships().addRelationship(XLRelationshipType::Slicer, "../slicers/" + slicerFilename.substr(11));
    std::string rId = relationships().relationshipByTarget("../slicers/" + slicerFilename.substr(11)).id();

    // 3. Add extLst to worksheet.xml
    XMLNode sheetNode = xmlDocument().document_element();
    XMLNode extLst    = sheetNode.child("extLst");
    if (!sheetNode.attribute("xmlns:mc"))
        sheetNode.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
    if (!sheetNode.attribute("xmlns:x15"))
        sheetNode.append_attribute("xmlns:x15").set_value("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
    if (!sheetNode.attribute("mc:Ignorable")) { sheetNode.append_attribute("mc:Ignorable").set_value("x15"); }
    else {
        std::string ign = sheetNode.attribute("mc:Ignorable").value();
        if (ign.find("x15") == std::string::npos) { sheetNode.attribute("mc:Ignorable").set_value((ign + " x15").c_str()); }
    }

    if (extLst.empty()) extLst = sheetNode.append_child("extLst");

    XMLNode ext = extLst.find_child_by_attribute("uri", "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}");
    if (ext.empty()) {
        ext = extLst.append_child("ext");
        ext.append_attribute("xmlns:x15").set_value("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
        ext.append_attribute("uri").set_value("{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}");
    }

    XMLNode slicerList = ext.child("x14:slicerList");
    if (slicerList.empty()) {
        slicerList = ext.append_child("x14:slicerList");
        // Need to add x14 namespace to root or here
        slicerList.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
    }

    slicerList.append_child("x14:slicer").append_attribute("r:id").set_value(rId.c_str());

    // 4. Add drawing for the slicer
    XLCellReference ref{std::string(cellReference)};
    uint32_t        row    = ref.row() - 1;
    uint16_t        colIdx = ref.column() - 1;

    XLDrawing& drw     = drawing();
    XMLNode    drwRoot = drw.xmlDocument().document_element();

    XMLNode anchor = drwRoot.append_child("xdr:twoCellAnchor");

    XMLNode from = anchor.append_child("xdr:from");
    from.append_child("xdr:col").text().set(colIdx);
    from.append_child("xdr:colOff").text().set(options.offsetX * 9525);
    from.append_child("xdr:row").text().set(row);
    from.append_child("xdr:rowOff").text().set(options.offsetY * 9525);

    XMLNode to = anchor.append_child("xdr:to");
    to.append_child("xdr:col").text().set(colIdx + 2);    // Approximate to
    to.append_child("xdr:colOff").text().set(options.offsetX * 9525);
    to.append_child("xdr:row").text().set(row + 10);
    to.append_child("xdr:rowOff").text().set(options.offsetY * 9525);

    XMLNode altContent = anchor.append_child("mc:AlternateContent");
    altContent.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");

    XMLNode choice = altContent.append_child("mc:Choice");
    choice.append_attribute("xmlns:sle15").set_value("http://schemas.microsoft.com/office/drawing/2012/slicer");
    choice.append_attribute("Requires").set_value("sle15");

    XMLNode graphicFrame = choice.append_child("xdr:graphicFrame");
    graphicFrame.append_attribute("macro").set_value("");

    XMLNode nvGraphicFramePr = graphicFrame.append_child("xdr:nvGraphicFramePr");
    XMLNode cNvPr            = nvGraphicFramePr.append_child("xdr:cNvPr");

    auto childCount = static_cast<size_t>(std::distance(drwRoot.children().begin(), drwRoot.children().end()));
    cNvPr.append_attribute("id").set_value(fmt::format("{}", childCount + 1).c_str());
    cNvPr.append_attribute("name").set_value(name.c_str());
    cNvPr.append_attribute("descr").set_value("");

    nvGraphicFramePr.append_child("xdr:cNvGraphicFramePr");

    XMLNode xfrm = graphicFrame.append_child("xdr:xfrm");
    xfrm.append_child("a:off").append_attribute("x").set_value("0");
    xfrm.child("a:off").append_attribute("y").set_value("0");
    xfrm.append_child("a:ext").append_attribute("cx").set_value("0");
    xfrm.child("a:ext").append_attribute("cy").set_value("0");

    XMLNode graphic     = graphicFrame.append_child("a:graphic");
    XMLNode graphicData = graphic.append_child("a:graphicData");
    graphicData.append_attribute("uri").set_value("http://schemas.microsoft.com/office/drawing/2010/slicer");

    XMLNode sleSlicer = graphicData.append_child("sle:slicer");
    sleSlicer.append_attribute("xmlns:sle").set_value("http://schemas.microsoft.com/office/drawing/2010/slicer");
    sleSlicer.append_attribute("name").set_value(name.c_str());

    XMLNode fallback = altContent.append_child("mc:Fallback");
    XMLNode sp       = fallback.append_child("xdr:sp");
    sp.append_attribute("macro").set_value("");
    sp.append_attribute("textlink").set_value("");

    XMLNode nvSpPr  = sp.append_child("xdr:nvSpPr");
    XMLNode spCNvPr = nvSpPr.append_child("xdr:cNvPr");
    spCNvPr.append_attribute("id").set_value(fmt::format("{}", childCount + 1).c_str());
    spCNvPr.append_attribute("name").set_value("");
    nvSpPr.append_child("xdr:cNvSpPr").append_attribute("txBox").set_value("true");

    XMLNode spPr   = sp.append_child("xdr:spPr");
    XMLNode spXfrm = spPr.append_child("a:xfrm");
    spXfrm.append_child("a:off").append_attribute("x").set_value("0");
    spXfrm.child("a:off").append_attribute("y").set_value("0");
    spXfrm.append_child("a:ext").append_attribute("cx").set_value(fmt::format("{}", static_cast<uint64_t>(options.width) * 9525).c_str());
    spXfrm.child("a:ext").append_attribute("cy").set_value(fmt::format("{}", static_cast<uint64_t>(options.height) * 9525).c_str());
    spPr.append_child("a:prstGeom").append_attribute("prst").set_value("rect");
    spPr.append_child("a:solidFill").append_child("a:srgbClr").append_attribute("val").set_value("FFFFFF");
    XMLNode ln = spPr.append_child("a:ln");
    ln.append_attribute("w").set_value("1");
    ln.append_child("a:solidFill").append_child("a:prstClr").append_attribute("val").set_value("black");

    XMLNode txBody = sp.append_child("xdr:txBody");
    XMLNode bodyPr = txBody.append_child("a:bodyPr");
    bodyPr.append_attribute("anchorCtr").set_value("false");
    bodyPr.append_attribute("rot").set_value("0");
    bodyPr.append_attribute("horzOverflow").set_value("clip");
    bodyPr.append_attribute("spcFirstLastPara").set_value("false");
    bodyPr.append_attribute("vertOverflow").set_value("clip");

    XMLNode p   = txBody.append_child("a:p");
    XMLNode r   = p.append_child("a:r");
    XMLNode rPr = r.append_child("a:rPr");
    rPr.append_attribute("b").set_value("false");
    rPr.append_attribute("i").set_value("false");
    r.append_child("a:t").text().set("This shape represents a table slicer. Table slicers are not supported in this version of Excel.");

    XMLNode cd = anchor.append_child("xdr:clientData");
    cd.append_attribute("fLocksWithSheet").set_value("true");
    cd.append_attribute("fPrintsWithSheet").set_value("true");
}

void XLWorksheet::addPivotSlicer(std::string_view       cellReference,
                                 const XLPivotTable&    pivotTable,
                                 std::string_view       columnName,
                                 const XLSlicerOptions& options)
{
    std::string ptName  = pivotTable.xmlDocument().document_element().attribute("name").value();
    uint32_t    cacheId = pivotTable.xmlDocument().document_element().attribute("cacheId").as_uint();

    // Find the sheetId where the pivot table lives (it should be this worksheet)
    uint32_t sheetId   = 0;
    XMLNode  wbkSheets = parentDoc().workbook().xmlDocument().document_element().child("sheets");
    for (auto sheet : wbkSheets.children("sheet")) {
        if (std::string(sheet.attribute("name").value()) == name()) {
            sheetId = sheet.attribute("sheetId").as_uint();
            break;
        }
    }

    std::string sName     = options.name.empty() ? std::string(columnName) : options.name;
    std::string caption   = options.caption.empty() ? std::string(columnName) : options.caption;
    std::string cacheName = "Slicer_" + sName;

    // Find PivotCacheDefinition and add the extLst required for slicers
    std::string pcRId;
    XMLNode     wbkPivotCaches = parentDoc().workbook().xmlDocument().document_element().child("pivotCaches");
    for (auto pc : wbkPivotCaches.children("pivotCache")) {
        if (pc.attribute("cacheId").as_uint() == cacheId) {
            pcRId = pc.attribute("r:id").value();
            break;
        }
    }
    if (!pcRId.empty()) {
        std::string pcTargetPath = parentDoc().workbookRelationships().relationshipById(pcRId).target();
        if (!pcTargetPath.empty() && pcTargetPath[0] != '/') pcTargetPath = "/xl/" + pcTargetPath;

        std::string targetPath = !pcTargetPath.empty() && pcTargetPath[0] == '/' ? pcTargetPath.substr(1) : pcTargetPath;
        
        // Use XLDocument's private getXmlData via our new friend status
        XLXmlData* data = parentDoc().getXmlData(targetPath);
        if (data) {
            XMLNode cacheRoot = data->getXmlDocument()->document_element();
            XMLNode extLst    = cacheRoot.child("extLst");
            if (extLst.empty()) extLst = cacheRoot.append_child("extLst");
            XMLNode ext = extLst.find_child_by_attribute("uri", "{725AE2AE-9491-48be-B2B4-4EB974FC3084}");
            if (ext.empty()) {
                ext = extLst.append_child("ext");
                ext.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
                ext.append_attribute("uri").set_value("{725AE2AE-9491-48be-B2B4-4EB974FC3084}");
                ext.append_child("x14:pivotCacheDefinition").append_attribute("pivotCacheId").set_value(cacheId);
            }
        }
    }

    // 1. Create Slicer Cache and Slicer files
    parentDoc().createPivotSlicerCache(cacheId, sheetId, ptName, cacheName, columnName);
    std::string slicerFilename = parentDoc().createSlicer(sName, cacheName, caption);

    // 2. Add Slicer relationship to worksheet
    relationships().addRelationship(XLRelationshipType::Slicer, "../slicers/" + slicerFilename.substr(11));
    std::string rId = relationships().relationshipByTarget("../slicers/" + slicerFilename.substr(11)).id();

    // 3. Add extLst to worksheet.xml
    XMLNode sheetNode = xmlDocument().document_element();
    XMLNode extLst    = sheetNode.child("extLst");
    if (extLst.empty()) extLst = sheetNode.append_child("extLst");

    XMLNode ext = extLst.find_child_by_attribute("uri", "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}");
    if (ext.empty()) {
        ext = extLst.append_child("ext");
        ext.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
        ext.append_attribute("uri").set_value("{A8765BA9-456A-4dab-B4F3-ACF838C121DE}");
    }

    XMLNode slicerList = ext.child("x14:slicerList");
    if (slicerList.empty()) { slicerList = ext.append_child("x14:slicerList"); }

    slicerList.append_child("x14:slicer").append_attribute("r:id").set_value(rId.c_str());

    // Ensure sheet namespaces for mc:Ignorable
    if (!sheetNode.attribute("xmlns:mc"))
        sheetNode.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
    if (!sheetNode.attribute("xmlns:x14"))
        sheetNode.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
    if (!sheetNode.attribute("mc:Ignorable")) { sheetNode.append_attribute("mc:Ignorable").set_value("x14"); }
    else {
        std::string ign = sheetNode.attribute("mc:Ignorable").value();
        if (ign.find("x14") == std::string::npos) { sheetNode.attribute("mc:Ignorable").set_value((ign + " x14").c_str()); }
    }

    // 4. Add drawing for the slicer
    XLCellReference ref{std::string(cellReference)};
    uint32_t        row    = ref.row() - 1;
    uint16_t        colIdx = ref.column() - 1;

    XLDrawing& drw     = drawing();
    XMLNode    drwRoot = drw.xmlDocument().document_element();

    XMLNode anchor = drwRoot.append_child("xdr:twoCellAnchor");

    XMLNode from = anchor.append_child("xdr:from");
    from.append_child("xdr:col").text().set(colIdx);
    from.append_child("xdr:colOff").text().set(options.offsetX * 9525);
    from.append_child("xdr:row").text().set(row);
    from.append_child("xdr:rowOff").text().set(options.offsetY * 9525);

    XMLNode to = anchor.append_child("xdr:to");
    to.append_child("xdr:col").text().set(colIdx + 2);
    to.append_child("xdr:colOff").text().set(options.offsetX * 9525);
    to.append_child("xdr:row").text().set(row + 10);
    to.append_child("xdr:rowOff").text().set(options.offsetY * 9525);

    XMLNode altContent = anchor.append_child("mc:AlternateContent");
    altContent.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");

    XMLNode choice = altContent.append_child("mc:Choice");
    choice.append_attribute("xmlns:a14").set_value("http://schemas.microsoft.com/office/drawing/2010/main");
    choice.append_attribute("Requires").set_value("a14");

    XMLNode graphicFrame = choice.append_child("xdr:graphicFrame");
    graphicFrame.append_attribute("macro").set_value("");

    XMLNode nvGraphicFramePr = graphicFrame.append_child("xdr:nvGraphicFramePr");
    XMLNode cNvPr            = nvGraphicFramePr.append_child("xdr:cNvPr");

    auto childCount = static_cast<size_t>(std::distance(drwRoot.children().begin(), drwRoot.children().end()));
    cNvPr.append_attribute("id").set_value(fmt::format("{}", childCount + 1).c_str());
    cNvPr.append_attribute("name").set_value(sName.c_str());
    cNvPr.append_attribute("descr").set_value("");

    nvGraphicFramePr.append_child("xdr:cNvGraphicFramePr");

    XMLNode xfrm = graphicFrame.append_child("xdr:xfrm");
    xfrm.append_child("a:off").append_attribute("x").set_value("0");
    xfrm.child("a:off").append_attribute("y").set_value("0");
    xfrm.append_child("a:ext").append_attribute("cx").set_value("0");
    xfrm.child("a:ext").append_attribute("cy").set_value("0");

    XMLNode graphic     = graphicFrame.append_child("a:graphic");
    XMLNode graphicData = graphic.append_child("a:graphicData");
    graphicData.append_attribute("uri").set_value("http://schemas.microsoft.com/office/drawing/2010/slicer");

    XMLNode sleSlicer = graphicData.append_child("sle:slicer");
    sleSlicer.append_attribute("xmlns:sle").set_value("http://schemas.microsoft.com/office/drawing/2010/slicer");
    sleSlicer.append_attribute("name").set_value(sName.c_str());

    XMLNode fallback = altContent.append_child("mc:Fallback");
    XMLNode sp       = fallback.append_child("xdr:sp");
    sp.append_attribute("macro").set_value("");
    sp.append_attribute("textlink").set_value("");

    XMLNode nvSpPr  = sp.append_child("xdr:nvSpPr");
    XMLNode spCNvPr = nvSpPr.append_child("xdr:cNvPr");
    spCNvPr.append_attribute("id").set_value(fmt::format("{}", childCount + 1).c_str());
    spCNvPr.append_attribute("name").set_value("");
    nvSpPr.append_child("xdr:cNvSpPr").append_attribute("txBox").set_value("true");

    XMLNode spPr   = sp.append_child("xdr:spPr");
    XMLNode spXfrm = spPr.append_child("a:xfrm");
    spXfrm.append_child("a:off").append_attribute("x").set_value("0");
    spXfrm.child("a:off").append_attribute("y").set_value("0");
    spXfrm.append_child("a:ext").append_attribute("cx").set_value(fmt::format("{}", static_cast<uint64_t>(options.width) * 9525).c_str());
    spXfrm.child("a:ext").append_attribute("cy").set_value(fmt::format("{}", static_cast<uint64_t>(options.height) * 9525).c_str());
    spPr.append_child("a:prstGeom").append_attribute("prst").set_value("rect");
    spPr.append_child("a:solidFill").append_child("a:srgbClr").append_attribute("val").set_value("FFFFFF");
    XMLNode ln = spPr.append_child("a:ln");
    ln.append_attribute("w").set_value("1");
    ln.append_child("a:solidFill").append_child("a:prstClr").append_attribute("val").set_value("black");

    XMLNode txBody = sp.append_child("xdr:txBody");
    XMLNode bodyPr = txBody.append_child("a:bodyPr");
    bodyPr.append_attribute("anchorCtr").set_value("false");
    bodyPr.append_attribute("rot").set_value("0");
    bodyPr.append_attribute("horzOverflow").set_value("clip");
    bodyPr.append_attribute("spcFirstLastPara").set_value("false");
    bodyPr.append_attribute("vertOverflow").set_value("clip");

    XMLNode p   = txBody.append_child("a:p");
    XMLNode r   = p.append_child("a:r");
    XMLNode rPr = r.append_child("a:rPr");
    rPr.append_attribute("b").set_value("false");
    rPr.append_attribute("i").set_value("false");
    r.append_child("a:t").text().set("This shape represents a pivot slicer. Slicers are not supported in this version of Excel.");

    XMLNode cd = anchor.append_child("xdr:clientData");
    cd.append_attribute("fLocksWithSheet").set_value("true");
    cd.append_attribute("fPrintsWithSheet").set_value("true");
}

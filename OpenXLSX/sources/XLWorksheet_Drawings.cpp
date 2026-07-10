#include <string>
#include <vector>

#include "XLDocument.hpp"
#include "XLUtilities.hpp"
#include "XLWorksheet.hpp"
#include <pugixml.hpp>

#include "XLChart.hpp"
#include "XLComments.hpp"
#include "XLConditionalFormatting.hpp"
#include "XLDataValidation.hpp"
#include "XLDrawing.hpp"
#include "XLImageOptions.hpp"
#include "XLMergeCells.hpp"
#include "XLPageSetup.hpp"
#include "XLPivotTable.hpp"
#include "XLRelationships.hpp"
#include "XLSparkline.hpp"
#include "XLStreamReader.hpp"
#include "XLStreamWriter.hpp"
#include "XLTables.hpp"
#include "XLThreadedComments.hpp"
#include "XLSlicer.hpp"
#include "XLSlicerCache.hpp"
#include "XLSlicerCollection.hpp"
#include "XLAutoFilter.hpp"
#include "XLWorksheetImpl.hpp"

using namespace OpenXLSX;

namespace
{
    uint64_t fnv1a_hash(gsl::span<const uint8_t> data)
    {
        uint64_t hash = 0xcbf29ce484222325ULL;
        for (uint8_t byte : data) {
            hash ^= byte;
            hash *= 0x100000001b3ULL;
        }
        return hash;
    }
}

bool XLWorksheet::hasRelationships() const { return parentDoc().hasSheetRelationships(sheetXmlNumber()); }
bool XLWorksheet::hasDrawing() const { return parentDoc().hasSheetDrawing(sheetXmlNumber()); }
bool XLWorksheet::hasVmlDrawing() const { return parentDoc().hasSheetVmlDrawing(sheetXmlNumber()); }
bool XLWorksheet::hasComments() const { return parentDoc().hasSheetComments(sheetXmlNumber()); }
bool XLWorksheet::hasThreadedComments() const { return parentDoc().hasSheetThreadedComments(sheetXmlNumber()); }
bool XLWorksheet::hasTables() const { return parentDoc().hasSheetTables(sheetXmlNumber()); }

XLRelationships& XLWorksheet::relationships()
{
    if (!m_impl->m_relationships.valid()) { m_impl->m_relationships = parentDoc().sheetRelationships(sheetXmlNumber()); }
    if (!m_impl->m_relationships.valid()) throw XLException("XLWorksheet::relationships(): could not create relationships XML");
    return m_impl->m_relationships;
}

XLDrawing& XLWorksheet::drawing()
{
    if (!m_impl->m_drawing.valid()) {
        XMLNode docElement = xmlDocument().document_element();
        std::ignore        = relationships();

        uint16_t sheetXmlNo = sheetXmlNumber();
        m_impl->m_drawing           = parentDoc().sheetDrawing(sheetXmlNo);
        if (!m_impl->m_drawing.valid()) throw XLException("XLWorksheet::drawing(): could not create drawing XML");

        std::string drawingPath         = m_impl->m_drawing.getXmlPath();
        std::string drawingRelativePath = getPathARelativeToPathB(drawingPath, getXmlPath());

        XLRelationshipItem drawingRelationship;
        if (!m_impl->m_relationships.targetExists(drawingRelativePath))
            drawingRelationship = m_impl->m_relationships.addRelationship(XLRelationshipType::Drawing, drawingRelativePath);
        else
            drawingRelationship = m_impl->m_relationships.relationshipByTarget(drawingRelativePath);

        if (drawingRelationship.empty()) throw XLException("XLWorksheet::drawing(): could not add sheet relationship for Drawing");

        if (docElement.child("drawing").empty()) {
            XMLNode drawingNode = ensureChild(docElement, "drawing", m_nodeOrder);
            setAttr(drawingNode, "r:id", drawingRelationship.id());
        }
    }

    return m_impl->m_drawing;
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

void XLWorksheet::addImage(const std::string&    name,
                           gsl::span<const uint8_t> data,
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
    if (!m_impl->m_drawing.valid()) {
        std::string drawingRelativePath = "";
        for (auto& rel : relationships().relationships()) {
            if (rel.type() == XLRelationshipType::Drawing) {
                drawingRelativePath = rel.target();
                break;
            }
        }
        if (drawingRelativePath.empty()) return result;
        m_impl->m_drawing = drawing();
    }

    uint32_t count = m_impl->m_drawing.imageCount();
    for (uint32_t i = 0; i < count; ++i) { result.push_back(m_impl->m_drawing.image(i)); }
    return result;
}

XLVmlDrawing& XLWorksheet::vmlDrawing()
{
    if (!m_impl->m_vmlDrawing.valid()) {
        XMLNode docElement = xmlDocument().document_element();
        std::ignore        = relationships();

        uint16_t sheetXmlNo = sheetXmlNumber();
        m_impl->m_vmlDrawing        = parentDoc().sheetVmlDrawing(sheetXmlNo);
        if (!m_impl->m_vmlDrawing.valid()) throw XLException("XLWorksheet::vmlDrawing(): could not create drawing XML");
        std::string        drawingRelativePath = getPathARelativeToPathB(m_impl->m_vmlDrawing.getXmlPath(), getXmlPath());
        XLRelationshipItem vmlDrawingRelationship;
        if (!m_impl->m_relationships.targetExists(drawingRelativePath))
            vmlDrawingRelationship = m_impl->m_relationships.addRelationship(XLRelationshipType::VMLDrawing, drawingRelativePath);
        else
            vmlDrawingRelationship = m_impl->m_relationships.relationshipByTarget(drawingRelativePath);
        if (vmlDrawingRelationship.empty())
            throw XLException("XLWorksheet::vmlDrawing(): could not add determine sheet relationship for VML Drawing");
        XMLNode legacyDrawing = ensureChild(docElement, "legacyDrawing", m_nodeOrder);
        if (legacyDrawing.empty()) throw XLException("XLWorksheet::vmlDrawing(): could not add <legacyDrawing> element to worksheet XML");
        setAttr(legacyDrawing, "r:id", vmlDrawingRelationship.id());
    }

    return m_impl->m_vmlDrawing;
}

XLComments& XLWorksheet::comments()
{
    if (!m_impl->m_comments.valid()) {
        std::ignore = relationships();
        std::ignore = vmlDrawing();

        uint16_t sheetXmlNo = sheetXmlNumber();
        m_impl->m_comments          = parentDoc().sheetComments(sheetXmlNo);
        if (!m_impl->m_comments.valid()) throw XLException("XLWorksheet::comments(): could not create comments XML");
        m_impl->m_comments.setVmlDrawing(m_impl->m_vmlDrawing);
        std::string commentsRelativePath = getPathARelativeToPathB(m_impl->m_comments.getXmlPath(), getXmlPath());
        if (!m_impl->m_relationships.targetExists(commentsRelativePath))
            m_impl->m_relationships.addRelationship(XLRelationshipType::Comments, commentsRelativePath);
    }

    return m_impl->m_comments;
}

XLThreadedComments& XLWorksheet::threadedComments()
{
    if (!m_impl->m_threadedComments.valid()) {
        std::ignore        = relationships();
        std::ignore        = comments(); // Trigger Comments & VML Drawing first to ensure stable relationship ordering (rId1=VML, rId2=Comments, rId3=Threaded)
        uint16_t sheetXmlNo = sheetXmlNumber();
        m_impl->m_threadedComments  = parentDoc().sheetThreadedComments(sheetXmlNo);
        if (!m_impl->m_threadedComments.valid()) throw XLException("XLWorksheet::threadedComments(): could not create threadedComments XML");

        std::string commentsPath         = m_impl->m_threadedComments.getXmlPath();
        std::string commentsRelativePath = getPathARelativeToPathB(commentsPath, getXmlPath());

        if (!m_impl->m_relationships.targetExists(commentsRelativePath))
            m_impl->m_relationships.addRelationship(XLRelationshipType::ThreadedComments, commentsRelativePath);
    }
    return m_impl->m_threadedComments;
}

XLTableCollection& XLWorksheet::tables()
{
    if (!m_impl->m_tables.valid()) { m_impl->m_tables = XLTableCollection(xmlDocument().document_element(), this); }

    return m_impl->m_tables;
}

void XLWorksheet::addHyperlink(std::string_view cellRef, std::string_view url, std::string_view tooltip)
{
    removeHyperlink(cellRef);
    std::ignore               = relationships();
    const auto rel            = m_impl->m_relationships.addRelationship(XLRelationshipType::Hyperlink, std::string(url), true);
    XMLNode    docElement     = xmlDocument().document_element();
    XMLNode    hyperlinksNode = docElement.child("hyperlinks");
    if (hyperlinksNode.empty()) { hyperlinksNode = ensureChild(docElement, "hyperlinks", m_nodeOrder); }
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
    if (hyperlinksNode.empty()) { hyperlinksNode = ensureChild(docElement, "hyperlinks", m_nodeOrder); }
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

    // Register revisions namespaces on worksheet root node
    XMLNode root = xmlDocument().document_element();
    ensureAttr(root, "xmlns:xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
    ensureAttr(root, "xmlns:xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
    ensureAttr(root, "xmlns:xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
    ensureAttr(root, "xr:uid", "{00000000-0001-0000-0000-000000000000}");
    auto worksheetIgnorableAttr = root.attribute("mc:Ignorable");
    if (!worksheetIgnorableAttr.empty()) {
        std::string ignorable = worksheetIgnorableAttr.value();
        if (ignorable.find("xr") == std::string::npos) {
            worksheetIgnorableAttr.set_value((ignorable + " xr xr2 xr3").c_str());
        }
    }
    else {
        setAttr(root, "mc:Ignorable", "x14ac xr xr2 xr3");
    }

    // Ensure workbook has co-authoring & revisions registered
    {
        XMLNode wbkRoot = parentDoc().workbook().xmlDocument().document_element();
        ensureAttr(wbkRoot, "xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        ensureAttr(wbkRoot, "xmlns:x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
        ensureAttr(wbkRoot, "xmlns:xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
        ensureAttr(wbkRoot, "xmlns:xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6");
        ensureAttr(wbkRoot, "xmlns:xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10");
        ensureAttr(wbkRoot, "xmlns:xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");

        auto wbkIgnorable = wbkRoot.attribute("mc:Ignorable");
        if (!wbkIgnorable.empty()) {
            std::string ignorable = wbkIgnorable.value();
            if (ignorable.find("xr") == std::string::npos) {
                wbkIgnorable.set_value((ignorable + " x15 xr xr6 xr10 xr2").c_str());
            }
        } else {
            setAttr(wbkRoot, "mc:Ignorable", "x15 xr xr6 xr10 xr2");
        }

        // Insert <xr:revisionPtr> in the correct order in workbook.xml
        if (wbkRoot.child("xr:revisionPtr").empty()) {
            XMLNode refNode = wbkRoot.child("bookViews");
            if (refNode.empty()) refNode = wbkRoot.child("sheets");

            XMLNode revPtr;
            if (!refNode.empty()) {
                revPtr = wbkRoot.insert_child_before("xr:revisionPtr", refNode);
            } else {
                revPtr = wbkRoot.append_child("xr:revisionPtr");
            }

            setAttr(revPtr, "revIDLastSave", "0");
            setAttr(revPtr, "documentId", "8_{9D879BB4-8C31-734F-9A02-992F49050029}");
            setAttr(revPtr, "xr6:coauthVersionLast", "47");
            setAttr(revPtr, "xr6:coauthVersionMax", "47");
            setAttr(revPtr, "xr10:uidLastSave", "{00000000-0000-0000-0000-000000000000}");
        }

        // Add xr2:uid to workbookView if exists
        XMLNode bookViews = wbkRoot.child("bookViews");
        if (!bookViews.empty()) {
            XMLNode workbookView = bookViews.child("workbookView");
            if (!workbookView.empty()) {
                ensureAttr(workbookView, "xr2:uid", "{0C52AA12-D58F-7D4C-86DF-F1B519ACB25B}");
            }
        }
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
        fallbackText += " ; Reply: " + text;

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

    // Generate internal name using FNV-1a hash to avoid collisions and enable deduplication
    uint64_t hashVal = fnv1a_hash(gsl::make_span(reinterpret_cast<const uint8_t*>(data.data()), data.size()));
    std::string name = "image_" + std::to_string(hashVal) + "." + size.extension;

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

void XLWorksheet::insertImage(const std::string& cellReference, gsl::span<const uint8_t> imageData, const XLImageOptions& options)
{
    XLImageSize size = XLImageParser::parseDimensions(imageData);
    if (!size.valid) throw XLInputError("Invalid or unsupported image file");

    // Generate internal name using FNV-1a hash to avoid collisions and enable deduplication
    uint64_t hashVal = fnv1a_hash(imageData);
    std::string name = "image_" + std::to_string(hashVal) + "." + size.extension;

    // Parse coordinate
    XLCellReference ref{std::string(cellReference)};
    uint32_t        row = ref.row() - 1;    // 0-indexed for addImage
    uint16_t        col = ref.column() - 1;

    // Scaled dimensions
    uint32_t scaledWidth  = static_cast<uint32_t>(size.width * options.scaleX);
    uint32_t scaledHeight = static_cast<uint32_t>(size.height * options.scaleY);

    // Call span-based addImage to avoid temporary string copies
    addImage(name, imageData, row, col, scaledWidth, scaledHeight, options);
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
    // Cache name is keyed on the source column, not on the user-provided slicer widget name.
    // This matches Excelize's setSlicerCache approach and prevents double-prefix
    // (e.g., when name="Slicer_Region", cacheName must still be "Slicer_Region" not "Slicer_Slicer_Region").
    std::string cacheName = "Slicer_" + std::string(columnName);

    // IMPORTANT: Force drawing() to be initialized first so it gets the lowest rId.
    // Excel always registers the drawing relationship before the slicer relationship.
    // This ensures drawing=rId1, table=rId2, slicer=rId3 (matching Excel's order).
    std::ignore = drawing();

    // 1. Create/Find Slicer Cache and Slicer files
    std::string actualCacheName = parentDoc().findOrCreateTableSlicerCache(tableId, colId, cacheName, columnName);

    // ── Per-sheet slicer XML sharing ─────────────────────────────────────
    // OOXML: all slicers on the same sheet share one slicer.xml file.
    // Look for an existing Slicer relationship on this worksheet.
    std::string existingSlicerFile;
    std::string existingSlicerRId;
    for (const auto& rel : relationships().relationships()) {
        if (rel.type() == XLRelationshipType::Slicer) {
            // Target is "../slicers/slicerN.xml" → resolve to "xl/slicers/slicerN.xml"
            existingSlicerFile = eliminateDotAndDotDotFromPath("xl/worksheets/" + rel.target());
            existingSlicerRId  = rel.id();
            break;
        }
    }

    // Append slicer node to existing file, or create new file
    std::string slicerFilename = parentDoc().createSlicer(name, actualCacheName, caption, existingSlicerFile);

    // Apply style/options to the newly appended slicer node (direct XML)
    if (!options.slicerStyle.empty()) {
        XLXmlData* slicerData = parentDoc().getXmlData(XLInternalAccess{}, slicerFilename);
        if (slicerData) {
            XMLNode root = slicerData->getXmlDocument()->document_element();
            for (auto node : root.children("slicer")) {
                if (std::string_view(node.attribute("name").value()) == name) {
                    // Set or update style attribute directly
                    auto styleAttr = node.attribute("style");
                    if (styleAttr) styleAttr.set_value(options.slicerStyle.c_str());
                    else node.append_attribute("style").set_value(options.slicerStyle.c_str());
                    break;
                }
            }
        }
    }

    if (!options.selectedItems.empty()) {
        auto tableFilter = table.autoFilter();
        if (tableFilter) {
            auto filterCol = tableFilter.filterColumn(colId - 1);
            filterCol.clearFilters();
            for (const auto& item : options.selectedItems) {
                filterCol.addFilter(item);
            }
        }
    }

    // 2. Add Slicer relationship to worksheet (only if this is the first slicer on this sheet)
    std::string slicerRelTarget = "../slicers/" + slicerFilename.substr(11);
    std::string rId;
    if (existingSlicerFile.empty()) {
        // First slicer on this sheet — register the relationship
        relationships().addRelationship(XLRelationshipType::Slicer, slicerRelTarget);
        rId = relationships().relationshipByTarget(slicerRelTarget).id();
    } else {
        // Reuse the existing relationship ID
        rId = existingSlicerRId;
    }

    // 3. Add extLst to worksheet.xml
    // Excel does NOT add xmlns:x14 to the worksheet root for slicers.
    // xmlns:x14 is declared inline on x14:slicerList only.
    // Do NOT modify mc:Ignorable on the worksheet root for slicers.
    XMLNode sheetNode = xmlDocument().document_element();
    XMLNode extLst    = sheetNode.child("extLst");
    if (!sheetNode.attribute("xmlns:mc"))
        sheetNode.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");

    if (extLst.empty()) extLst = ensureChild(sheetNode, "extLst", m_nodeOrder);

    // URI {3A4CF648-6AED-40f4-86FF-DC5316D8AED3}: table slicer list reference in worksheet extLst
    // Excel puts uri first, then xmlns:x15 on ext; x14:slicerList must declare xmlns:x14 inline
    XMLNode ext = extLst.find_child_by_attribute("uri", "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}");
    if (ext.empty()) {
        ext = extLst.append_child("ext");
        ext.append_attribute("uri").set_value("{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}");
        ext.append_attribute("xmlns:x15").set_value("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
    }

    XMLNode slicerList = ext.child("x14:slicerList");
    if (slicerList.empty()) {
        slicerList = ext.append_child("x14:slicerList");
        // xmlns:x14 must be declared inline on x14:slicerList (not on worksheet root)
        slicerList.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
    }

    // Only add the relationship reference once (first slicer registers it)
    if (existingSlicerFile.empty()) {
        slicerList.append_child("x14:slicer").append_attribute("r:id").set_value(rId.c_str());
    }

    // Build drawing anchor using oneCellAnchor + ext for pixel-accurate sizing
    // Bug fix: old code used twoCellAnchor with hardcoded +2 col / +10 row offsets
    // which made options.width / options.height completely ineffective.
    XLCellReference ref{std::string(cellReference)};
    uint32_t        row    = ref.row() - 1;    // 0-indexed
    uint16_t        colIdx = ref.column() - 1;

    XLDrawing& drw     = drawing();
    XMLNode    drwRoot = drw.xmlDocument().document_element();

    // Excel uses twoCellAnchor with editAs="absolute" for slicers.
    // Size is encoded in the "to" corner cell/offset coords.
    // 1 EMU = 1/914400 inch; default col width ~9 chars ≈ 914400 EMU, row height ≈ 190500 EMU
    // We approximate: each col ≈ 914400 EMU wide, each row ≈ 190500 EMU tall
    // cols spanned = ceil(width_px * 9525 / 914400), rows spanned = ceil(height_px * 9525 / 190500)
    const int64_t cx = static_cast<int64_t>(options.width)  * 9525;
    const int64_t cy = static_cast<int64_t>(options.height) * 9525;

    // Compute "to" cell: find the cell that the bottom-right corner lands in
    const int64_t colWidthEmu = 914400;   // ~1 inch per column (approx)
    const int64_t rowHeightEmu = 190500;  // default row height in EMU
    int64_t toCol    = colIdx + cx / colWidthEmu;
    int64_t toColOff = cx % colWidthEmu;
    int64_t toRow    = row + cy / rowHeightEmu;
    int64_t toRowOff = cy % rowHeightEmu;

    // Count existing anchors in THIS drawing to generate a drawing-local unique shapeId.
    // OOXML requires cNvPr id to be unique within the drawing (not workbook-wide).
    uint32_t shapeId = 1;
    for (auto child = drwRoot.first_child_of_type(pugi::node_element); !child.empty();
         child = child.next_sibling_of_type(pugi::node_element)) {
        ++shapeId;
    }

    XMLNode anchor = drwRoot.append_child("xdr:twoCellAnchor");
    anchor.append_attribute("editAs").set_value("absolute");

    XMLNode from = anchor.append_child("xdr:from");
    from.append_child("xdr:col").text().set(colIdx);
    from.append_child("xdr:colOff").text().set(static_cast<int64_t>(options.offsetX) * 9525);
    from.append_child("xdr:row").text().set(row);
    from.append_child("xdr:rowOff").text().set(static_cast<int64_t>(options.offsetY) * 9525);

    XMLNode to = anchor.append_child("xdr:to");
    to.append_child("xdr:col").text().set(toCol);
    to.append_child("xdr:colOff").text().set(toColOff);
    to.append_child("xdr:row").text().set(toRow);
    to.append_child("xdr:rowOff").text().set(toRowOff);

    // OOXML ECMA-376 §22.1.2: namespace declarations scoped to the minimum required element.
    // Excel writes xmlns:sle15 on mc:Choice and xmlns:sle directly on sle:slicer (not on AlternateContent).
    XMLNode altContent = anchor.append_child("mc:AlternateContent");
    altContent.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");

    XMLNode choice = altContent.append_child("mc:Choice");
    choice.append_attribute("xmlns:sle15").set_value("http://schemas.microsoft.com/office/drawing/2012/slicer");
    choice.append_attribute("Requires").set_value("sle15");

    XMLNode graphicFrame = choice.append_child("xdr:graphicFrame");
    graphicFrame.append_attribute("macro").set_value("");

    XMLNode nvGraphicFramePr = graphicFrame.append_child("xdr:nvGraphicFramePr");
    XMLNode cNvPr            = nvGraphicFramePr.append_child("xdr:cNvPr");

    cNvPr.append_attribute("id").set_value(fmt::format("{}", shapeId).c_str());
    cNvPr.append_attribute("name").set_value(name.c_str());

    // ECMA-376 / Office Open XML: <a:extLst> with a16:creationId is required for slicer shape identity
    // (Excel 2013+ uses this to link drawing shape → slicer object)
    {
        XMLNode extLst    = cNvPr.append_child("a:extLst");
        XMLNode ext       = extLst.append_child("a:ext");
        ext.append_attribute("uri").set_value("{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}");
        XMLNode creationId = ext.append_child("a16:creationId");
        creationId.append_attribute("xmlns:a16").set_value("http://schemas.microsoft.com/office/drawing/2014/main");
        // Generate a stable UID based on shape index: {00000000-0008-0000-0000-00000N000000}
        creationId.append_attribute("id").set_value(
            fmt::format("{{00000000-0008-0000-0000-{:012X}}}", shapeId).c_str());
    }

    nvGraphicFramePr.append_child("xdr:cNvGraphicFramePr");

    XMLNode xfrm = graphicFrame.append_child("xdr:xfrm");
    xfrm.append_child("a:off").append_attribute("x").set_value("0");
    xfrm.child("a:off").append_attribute("y").set_value("0");
    xfrm.append_child("a:ext").append_attribute("cx").set_value("0");
    xfrm.child("a:ext").append_attribute("cy").set_value("0");

    XMLNode graphic     = graphicFrame.append_child("a:graphic");
    XMLNode graphicData = graphic.append_child("a:graphicData");
    // Excel uses 2010/slicer URI for table slicers in the graphicData element
    graphicData.append_attribute("uri").set_value("http://schemas.microsoft.com/office/drawing/2010/slicer");

    // xmlns:sle is declared on sle:slicer itself (not hoisted to AlternateContent)
    XMLNode sleSlicer = graphicData.append_child("sle:slicer");
    sleSlicer.append_attribute("xmlns:sle").set_value("http://schemas.microsoft.com/office/drawing/2010/slicer");
    sleSlicer.append_attribute("name").set_value(name.c_str());

    // Fallback: no xmlns="" — Table slicer fallback does not reset the default namespace.
    XMLNode fallback = altContent.append_child("mc:Fallback");
    XMLNode sp       = fallback.append_child("xdr:sp");
    sp.append_attribute("macro").set_value("");
    sp.append_attribute("textlink").set_value("");

    XMLNode nvSpPr  = sp.append_child("xdr:nvSpPr");
    XMLNode spCNvPr = nvSpPr.append_child("xdr:cNvPr");
    spCNvPr.append_attribute("id").set_value(fmt::format("{}", shapeId).c_str());
    spCNvPr.append_attribute("name").set_value("");
    // Excel adds <a:spLocks noTextEdit="1"/> to mark the fallback shape as non-editable
    XMLNode cNvSpPr = nvSpPr.append_child("xdr:cNvSpPr");
    cNvSpPr.append_child("a:spLocks").append_attribute("noTextEdit").set_value("1");

    XMLNode spPr   = sp.append_child("xdr:spPr");
    XMLNode spXfrm = spPr.append_child("a:xfrm");
    // Fallback uses absolute EMU position computed from the anchor cell + offset
    int64_t fallbackX = static_cast<int64_t>(colIdx) * 914400 +
                        static_cast<int64_t>(options.offsetX) * 9525;
    int64_t fallbackY = static_cast<int64_t>(row) * 190500 +
                        static_cast<int64_t>(options.offsetY) * 9525;
    spXfrm.append_child("a:off").append_attribute("x").set_value(std::to_string(fallbackX).c_str());
    spXfrm.child("a:off").append_attribute("y").set_value(std::to_string(fallbackY).c_str());
    spXfrm.append_child("a:ext").append_attribute("cx").set_value(std::to_string(cx).c_str());
    spXfrm.child("a:ext").append_attribute("cy").set_value(std::to_string(cy).c_str());
    XMLNode prstGeom = spPr.append_child("a:prstGeom");
    prstGeom.append_attribute("prst").set_value("rect");
    prstGeom.append_child("a:avLst");
    spPr.append_child("a:solidFill").append_child("a:prstClr").append_attribute("val").set_value("white");
    XMLNode ln = spPr.append_child("a:ln");
    ln.append_attribute("w").set_value("1");
    ln.append_child("a:solidFill").append_child("a:prstClr").append_attribute("val").set_value("green");

    XMLNode txBody = sp.append_child("xdr:txBody");
    XMLNode bodyPr = txBody.append_child("a:bodyPr");
    bodyPr.append_attribute("vertOverflow").set_value("clip");
    bodyPr.append_attribute("horzOverflow").set_value("clip");
    txBody.append_child("a:lstStyle");
    XMLNode p   = txBody.append_child("a:p");
    XMLNode r   = p.append_child("a:r");
    r.append_child("a:rPr").append_attribute("lang").set_value("en-US");
    r.append_child("a:t").text().set("This shape represents a table slicer. Table slicers are not supported in this version of Excel.\n\nIf the shape was modified in an earlier version of Excel, or if the workbook was saved in Excel 2007 or earlier, the slicer cannot be used.");

    // ECMA-376 §20.5.2.3 CT_TwoCellAnchor requires <xdr:clientData/> as the last child
    anchor.append_child("xdr:clientData");

    // Register definedName pointing to #N/A (required by Excel for slicers to not trigger corruption warning)
    if (!parentDoc().workbook().definedNames().exists(name)) {
        parentDoc().workbook().definedNames().append(name, "#N/A");
    }
    if (actualCacheName != name && !parentDoc().workbook().definedNames().exists(actualCacheName)) {
        parentDoc().workbook().definedNames().append(actualCacheName, "#N/A");
    }
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
    // Cache name is keyed on source column, not user-provided slicer widget name.
    std::string cacheName = "Slicer_" + std::string(columnName);

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
        XLXmlData* data = parentDoc().getXmlData(XLInternalAccess{}, targetPath);
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
    std::string actualCacheName = parentDoc().findOrCreatePivotSlicerCache(cacheId, sheetId, ptName, cacheName, columnName);

    // Per-sheet slicer XML sharing: look for existing slicer relationship on this worksheet
    std::string existingSlicerFile;
    std::string existingSlicerRId;
    for (const auto& rel : relationships().relationships()) {
        if (rel.type() == XLRelationshipType::Slicer) {
            existingSlicerFile = eliminateDotAndDotDotFromPath("xl/worksheets/" + rel.target());
            existingSlicerRId  = rel.id();
            break;
        }
    }
    std::string slicerFilename = parentDoc().createSlicer(sName, actualCacheName, caption, existingSlicerFile);

    // Apply style via direct XML attribute (no XLSlicer construction needed)
    if (!options.slicerStyle.empty()) {
        XLXmlData* slicerData = parentDoc().getXmlData(XLInternalAccess{}, slicerFilename);
        if (slicerData) {
            XMLNode root = slicerData->getXmlDocument()->document_element();
            for (auto node : root.children("slicer")) {
                if (std::string_view(node.attribute("name").value()) == sName) {
                    auto styleAttr = node.attribute("style");
                    if (styleAttr) styleAttr.set_value(options.slicerStyle.c_str());
                    else node.append_attribute("style").set_value(options.slicerStyle.c_str());
                    break;
                }
            }
        }
    }

    for (const auto& item : parentDoc().archive().entryNames()) {
        if (item.find("xl/slicerCaches/slicerCache") != std::string::npos) {
            XLXmlData* cacheData = parentDoc().getXmlData(XLInternalAccess{}, item);
            if (cacheData) {
                XLSlicerCache slicerCache(cacheData);
                if (slicerCache.name() == actualCacheName) {
                    slicerCache.syncWithPivotCache(pivotTable.cacheDefinition(), options.selectedItems);
                    break;
                }
            }
        }
    }

    // 2. Add Slicer relationship to worksheet (only for first slicer on this sheet)
    std::string slicerRelTarget = "../slicers/" + slicerFilename.substr(11);
    std::string rId;
    if (existingSlicerFile.empty()) {
        relationships().addRelationship(XLRelationshipType::Slicer, slicerRelTarget);
        rId = relationships().relationshipByTarget(slicerRelTarget).id();
    } else {
        rId = existingSlicerRId;
    }

    // 3. Add extLst to worksheet.xml
    XMLNode sheetNode = xmlDocument().document_element();
    XMLNode extLst    = sheetNode.child("extLst");
    if (extLst.empty()) extLst = ensureChild(sheetNode, "extLst", m_nodeOrder);

    XMLNode ext = extLst.find_child_by_attribute("uri", "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}");
    if (ext.empty()) {
        ext = extLst.append_child("ext");
        ext.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
        ext.append_attribute("uri").set_value("{A8765BA9-456A-4dab-B4F3-ACF838C121DE}");
    }

    XMLNode slicerList = ext.child("x14:slicerList");
    if (slicerList.empty()) { slicerList = ext.append_child("x14:slicerList"); }

    slicerList.append_child("x14:slicer").append_attribute("r:id").set_value(rId.c_str());

    // 4. Add drawing for the slicer (twoCellAnchor + absolute for Excel compliance)
    XLCellReference ref{std::string(cellReference)};
    uint32_t        row    = ref.row() - 1;
    uint16_t        colIdx = ref.column() - 1;

    XLDrawing& drw     = drawing();
    XMLNode    drwRoot = drw.xmlDocument().document_element();

    const int64_t cx = static_cast<int64_t>(options.width)  * 9525;
    const int64_t cy = static_cast<int64_t>(options.height) * 9525;

    // Compute "to" cell: find the cell that the bottom-right corner lands in
    const int64_t colWidthEmu = 914400;   // ~1 inch per column (approx)
    const int64_t rowHeightEmu = 190500;  // default row height in EMU
    int64_t toCol    = colIdx + cx / colWidthEmu;
    int64_t toColOff = cx % colWidthEmu;
    int64_t toRow    = row + cy / rowHeightEmu;
    int64_t toRowOff = cy % rowHeightEmu;

    // Count existing anchors in THIS drawing to generate a drawing-local unique shapeId.
    // OOXML requires cNvPr id to be unique within the drawing (not workbook-wide).
    uint32_t shapeId = 1;
    for (auto child = drwRoot.first_child_of_type(pugi::node_element); !child.empty();
         child = child.next_sibling_of_type(pugi::node_element)) {
        ++shapeId;
    }

    XMLNode anchor = drwRoot.append_child("xdr:twoCellAnchor");
    anchor.append_attribute("editAs").set_value("absolute");

    XMLNode from = anchor.append_child("xdr:from");
    from.append_child("xdr:col").text().set(colIdx);
    from.append_child("xdr:colOff").text().set(static_cast<int64_t>(options.offsetX) * 9525);
    from.append_child("xdr:row").text().set(row);
    from.append_child("xdr:rowOff").text().set(static_cast<int64_t>(options.offsetY) * 9525);

    XMLNode to = anchor.append_child("xdr:to");
    to.append_child("xdr:col").text().set(toCol);
    to.append_child("xdr:colOff").text().set(toColOff);
    to.append_child("xdr:row").text().set(toRow);
    to.append_child("xdr:rowOff").text().set(toRowOff);

    // OOXML ECMA-376 §22.1.2: xmlns:sle MUST be declared on mc:AlternateContent (or ancestor)
    XMLNode altContent = anchor.append_child("mc:AlternateContent");
    altContent.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
    altContent.append_attribute("xmlns:sle").set_value("http://schemas.microsoft.com/office/drawing/2010/slicer");

    XMLNode choice = altContent.append_child("mc:Choice");
    choice.append_attribute("Requires").set_value("sle");

    XMLNode graphicFrame = choice.append_child("xdr:graphicFrame");
    graphicFrame.append_attribute("macro").set_value("");

    XMLNode nvGraphicFramePr = graphicFrame.append_child("xdr:nvGraphicFramePr");
    XMLNode cNvPr            = nvGraphicFramePr.append_child("xdr:cNvPr");

    cNvPr.append_attribute("id").set_value(fmt::format("{}", shapeId).c_str());
    cNvPr.append_attribute("name").set_value(sName.c_str());

    // ECMA-376 / Office Open XML: <a:extLst> with a16:creationId is required for slicer shape identity
    {
        XMLNode extLst    = cNvPr.append_child("a:extLst");
        XMLNode ext       = extLst.append_child("a:ext");
        ext.append_attribute("uri").set_value("{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}");
        XMLNode creationId = ext.append_child("a16:creationId");
        creationId.append_attribute("xmlns:a16").set_value("http://schemas.microsoft.com/office/drawing/2014/main");
        creationId.append_attribute("id").set_value(
            fmt::format("{{00000000-0008-0000-0000-{:012X}}}", shapeId).c_str());
    }

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
    sleSlicer.append_attribute("name").set_value(sName.c_str());

    XMLNode fallback = altContent.append_child("mc:Fallback");
    fallback.append_attribute("xmlns").set_value("");
    XMLNode sp       = fallback.append_child("xdr:sp");
    sp.append_attribute("macro").set_value("");
    sp.append_attribute("textlink").set_value("");

    XMLNode nvSpPr  = sp.append_child("xdr:nvSpPr");
    XMLNode spCNvPr = nvSpPr.append_child("xdr:cNvPr");
    spCNvPr.append_attribute("id").set_value(fmt::format("{}", shapeId).c_str());
    spCNvPr.append_attribute("name").set_value("");
    nvSpPr.append_child("xdr:cNvSpPr").append_attribute("txBox").set_value("true");

    XMLNode spPr   = sp.append_child("xdr:spPr");
    XMLNode spXfrm = spPr.append_child("a:xfrm");
    spXfrm.append_child("a:off").append_attribute("x").set_value("0");
    spXfrm.child("a:off").append_attribute("y").set_value("0");
    spXfrm.append_child("a:ext").append_attribute("cx").set_value(std::to_string(cx).c_str());
    spXfrm.child("a:ext").append_attribute("cy").set_value(std::to_string(cy).c_str());
    XMLNode prstGeom = spPr.append_child("a:prstGeom");
    prstGeom.append_attribute("prst").set_value("rect");
    prstGeom.append_child("a:avLst");
    spPr.append_child("a:solidFill").append_child("a:srgbClr").append_attribute("val").set_value("FFFFFF");
    XMLNode ln = spPr.append_child("a:ln");
    ln.append_attribute("w").set_value("1");
    ln.append_child("a:solidFill").append_child("a:prstClr").append_attribute("val").set_value("green");

    XMLNode txBody = sp.append_child("xdr:txBody");
    XMLNode bodyPr = txBody.append_child("a:bodyPr");
    bodyPr.append_attribute("vertOverflow").set_value("clip");
    bodyPr.append_attribute("horzOverflow").set_value("clip");
    txBody.append_child("a:lstStyle");
    XMLNode p   = txBody.append_child("a:p");
    XMLNode r   = p.append_child("a:r");
    r.append_child("a:rPr").append_attribute("lang").set_value("en-US");
    r.append_child("a:t").text().set("This shape represents a pivot slicer. Slicers are not supported in this version of Excel.\n\nIf the shape was modified in an earlier version of Excel, or if the workbook was saved in Excel 2007 or earlier, the slicer cannot be used.");

    anchor.append_child("xdr:clientData");

    // Register definedName pointing to #N/A (required by Excel for slicers to not trigger corruption warning)
    if (!parentDoc().workbook().definedNames().exists(sName)) {
        parentDoc().workbook().definedNames().append(sName, "#N/A");
    }
    if (actualCacheName != sName && !parentDoc().workbook().definedNames().exists(actualCacheName)) {
        parentDoc().workbook().definedNames().append(actualCacheName, "#N/A");
    }
}

XLSlicerCollection& XLWorksheet::slicers()
{
    if (!m_impl->m_slicers.valid()) {
        m_impl->m_slicers = XLSlicerCollection(this);
    }
    return m_impl->m_slicers;
}

void XLWorksheet::deleteSlicer(const std::string& name)
{
    // 1. Remove drawing shape representation (search inside mc:AlternateContent)
    if (hasDrawing()) {
        XLDrawing& drw = drawing();
        auto drwRoot = drw.xmlDocument().document_element();
        pugi::xml_node targetAnchor;
        for (auto anchor : drwRoot.children()) {
            // New anchor style: oneCellAnchor/twoCellAnchor with mc:AlternateContent
            auto alt = anchor.child("mc:AlternateContent");
            if (alt) {
                // Look inside mc:Choice → xdr:graphicFrame → ... → sle:slicer[name]
                for (auto choice : alt.children("mc:Choice")) {
                    auto gf = choice.child("xdr:graphicFrame");
                    if (!gf) continue;
                    auto cNvPr = gf.child("xdr:nvGraphicFramePr").child("xdr:cNvPr");
                    if (cNvPr && std::string(cNvPr.attribute("name").value()) == name) {
                        targetAnchor = anchor;
                        break;
                    }
                    // Also check sle:slicer name directly
                    auto sle = gf.child("a:graphic").child("a:graphicData").child("sle:slicer");
                    if (!sle)
                        sle = gf.child("a:graphic").child("a:graphicData").child("sle15:slicer");
                    if (sle && std::string(sle.attribute("name").value()) == name) {
                        targetAnchor = anchor;
                        break;
                    }
                }
                if (targetAnchor) break;
            }
            // Fallback: direct xdr:graphicFrame (no AlternateContent)
            auto graphicFrame = anchor.child("xdr:graphicFrame");
            if (graphicFrame) {
                auto nameAttr = graphicFrame.child("xdr:nvGraphicFramePr").child("xdr:cNvPr").attribute("name");
                if (nameAttr && std::string(nameAttr.value()) == name) {
                    targetAnchor = anchor;
                    break;
                }
            }
        }
        if (targetAnchor) {
            drwRoot.remove_child(targetAnchor);
        }
    }

    // 2. Remove relationship and sheet extLst reference only if no other slicers share the file
    std::string rId;
    bool otherSlicersLeft = false;
    auto& rels = relationships();
    auto relList = rels.relationships();
    for (const auto& rel : relList) {
        if (rel.type() == XLRelationshipType::Slicer) {
            std::string target = rel.target();
            std::string absPath = eliminateDotAndDotDotFromPath("xl/worksheets/" + target);
            XLXmlData* xmlData = parentDoc().getXmlData(XLInternalAccess{}, absPath, true);
            if (!xmlData) {
                xmlData = const_cast<XLDocument&>(parentDoc()).addXmlData(XLInternalAccess{}, absPath, "", XLContentType::Slicer);
            }
            if (xmlData) {
                auto root = xmlData->getXmlDocument()->document_element();
                bool foundThisSlicer = false;
                for (auto slicerNode : root.children("slicer")) {
                    if (std::string(slicerNode.attribute("name").value()) == name) {
                        foundThisSlicer = true;
                    } else {
                        otherSlicersLeft = true;
                    }
                }
                if (foundThisSlicer) {
                    rId = rel.id();
                    if (!otherSlicersLeft) {
                        rels.deleteRelationship(rel);
                    }
                    break;
                }
            }
        }
    }

    if (!rId.empty() && !otherSlicersLeft) {
        auto sheetNode = xmlDocument().document_element();
        auto extLst = sheetNode.child("extLst");
        if (extLst) {
            for (auto ext : extLst.children("ext")) {
                auto slicerList = ext.child("x14:slicerList");
                if (slicerList) {
                    auto slicerNode = slicerList.find_child_by_attribute("r:id", rId.c_str());
                    if (slicerNode) {
                        slicerList.remove_child(slicerNode);
                    }
                    if (slicerList.first_child().empty()) {
                        ext.remove_child(slicerList);
                    }
                }
            }
        }
    }

    // 3. Delete slicer file and orphan cache
    parentDoc().deleteSlicerFileAndOrphanCache(name);

    // Remove the definedName associated with the slicer
    if (parentDoc().workbook().definedNames().exists(name)) {
        parentDoc().workbook().definedNames().remove(name);
    }

    // 4. Invalidate the collection cache so next access re-enumerates
    m_impl->m_slicers.reload();
}


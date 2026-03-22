// ===== External Includes ===== //
// #include <algorithm>
#include <gsl/gsl>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLComments.hpp"
#include "XLDocument.hpp"     // pugi_parse_settings
#include "XLUtilities.hpp"    // OpenXLSX::ignore, appendAndGetNode

using namespace OpenXLSX;

namespace
{
    /**
     * @details Helper function to parse rich text from a comment node.
     */
    XLRichText parseRichText(XMLNode commentNode)
    {
        XLRichText result;
        XMLNode    textNode = commentNode.child("text");
        if (textNode.empty()) return result;

        XMLNode element = textNode.first_child_of_type(pugi::node_element);
        while (!element.empty()) {
            if (std::string(element.name()) == "t") { result.addRun(XLRichTextRun(element.text().get())); }
            else if (std::string(element.name()) == "r") {
                XLRichTextRun run;
                XMLNode       tNode = element.child("t");
                if (!tNode.empty()) run.setText(tNode.text().get());

                XMLNode rPrNode = element.child("rPr");
                if (!rPrNode.empty()) {
                    if (!rPrNode.child("b").empty()) run.setBold(true);
                    if (!rPrNode.child("i").empty()) run.setItalic(true);
                    if (!rPrNode.child("u").empty()) run.setUnderline(true);
                    if (!rPrNode.child("strike").empty()) run.setStrikethrough(true);

                    XMLNode rFontNode = rPrNode.child("rFont");
                    if (!rFontNode.empty()) run.setFontName(rFontNode.attribute("val").value());

                    XMLNode szNode = rPrNode.child("sz");
                    if (!szNode.empty()) run.setFontSize(szNode.attribute("val").as_uint());

                    XMLNode colorNode = rPrNode.child("color");
                    if (!colorNode.empty()) {
                        if (colorNode.attribute("rgb")) run.setFontColor(XLColor(colorNode.attribute("rgb").value()));
                    }
                }
                result.addRun(run);
            }
            element = element.next_sibling_of_type(pugi::node_element);
        }
        return result;
    }

    std::string getCommentString(XMLNode commentNode)
    {
        std::string result{};

        using namespace std::literals::string_literals;
        XMLNode textElement = commentNode.child("text").first_child_of_type(pugi::node_element);
        while (not textElement.empty()) {
            if (textElement.name() == "t"s) { result += textElement.first_child().value(); }
            else if (textElement.name() == "r"s) {    // rich text
                XMLNode richTextSubnode = textElement.first_child_of_type(pugi::node_element);
                while (not richTextSubnode.empty()) {
                    if (richTextSubnode.name() == "t"s) { result += richTextSubnode.first_child().value(); }
                    else if (richTextSubnode.name() == "rPr"s) {}    // ignore rich text formatting info
                    else {}                                          // ignore other nodes
                    richTextSubnode = richTextSubnode.next_sibling_of_type(pugi::node_element);
                }
            }
            else {}    // ignore other elements (for now)
            textElement = textElement.next_sibling_of_type(pugi::node_element);
        }
        return result;
    }
}    // namespace
// XLComment::XLComment()
//  : m_commentNode(std::make_unique<XMLNode>())
//  {}
XLComment::XLComment(XMLNode node) : m_commentNode(node) {}

/**
 * @details
 * @note Function body moved to cpp module as it uses "not" keyword for readability, which MSVC sabotages with non-CPP compatibility.
 * @note For the library it is reasonable to expect users to compile it with MSCV /permissive- flag, but for the user's own projects the
 * header files shall "just work"
 */
bool XLComment::valid() const { return not m_commentNode.empty(); }
std::string XLComment::ref() const { return m_commentNode.attribute("ref").value(); }
std::string XLComment::text() const { return getCommentString(m_commentNode); }
XLRichText  XLComment::richText() const { return parseRichText(m_commentNode); }
uint16_t    XLComment::authorId() const { return gsl::narrow_cast<uint16_t>(m_commentNode.attribute("authorId").as_uint()); }
XLComment& XLComment::setText(const std::string& newText)
{
    m_commentNode.remove_child("text");
    XMLNode tNode = m_commentNode.append_child("text").append_child("t");
    tNode.append_attribute("xml:space").set_value("preserve");
    tNode.text().set(newText.c_str());
    return *this;
}

XLComment& XLComment::setRichText(const XLRichText& richText)
{
    m_commentNode.remove_child("text");
    XMLNode textNode = m_commentNode.append_child("text");

    for (const auto& run : richText.runs()) {
        XMLNode rNode = textNode.append_child("r");

        // Add properties if any are set
        if (run.fontName() || run.fontSize() || run.fontColor() || run.bold() || run.italic() || run.underline() || run.strikethrough()) {
            XMLNode rPrNode = rNode.append_child("rPr");
            if (run.fontName()) {
                XMLNode rFontNode = rPrNode.append_child("rFont");
                rFontNode.append_attribute("val").set_value(run.fontName()->c_str());
            }
            if (run.fontSize()) {
                XMLNode szNode = rPrNode.append_child("sz");
                szNode.append_attribute("val").set_value(*run.fontSize());
            }
            if (run.fontColor()) {
                XMLNode colorNode = rPrNode.append_child("color");
                colorNode.append_attribute("rgb").set_value(run.fontColor()->hex().c_str());
            }
            if (run.bold() && *run.bold()) { rPrNode.append_child("b"); }
            if (run.italic() && *run.italic()) { rPrNode.append_child("i"); }
            if (run.underline() && *run.underline()) { rPrNode.append_child("u"); }
            if (run.strikethrough() && *run.strikethrough()) { rPrNode.append_child("strike"); }
        }

        // Add text
        XMLNode tNode = rNode.append_child("t");
        tNode.text().set(run.text().c_str());

        // Handle space preservation
        if (!run.text().empty() && (run.text().front() == ' ' || run.text().back() == ' ')) {
            tNode.append_attribute("xml:space").set_value("preserve");
        }
    }
    return *this;
}
XLComment& XLComment::setAuthorId(uint16_t newAuthorId)
{ appendAndSetAttribute(m_commentNode, "authorId", std::to_string(newAuthorId)); return *this; }


XLComments::XLComments() : XLXmlFile(nullptr), m_vmlDrawing() {}

/**
 * @details The constructor creates an instance of the superclass, XLXmlFile
 */
XLComments::XLComments(XLXmlData* xmlData) : XLXmlFile(xmlData), m_vmlDrawing()
{
    if (xmlData->getXmlType() != XLContentType::Comments) throw XLInternalError("XLComments constructor: Invalid XML data.");

    XMLDocument& doc = xmlDocument();
    if (doc.document_element().empty())    // handle a bad (no document element) comments XML file
        doc.load_string("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
                        "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                        " xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n"
                        "</comments>",
                        pugi_parse_settings);

    XMLNode rootNode = doc.document_element();
    bool    docNew   = rootNode.first_child_of_type(pugi::node_element).empty();    // check and store status: was document empty?
    m_authors        = appendAndGetNode(rootNode, "authors", m_nodeOrder);          // (insert and) get authors node
    if (docNew)
        rootNode.prepend_child(pugi::node_pcdata)
            .set_value("\n\t");    // if document was empty: prefix the newly inserted authors node with a single tab
    m_commentList =
        appendAndGetNode(rootNode,
                         "commentList",
                         m_nodeOrder);    // (insert and) get commentList node -> this should now copy the whitespace prefix of m_authors

    // whitespace formatting / indentation for closing tags:
    if (m_authors.first_child().empty()) m_authors.append_child(pugi::node_pcdata).set_value("\n\t");
    if (m_commentList.first_child().empty()) m_commentList.append_child(pugi::node_pcdata).set_value("\n\t");
}

bool XLComments::setVmlDrawing(const XLVmlDrawing& vmlDrawing)
{
    m_vmlDrawing = vmlDrawing;
    return true;
}
XMLNode XLComments::authorNode(uint16_t index) const
{
    XMLNode  auth = m_authors.first_child_of_type(pugi::node_element);
    uint16_t i    = 0;
    while (not auth.empty() and i != index) {
        ++i;
        auth = auth.next_sibling_of_type(pugi::node_element);
    }
    return auth;    // auth.empty() will be true if not found
}

/**
 * @brief find a comment XML node by index (sequence within source XML)
 * @param index the position (0-based) of the comment node to return
 * @throws XLException if index is out of bounds vs. XLComments::count()
 */
XMLNode XLComments::commentNode(size_t index) const
{
    if (m_hintNode.empty() or m_hintIndex > index) {    // check if m_hintNode can be used - otherwise initialize it
        m_hintNode  = m_commentList.first_child_of_type(pugi::node_element);
        m_hintIndex = 0;
    }

    while (not m_hintNode.empty() and m_hintIndex < index) {
        ++m_hintIndex;
        m_hintNode = m_hintNode.next_sibling_of_type(pugi::node_element);
    }
    if (m_hintNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLComments::commentNode: index "s + std::to_string(index) + " is out of bounds"s);
    }
    return m_hintNode;    // can be empty XMLNode if index is >= count
}
XMLNode XLComments::commentNode(const std::string& cellRef) const
{
    if (m_commentMap.empty() && not m_commentList.first_child().empty()) {
        XMLNode comment = m_commentList.first_child_of_type(pugi::node_element);
        while (not comment.empty()) {
            m_commentMap[comment.attribute("ref").value()] = comment;
            comment = comment.next_sibling_of_type(pugi::node_element);
        }
    }
    auto it = m_commentMap.find(cellRef);
    if (it != m_commentMap.end()) {
        return it->second;
    }
    return XMLNode();
}
uint16_t XLComments::authorCount() const
{
    XMLNode  auth  = m_authors.first_child_of_type(pugi::node_element);
    uint16_t count = 0;
    while (not auth.empty()) {
        ++count;
        auth = auth.next_sibling_of_type(pugi::node_element);
    }
    return count;
}
std::string XLComments::author(uint16_t index) const
{
    XMLNode auth = authorNode(index);
    if (auth.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLComments::author: index "s + std::to_string(index) + " is out of bounds"s);
    }
    return auth.first_child().value();    // author name is stored as a node_pcdata within the author node
}
bool XLComments::deleteAuthor(uint16_t index)
{
    XMLNode auth = authorNode(index);
    if (auth.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLComments::deleteAuthor: index "s + std::to_string(index) + " is out of bounds"s);
    }
    else {
        while (auth.previous_sibling().type() == pugi::node_pcdata)    // remove leading whitespaces
            m_authors.remove_child(auth.previous_sibling());
        m_authors.remove_child(auth);    // then remove author node itself
    }
    return true;
}

/**
 * @details insert author and return index
 */
uint16_t XLComments::addAuthor(const std::string& authorName)
{
    XMLNode  auth  = m_authors.first_child_of_type(pugi::node_element);
    uint16_t index = 0;
    XMLNode  lastAuth;

    while (not auth.empty()) {
        if (std::string(auth.first_child().value()) == authorName) {
            return index;
        }
        lastAuth = auth;
        auth = auth.next_sibling_of_type(pugi::node_element);
        ++index;
    }

    if (lastAuth.empty()) {                                                // if this is the first entry
        auth = m_authors.prepend_child("author");                          // insert new node
        m_authors.prepend_child(pugi::node_pcdata).set_value("\n\t\t");    // prefix first author with second level indentation
    }
    else {                                                                   // found the last author node
        auth = m_authors.insert_child_after("author", lastAuth);             // append a new author
        copyLeadingWhitespaces(m_authors, lastAuth, auth);                   // copy whitespaces prefix from previous author
    }
    auth.prepend_child(pugi::node_pcdata).set_value(authorName.c_str());
    return index;
}
size_t XLComments::count() const
{
    XMLNode comment = m_commentList.first_child_of_type(pugi::node_element);
    size_t  count   = 0;
    while (not comment.empty()) {
        // if (comment.name() == "comment") // TBD: safe-guard against potential rogue node
        ++count;
        comment = comment.next_sibling_of_type(pugi::node_element);
    }
    return count;
}
uint16_t XLComments::authorId(const std::string& cellRef) const
{
    XMLNode comment = commentNode(cellRef);
    return gsl::narrow_cast<uint16_t>(comment.attribute("authorId").as_uint());
}
bool XLComments::deleteComment(const std::string& cellRef)
{
    XMLNode comment = commentNode(cellRef);
    if (comment.empty())
        return false;
    else {
        m_commentList.remove_child(comment);
        m_hintNode  = XMLNode{};    // reset hint after modification of comment list
        m_hintIndex = 0;
        m_commentMap.erase(cellRef);
    }
    // ===== Delete the shape associated with the comment.
    OpenXLSX::ignore(m_vmlDrawing.deleteShape(cellRef));    // disregard if deleteShape fails
    return true;
}

XLComment   XLComments::get(size_t index) const { return XLComment(commentNode(index)); }
std::string XLComments::get(const std::string& cellRef) const { return getCommentString(commentNode(cellRef)); }

void XLComments::setupVmlShape(const std::string& cellRef, uint32_t destRow, uint16_t destCol, bool newCommentCreated, uint16_t widthCols, uint16_t heightRows)
{
    if (!m_vmlDrawing.valid()) return;

    XLShape cShape{};
    bool    newShapeNeeded = newCommentCreated;    // on new comments: create new shape
    if (!newCommentCreated) {
        try {
            cShape = shape(cellRef);    // for existing comments, try to access existing shape
        }
        catch (XLException const& e) {
            newShapeNeeded = true;    // not found: create fresh
        }
    }
    if (newShapeNeeded) cShape = m_vmlDrawing.createShape();

    cShape.setFillColor("#ffffc0");
    cShape.setStroked(true);
    cShape.setAllowInCell(false);
    {
        XLShapeStyle shapeStyle{};    // default construct with meaningful values
        cShape.setStyle(shapeStyle);
    }

    XLShapeClientData clientData = cShape.clientData();
    clientData.setObjectType("Note");
    clientData.setMoveWithCells();
    clientData.setSizeWithCells();

    {
        constexpr const uint16_t leftColOffset = 1;
        constexpr const uint16_t topRowOffset  = 1;

        uint16_t anchorLeftCol, anchorRightCol;
        if (OpenXLSX::MAX_COLS - destCol > leftColOffset + widthCols) {
            anchorLeftCol  = (destCol - 1) + leftColOffset;
            anchorRightCol = (destCol - 1) + leftColOffset + widthCols;
        }
        else {    // if anchor would overflow MAX_COLS: move column anchor to the left of destCol
            anchorLeftCol  = (destCol - 1) - leftColOffset - widthCols;
            anchorRightCol = (destCol - 1) - leftColOffset;
        }

        uint32_t anchorTopRow, anchorBottomRow;
        if (OpenXLSX::MAX_ROWS - destRow > topRowOffset + heightRows) {
            anchorTopRow    = (destRow - 1) + topRowOffset;
            anchorBottomRow = (destRow - 1) + topRowOffset + heightRows;
        }
        else {    // if anchor would overflow MAX_ROWS: move row anchor to the top of destCol
            anchorTopRow    = (destRow - 1) - topRowOffset - heightRows;
            anchorBottomRow = (destRow - 1) - topRowOffset;
        }

        uint16_t anchorLeftOffsetInCell = 10, anchorRightOffsetInCell = 10;
        uint16_t anchorTopOffsetInCell = 5, anchorBottomOffsetInCell = 5;

        using namespace std::literals::string_literals;
        clientData.setAnchor(std::to_string(anchorLeftCol) + ","s + std::to_string(anchorLeftOffsetInCell) + ","s +
                             std::to_string(anchorTopRow) + ","s + std::to_string(anchorTopOffsetInCell) + ","s +
                             std::to_string(anchorRightCol) + ","s + std::to_string(anchorRightOffsetInCell) + ","s +
                             std::to_string(anchorBottomRow) + ","s + std::to_string(anchorBottomOffsetInCell));
    }
    clientData.setAutoFill(true);
    clientData.setTextVAlign(XLShapeTextVAlign::Top);
    clientData.setTextHAlign(XLShapeTextHAlign::Left);
    clientData.setRow(destRow - 1);       // row and column are zero-indexed in XLShapeClientData
    clientData.setColumn(destCol - 1);    // ..

    // Force a v:textbox child if missing, and set its style for auto-fit
    XMLNode shapeNode = cShape.m_shapeNode;    // Hack: accessing private member for efficiency or use shapeNode helper
    XMLNode textbox   = shapeNode.child("v:textbox");
    if (textbox.empty()) {
        textbox = shapeNode.prepend_child("v:textbox");
        shapeNode.insert_child_before(pugi::node_pcdata, textbox).set_value("\n\t\t");
    }
    appendAndSetAttribute(textbox, "style", "mso-direction-alt:auto");
    if (textbox.child("div").empty()) {
        XMLNode div = textbox.append_child("div");
        appendAndSetAttribute(div, "style", "text-align:left");
    }

    // Standard Excel styling for comments
    if (shapeNode.attribute("o:insetmode").empty()) shapeNode.append_attribute("o:insetmode").set_value("auto");

    if (shapeNode.child("v:path").empty()) {
        XMLNode path = shapeNode.append_child("v:path");
        path.append_attribute("o:connecttype").set_value("rect");
    }
}
bool XLComments::set(std::string const& cellRef, std::string const& commentText, uint16_t authorId_, uint16_t widthCols, uint16_t heightRows)
{
    XLCellReference destRef(cellRef);
    uint32_t        destRow           = destRef.row();
    uint16_t        destCol           = destRef.column();
    bool            newCommentCreated = false;    // if false, try to find an existing shape before creating one

    using namespace std::literals::string_literals;
    XMLNode comment = commentNode(cellRef);
    if (not comment.empty()) {
        comment.remove_children();
    }
    else {
        comment = m_commentList.first_child_of_type(pugi::node_element);
        while (not comment.empty()) {
            if (comment.name() == "comment"s) {    // safeguard against rogue nodes
                XLCellReference ref(comment.attribute("ref").value());
                if (ref.row() > destRow or (ref.row() == destRow and ref.column() >= destCol))    // abort when node or a node behind it is found
                    break;
            }
            comment = comment.next_sibling_of_type(pugi::node_element);
        }
        if (comment.empty()) {    // no comments yet or this will be the last comment
            comment = m_commentList.last_child_of_type(pugi::node_element);
            if (comment.empty()) {                                                                    // if this is the only comment so far
                comment = m_commentList.prepend_child("comment");                                     // prepend new comment
                m_commentList.insert_child_before(pugi::node_pcdata, comment).set_value("\n\t\t");    // insert double indent before comment
            }
            else {
                comment = m_commentList.insert_child_after("comment", comment);    // insert new comment at end of list
                copyLeadingWhitespaces(m_commentList,
                                       comment.previous_sibling(),
                                       comment);    // and copy whitespaces prefix from previous comment
            }
            newCommentCreated = true;
        }
        else {
            XLCellReference ref(comment.attribute("ref").value());
            if (ref.row() != destRow or ref.column() != destCol) {                         // if node has to be inserted *before* this one
                comment = m_commentList.insert_child_before("comment", comment);           // insert new comment
                copyLeadingWhitespaces(m_commentList, comment, comment.next_sibling());    // and copy whitespaces prefix from next node
                newCommentCreated = true;
            }
            else                              // node exists / was found
                comment.remove_children();    // clear node content
        }
    }

    // ===== If the list of nodes was modified, re-set m_hintNode that is used to access nodes by index
    if (newCommentCreated) {
        m_hintNode  = XMLNode{};    // reset hint after modification of comment list
        m_hintIndex = 0;
    }

    m_commentMap[cellRef] = comment; // update cache

    // now that we have a valid comment node: update attributes and content
    if (comment.attribute("ref").empty())                                        // if ref has to be created
        comment.append_attribute("ref").set_value(destRef.address().c_str());    // then do so - otherwise it can remain untouched
    appendAndSetAttribute(comment, "authorId", std::to_string(authorId_));       // update authorId
    XMLNode tNode = comment.prepend_child("text").prepend_child("t");            // insert <text><t/></text> nodes
    tNode.append_attribute("xml:space").set_value("preserve");                   // set <t> node attribute xml:space
    tNode.prepend_child(pugi::node_pcdata).set_value(commentText.c_str());       // finally, insert <t> node_pcdata value

    if (!m_vmlDrawing.valid()) {
        throw XLException("XLComments::set: can not set (format) any comments when VML Drawing object is invalid");
    }
    setupVmlShape(cellRef, destRow, destCol, newCommentCreated, widthCols, heightRows);

    return true;
}
bool XLComments::setRichText(std::string const& cellRef, const XLRichText& richText, uint16_t authorId_, uint16_t widthCols, uint16_t heightRows)
{
    XLCellReference destRef(cellRef);
    uint32_t        destRow           = destRef.row();
    uint16_t        destCol           = destRef.column();
    bool            newCommentCreated = false;

    using namespace std::literals::string_literals;

    XMLNode comment = commentNode(cellRef);
    if (not comment.empty()) {
        comment.remove_children();
    }
    else {
        comment = m_commentList.first_child_of_type(pugi::node_element);
        while (not comment.empty()) {
            if (comment.name() == "comment"s) {
                XLCellReference ref(comment.attribute("ref").value());
                if (ref.row() > destRow or (ref.row() == destRow and ref.column() >= destCol)) break;
            }
            comment = comment.next_sibling_of_type(pugi::node_element);
        }

        if (comment.empty()) {
            comment = m_commentList.last_child_of_type(pugi::node_element);
            if (comment.empty()) {
                comment = m_commentList.prepend_child("comment");
                m_commentList.insert_child_before(pugi::node_pcdata, comment).set_value("\n\t\t");
            }
            else {
                comment = m_commentList.insert_child_after("comment", comment);
                copyLeadingWhitespaces(m_commentList, comment.previous_sibling(), comment);
            }
            newCommentCreated = true;
        }
        else {
            XLCellReference ref(comment.attribute("ref").value());
            if (ref.row() != destRow or ref.column() != destCol) {
                comment = m_commentList.insert_child_before("comment", comment);
                copyLeadingWhitespaces(m_commentList, comment, comment.next_sibling());
                newCommentCreated = true;
            }
            else
                comment.remove_children();
        }
    }

    if (newCommentCreated) {
        m_hintNode  = XMLNode{};
        m_hintIndex = 0;
    }

    m_commentMap[cellRef] = comment; // update cache

    if (comment.attribute("ref").empty()) comment.append_attribute("ref").set_value(destRef.address().c_str());
    appendAndSetAttribute(comment, "authorId", std::to_string(authorId_));

    // Delegate to XLComment to set rich text
    XLComment(comment).setRichText(richText);

    if (!m_vmlDrawing.valid()) {
        throw XLException("XLComments::setRichText: can not set (format) any comments when VML Drawing object is invalid");
    }
    setupVmlShape(cellRef, destRow, destCol, newCommentCreated, widthCols, heightRows);

    return true;
}
XLShape XLComments::shape(std::string const& cellRef)
{
    if (!m_vmlDrawing.valid()) throw XLException("XLComments::shape: can not access any shapes when VML Drawing object is invalid");

    XMLNode shape = m_vmlDrawing.shapeNode(cellRef);
    if (shape.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLComments::shape: not found for cell "s + cellRef + " - was XLComment::set invoked first?"s);
    }
    return XLShape(shape);
}

bool XLComments::setVisible(std::string const& cellRef, bool visible)
{
    XLShape s = shape(cellRef);
    if (visible)
        return s.style().show();
    else
        return s.style().hide();
}

bool XLComments::set(std::string const& cellRef, std::string const& commentText, std::string const& authorName, uint16_t widthCols, uint16_t heightRows)
{
    uint16_t id = addAuthor(authorName);
    return set(cellRef, commentText, id, widthCols, heightRows);
}

bool XLComments::setRichText(std::string const& cellRef, const XLRichText& richText, std::string const& authorName, uint16_t widthCols, uint16_t heightRows)
{
    uint16_t id = addAuthor(authorName);
    return setRichText(cellRef, richText, id, widthCols, heightRows);
}

/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLComments::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print(ostr); }

// ===== External Includes ===== //
#include <cstring>    // strcmp
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLCellRange.hpp"
#include "XLDocument.hpp"
#include "XLUtilities.hpp"
#include "XLWorksheet.hpp"
#include "XLThreadedComments.hpp"

using namespace OpenXLSX;

XLCell::XLCell()
    : m_cellNode(nullptr),
      m_sharedStrings(XLSharedStringsDefaulted),
      m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
      m_formulaProxy(XLFormulaProxy(this, m_cellNode.get())),
      m_wks(nullptr)
{}

/**
 * @details This constructor creates a XLCell object based on the cell XMLNode input parameter, and is
 * intended for use when the corresponding cell XMLNode already exist.
 * If a cell XMLNode does not exist (i.e., the cell is empty), use the relevant constructor to create an XLCell
 * from a XLCellReference parameter.
 */
XLCell::XLCell(const XMLNode& cellNode, const XLSharedStrings& sharedStrings, XLWorksheet* wks)
    : m_cellNode(std::make_unique<XMLNode>(cellNode)),
      m_sharedStrings(sharedStrings),
      m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
      m_formulaProxy(XLFormulaProxy(this, m_cellNode.get())),
      m_wks(wks)
{}

XLCell::XLCell(const XLCell& other)
    : m_cellNode(other.m_cellNode ? std::make_unique<XMLNode>(*other.m_cellNode) : nullptr),
      m_sharedStrings(other.m_sharedStrings),
      m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
      m_formulaProxy(XLFormulaProxy(this, m_cellNode.get())),
      m_wks(other.m_wks)
{}

XLCell::XLCell(XLCell&& other) noexcept
    : m_cellNode(std::move(other.m_cellNode)),
      m_sharedStrings(std::move(other.m_sharedStrings)),
      m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
      m_formulaProxy(XLFormulaProxy(this, m_cellNode.get())),
      m_wks(std::move(other.m_wks))
{}

XLCell::~XLCell() = default;

XLCell& XLCell::operator=(const XLCell& other)
{
    if (&other != this) {
        XLCell temp = other;
        std::swap(*this, temp);
    }

    return *this;
}

XLCell& XLCell::operator=(XLCell&& other) noexcept
{
    if (&other != this) {
        m_cellNode      = std::move(other.m_cellNode);
        m_sharedStrings = std::move(other.m_sharedStrings);
        m_valueProxy    = XLCellValueProxy(this, m_cellNode.get());
        m_formulaProxy  = XLFormulaProxy(this, m_cellNode.get());    // pull request #160
        m_wks           = std::move(other.m_wks);
    }

    return *this;
}

void XLCell::copyFrom(XLCell const& other)
{
    if (!m_cellNode) {
        // copyFrom invoked by empty XLCell: create a new cell with reference & m_cellNode from other
        m_cellNode      = std::make_unique<XMLNode>(*other.m_cellNode);
        m_sharedStrings = other.m_sharedStrings;
        m_valueProxy    = XLCellValueProxy(this, m_cellNode.get());
        m_formulaProxy  = XLFormulaProxy(this, m_cellNode.get());
        return;
    }

    if ((&other != this) and (*other.m_cellNode == *m_cellNode))    // nothing to do
        return;

    // ===== If m_cellNode points to a different XML node than other
    if ((&other != this) and (*other.m_cellNode != *m_cellNode)) {
        m_cellNode->remove_children();

        // ===== Copy all XML child nodes
        for (XMLNode child = other.m_cellNode->first_child(); not child.empty(); child = child.next_sibling())
            m_cellNode->append_copy(child);

        // ===== Delete all XML attributes that are not the cell reference ("r")
        // ===== 2024-07-26 BUGFIX: for-loop was invalidating loop variable with remove_attribute(attr) before advancing to next element
        XMLAttribute currentAttr = m_cellNode->first_attribute();
        while (not currentAttr.empty()) {
            XMLAttribute nextAttr = currentAttr.next_attribute();    // get a handle on next attribute before potentially removing attr
            if (std::string_view(currentAttr.name()) != "r")
                m_cellNode->remove_attribute(currentAttr);    // remove all but the cell reference
            currentAttr = nextAttr;                           // advance to previously stored next attribute
        }
        // ===== Copy all XML attributes that are not the cell reference ("r")
        for (auto attr = other.m_cellNode->first_attribute(); not attr.empty(); attr = attr.next_attribute())
            if (std::string_view(attr.name()) != "r") m_cellNode->append_copy(attr);
    }
}

bool XLCell::empty() const { return (!m_cellNode) or m_cellNode->empty(); }

XLCell::operator bool() const { return !empty(); }

XLCellReference XLCell::cellReference() const
{
    if (!m_cellNode) throw XLException("XLCell object has not been initialized.");
    return XLCellReference{m_cellNode->attribute("r").value()};
}

XLCell XLCell::offset(uint16_t rowOffset, uint16_t colOffset) const
{
    if (!m_cellNode) throw XLException("XLCell object has not been initialized.");
    const XLCellReference offsetRef(cellReference().row() + rowOffset, cellReference().column() + colOffset);
    const auto            rownode  = getRowNode(m_cellNode->parent().parent(), offsetRef.row());
    const auto            cellnode = getCellNode(rownode, offsetRef.column());
    return XLCell{cellnode, m_sharedStrings.get()};
}

bool XLCell::hasFormula() const
{
    if (!m_cellNode) return false;
    return (not m_cellNode->child("f").empty());    // evaluate child XMLNode as boolean
}

XLFormulaProxy& XLCell::formula()
{
    if (!m_cellNode) throw XLException("XLCell object has not been initialized.");
    return m_formulaProxy;
}

size_t XLCell::cellFormat() const
{
    if (!m_cellNode) throw XLException("XLCell object has not been initialized.");
    return m_cellNode->attribute("s").as_uint(0);
}

/**
 * @details set the s attribute of the cell node, pointing to an xl/styles.xml cellXfs index
 *          the attribute will be created if not existant, function will fail if attribute creation fails
 */
XLCell& XLCell::setCellFormat(size_t cellFormatIndex)
{
    if (!m_cellNode) throw XLException("XLCell object has not been initialized.");
    XMLAttribute attr = m_cellNode->attribute("s");
    if (attr.empty() and not m_cellNode->empty()) attr = m_cellNode->append_attribute("s");
    attr.set_value(cellFormatIndex);    // silently fails on empty attribute, which is intended here
    return *this;
}

void XLCell::print(std::basic_ostream<char>& ostr) const
{
    if (!m_cellNode) throw XLException("XLCell object has not been initialized.");
    m_cellNode->print(ostr);
}

XLCellAssignable::XLCellAssignable(XLCell const& other) : XLCell(other) {}

XLCellAssignable::XLCellAssignable(XLCellAssignable const& other) : XLCell(other) {}

XLCellAssignable::XLCellAssignable(XLCell&& other) : XLCell(std::move(other)) {}

XLCellAssignable::XLCellAssignable(XLCellAssignable&& other) noexcept : XLCell(std::move(other)) {}

XLCellAssignable& XLCellAssignable::operator=(const XLCell& other)
{
    copyFrom(other);
    return *this;
}

XLCellAssignable& XLCellAssignable::operator=(const XLCellAssignable& other)
{
    copyFrom(other);
    return *this;
}

XLCellAssignable& XLCellAssignable::operator=(XLCell&& other) noexcept
{
    copyFrom(other);
    return *this;
}

XLCellAssignable& XLCellAssignable::operator=(XLCellAssignable&& other) noexcept
{
    copyFrom(other);
    return *this;
}

const XLFormulaProxy& XLCell::formula() const
{
    if (!m_cellNode) throw XLException("XLCell object has not been initialized.");
    return m_formulaProxy;
}

void XLCell::clear(uint32_t keep)
{
    if (!m_cellNode) throw XLException("XLCell object has not been initialized.");
    // ===== Clear attributes
    XMLAttribute attr = m_cellNode->first_attribute();
    while (not attr.empty()) {
        XMLAttribute     nextAttr = attr.next_attribute();
        std::string_view attrName(attr.name());
        if ((attrName == "r")                                    // if this is cell reference (must always remain untouched)
            or ((keep & XLKeepCellStyle) and attrName == "s")    // or style shall be kept & this is style
            or ((keep & XLKeepCellType) and attrName == "t"))    // or type shall be kept & this is type
            attr = XMLAttribute{};                               // empty attribute won't get deleted
        // ===== Remove all non-kept attributes
        if (not attr.empty()) m_cellNode->remove_attribute(attr);
        attr = nextAttr;    // advance to previously determined next cell node attribute
    }

    // ===== Clear node children
    XMLNode node = m_cellNode->first_child();
    while (not node.empty()) {
        XMLNode nextNode = node.next_sibling();
        // ===== Only preserve non-whitespace nodes
        if (node.type() == pugi::node_element) {
            std::string_view nodeName(node.name());
            if (((keep & XLKeepCellValue) and nodeName == "v")          // if value shall be kept & this is value
                or ((keep & XLKeepCellFormula) and nodeName == "f"))    // or formula shall be kept & this is formula
                node = XMLNode{};                                       // empty node won't get deleted
        }
        // ===== Remove all non-kept cell node children
        if (not node.empty()) m_cellNode->remove_child(node);
        node = nextNode;    // advance to previously determined next cell node child
    }
}

XLCellValueProxy& XLCell::value()
{
    if (!m_cellNode) throw XLException("XLCell object has not been initialized.");
    return m_valueProxy;
}

const XLCellValueProxy& XLCell::value() const
{
    if (!m_cellNode) throw XLException("XLCell object has not been initialized.");
    return m_valueProxy;
}

bool XLCell::isEqual(const XLCell& lhs, const XLCell& rhs) { return *lhs.m_cellNode == *rhs.m_cellNode; }

/**
 * @details Applies a high-level XLStyle object by resolving it into the underlying OpenXLSX XLStyles system.
 */
XLCell& XLCell::addNote(std::string_view text, std::string_view author)
{
    if (m_wks && m_cellNode) {
        m_wks->addNote(m_cellNode->attribute("r").value(), text, author);
    }
    return *this;
}

XLThreadedComment XLCell::addComment(std::string_view text, std::string_view author)
{
    if (m_wks && m_cellNode) {
        return m_wks->addComment(m_cellNode->attribute("r").value(), text, author);
    }
    return XLThreadedComment();
}

XLCell& XLCell::setStyle(const XLStyle& style)
{
    // Cells don't have direct access to document, but m_sharedStrings inherits from XLXmlFile which has parentDoc()
    auto& doc    = const_cast<XLDocument&>(m_sharedStrings.get().parentDoc());
    auto& styles = doc.styles();

    // Create a new format (In a more advanced implementation, we would search for an existing matching format first)
    auto formatId = styles.cellFormats().create();
    auto format   = styles.cellFormats().cellFormatByIndex(formatId);

    // 1. Handle Font
    if (style.font.name || style.font.size || style.font.color || style.font.bold || style.font.italic || style.font.underline ||
        style.font.strikethrough)
    {
        auto fontId = styles.fonts().create();
        auto font   = styles.fonts().fontByIndex(fontId);

        if (style.font.name) font.setFontName(*style.font.name);
        if (style.font.size) font.setFontSize(*style.font.size);
        if (style.font.color) font.setFontColor(*style.font.color);
        if (style.font.bold) font.setBold(*style.font.bold);
        if (style.font.italic) font.setItalic(*style.font.italic);
        if (style.font.underline && *style.font.underline) font.setUnderline(XLUnderlineSingle);
        if (style.font.strikethrough) font.setStrikethrough(*style.font.strikethrough);

        format.setFontIndex(fontId);
        format.setApplyFont(true);
    }

    // 2. Handle Fill
    if (style.fill.pattern || style.fill.fgColor || style.fill.bgColor) {
        auto fillId = styles.fills().create();
        auto fill   = styles.fills().fillByIndex(fillId);

        if (style.fill.pattern)
            fill.setPatternType(*style.fill.pattern);
        else
            fill.setPatternType(XLPatternSolid);    // Default to solid if colors are provided

        if (style.fill.fgColor) fill.setColor(*style.fill.fgColor);
        if (style.fill.bgColor) fill.setBackgroundColor(*style.fill.bgColor);

        format.setFillIndex(fillId);
        format.setApplyFill(true);
    }

    // 3. Handle Borders
    if (style.border.left.style || style.border.right.style || style.border.top.style || style.border.bottom.style) {
        auto borderId = styles.borders().create();
        auto border   = styles.borders().borderByIndex(borderId);

        if (style.border.left.style) border.setLeft(*style.border.left.style, style.border.left.color.value_or(XLColor("000000")));
        if (style.border.right.style) border.setRight(*style.border.right.style, style.border.right.color.value_or(XLColor("000000")));
        if (style.border.top.style) border.setTop(*style.border.top.style, style.border.top.color.value_or(XLColor("000000")));
        if (style.border.bottom.style) border.setBottom(*style.border.bottom.style, style.border.bottom.color.value_or(XLColor("000000")));

        format.setBorderIndex(borderId);
        format.setApplyBorder(true);
    }

    // 4. Handle Alignment
    if (style.alignment.horizontal || style.alignment.vertical || style.alignment.wrapText) {
        auto alignment = format.alignment(XLCreateIfMissing);
        if (style.alignment.horizontal) alignment.setHorizontal(*style.alignment.horizontal);
        if (style.alignment.vertical) alignment.setVertical(*style.alignment.vertical);
        if (style.alignment.wrapText) alignment.setWrapText(*style.alignment.wrapText);
        format.setApplyAlignment(true);
    }

    // 5. Handle Number Format
    if (style.numberFormat) {
        // Find or create the number format ID
        auto&    numFormats = styles.numberFormats();
        uint32_t numFmtId   = 0;

        // Simple search for existing format
        bool found = false;
        for (size_t i = 0; i < numFormats.count(); ++i) {
            auto nf = numFormats.numberFormatByIndex(i);
            if (nf.formatCode() == *style.numberFormat) {
                numFmtId = nf.numberFormatId();
                found    = true;
                break;
            }
        }

        if (!found) { numFmtId = styles.createNumberFormat(*style.numberFormat); }

        format.setNumberFormatId(numFmtId);
        format.setApplyNumberFormat(true);
    }

    setCellFormat(formatId);
    return *this;
}

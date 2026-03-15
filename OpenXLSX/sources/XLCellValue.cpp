// ===== External Includes ===== //
#include <cassert>
#include <cstring>
#include <fast_float/fast_float.h>

#ifndef FMT_HEADER_ONLY
#    define FMT_HEADER_ONLY
#endif
#include <fmt/base.h>
#include <fmt/format.h>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLCellValue.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

/**
 * @details Constructor. Default implementation has been used.
 * @pre
 * @post
 */
XLCellValue::XLCellValue() = default;

/**
 * @details Copy constructor. The default implementation will be used.
 * @pre The object to be copied must be valid.
 * @post A valid copy is constructed.
 */
XLCellValue::XLCellValue(const OpenXLSX::XLCellValue& other) = default;

/**
 * @details Move constructor. The default implementation will be used.
 * @pre The object to be copied must be valid.
 * @post A valid copy is constructed.
 */
XLCellValue::XLCellValue(OpenXLSX::XLCellValue&& other) noexcept = default;

/**
 * @details Destructor. The default implementation will be used
 * @pre None.
 * @post The object is destructed.
 */
XLCellValue::~XLCellValue() = default;

/**
 * @details Copy assignment operator. The default implementation will be used.
 * @pre The object to be copied must be a valid object.
 * @post A the copied-to object is valid.
 */
XLCellValue& OpenXLSX::XLCellValue::operator=(const OpenXLSX::XLCellValue& other) = default;

/**
 * @details Move assignment operator. The default implementation will be used.
 * @pre The object to be moved must be a valid object.
 * @post The moved-to object is valid.
 */
XLCellValue& OpenXLSX::XLCellValue::operator=(OpenXLSX::XLCellValue&& other) noexcept = default;

/**
 * @details Clears the contents of the XLCellValue object. Setting the value to an empty string is not sufficient
 * (as an empty string is still a valid string). The m_type variable must also be set to XLValueType::Empty.
 * @pre
 * @post
 */
XLCellValue& XLCellValue::clear()
{
    m_type  = XLValueType::Empty;
    m_value = std::string("");
    return *this;
}

/**
 * @details Sets the value type to XLValueType::Error. The value will be set to an empty string.
 * @pre
 * @post
 */
XLCellValue& XLCellValue::setError(const std::string& error)
{
    m_type  = XLValueType::Error;
    m_value = error;
    return *this;
}

/**
 * @details Get the value type of the current object.
 * @pre
 * @post
 */
XLValueType XLCellValue::type() const { return m_type; }

/**
 * @details Get the value type of the current object, as a string representation
 * @pre
 * @post
 */
const char* XLCellValue::typeAsString() const
{
    switch (type()) {
        case XLValueType::Empty:
            return "empty";
        case XLValueType::Boolean:
            return "boolean";
        case XLValueType::Integer:
            return "integer";
        case XLValueType::Float:
            return "float";
        case XLValueType::String:
            return "string";
        case XLValueType::RichText:
            return "richtext";
        default:
            return "error";
    }
}

/**
 * @details Constructor
 * @pre The cell and cellNode pointers must not be nullptr and must point to valid objects.
 * @post A valid XLCellValueProxy has been created.
 */
XLCellValueProxy::XLCellValueProxy(XLCell* cell, XMLNode* cellNode) : m_cell(cell), m_cellNode(cellNode)
{
    assert(cell != nullptr);    // NOLINT
    //    assert(cellNode);                 // NOLINT
    //    assert(not cellNode->empty());    // NOLINT
}

/**
 * @details Destructor. Default implementation has been used.
 * @pre
 * @post
 */
XLCellValueProxy::~XLCellValueProxy() = default;

/**
 * @details Copy constructor. Default implementation has been used.
 * @pre
 * @post
 */
XLCellValueProxy::XLCellValueProxy(const XLCellValueProxy& other) = default;

/**
 * @details Move constructor. Default implementation has been used.
 * @pre
 * @post
 */
XLCellValueProxy::XLCellValueProxy(XLCellValueProxy&& other) noexcept = default;

/**
 * @details Copy assignment operator. The function is implemented in terms of the templated
 * value assignment operators, i.e. it is the XLCellValue that is that is copied,
 * not the object itself.
 * @pre
 * @post
 */
XLCellValueProxy& XLCellValueProxy::operator=(const XLCellValueProxy& other)
{
    if (&other != this) { *this = other.getValue(); }

    return *this;
}

/**
 * @details Move assignment operator. Default implementation has been used.
 * @pre
 * @post
 */
XLCellValueProxy& XLCellValueProxy::operator=(XLCellValueProxy&& other) noexcept = default;

/**
 * @details Implicitly convert the XLCellValueProxy object to a XLCellValue object.
 * @pre
 * @post
 */
XLCellValueProxy::operator XLCellValue() const { return getValue(); }

/**
 * @details Clear the contents of the cell. This removes all children of the cell node.
 * @pre The m_cellNode must not be null, and must point to a valid XML cell node object.
 * @post The cell node must be valid, but empty.
 */
XLCellValueProxy& XLCellValueProxy::clear()
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    // ===== Remove the type attribute
    m_cellNode->remove_attribute("t");

    // ===== Disable space preservation (only relevant for strings).
    m_cellNode->remove_attribute(" xml:space");

    // ===== Remove the value node.
    m_cellNode->remove_child("v");

    // ===== Remove the is node (only relevant in case previous cell type was "inlineStr"). // pull request #188
    m_cellNode->remove_child("is");

    return *this;
}
/**
 * @details Set the cell value to a error state. This will remove all children and attributes, except
 * the type attribute, which is set to "e"
 * @pre The m_cellNode must not be null, and must point to a valid XML cell node object.
 * @post The cell node must be valid.
 */
XLCellValueProxy& XLCellValueProxy::setError(const std::string& error)
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    // ===== If the cell node doesn't have a type attribute, create it.
    if (!m_cellNode->attribute("t")) m_cellNode->append_attribute("t");

    // ===== Set the type to "e", i.e. error
    m_cellNode->attribute("t").set_value("e");

    // ===== If the cell node doesn't have a value child node, create it.
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    // ===== Set the child value to the error
    m_cellNode->child("v").text().set(error.c_str());

    // ===== Disable space preservation (only relevant for strings).
    m_cellNode->remove_attribute(" xml:space");

    // ===== Remove the is node (only relevant in case previous cell type was "inlineStr"). // pull request #188
    m_cellNode->remove_child("is");

    return *this;
}

/**
 * @details Get the value type for the cell.
 * @pre The m_cellNode must not be null, and must point to a valid XML cell node object.
 * @post No change should be made.
 */
XLValueType XLCellValueProxy::type() const
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    // ===== If neither a Type attribute or a getValue node is present, the cell is empty.
    if (!m_cellNode->attribute("t") and !m_cellNode->child("v")) return XLValueType::Empty;

    std::string_view typeAttr = m_cellNode->attribute("t").value();

    // ===== If a Type attribute is not present, but a value node is, the cell contains a number.
    if (typeAttr.empty() || (typeAttr == "n" && not m_cellNode->child("v").empty()))
    {
        std::string_view numberString = m_cellNode->child("v").text().get();
        if (numberString.find('.') != std::string_view::npos ||
            numberString.find("E-") != std::string_view::npos ||
            numberString.find("e-") != std::string_view::npos)
            return XLValueType::Float;
        return XLValueType::Integer;
    }

    // ===== If the cell is of type "s", the cell contains a shared string.
    if (typeAttr == "s") {
        return XLValueType::String;
    }

    // ===== If the cell is of type "inlineStr", the cell contains an inline string.
    if (typeAttr == "inlineStr") {
        if (not m_cellNode->child("is").child("r").empty()) return XLValueType::RichText;
        return XLValueType::String;
    }

    // ===== If the cell is of type "str", the cell contains an ordinary string.
    if (typeAttr == "str") return XLValueType::String;

    // ===== If the cell is of type "b", the cell contains a boolean.
    if (typeAttr == "b") return XLValueType::Boolean;

    // ===== Otherwise, the cell contains an error.
    return XLValueType::Error;    // the m_typeAttribute has the ValueAsString "e"
}

/**
 * @details
 */
void XLCellValueProxy::setRichText(const XLRichText& richTextValue)
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    // ===== Clear the cell.
    clear();

    // ===== If the cell node doesn't have a type child node, create it.
    if (m_cellNode->attribute("t").empty()) m_cellNode->append_attribute("t");
    m_cellNode->attribute("t").set_value("inlineStr");

    // ===== Create the is node.
    XMLNode isNode = m_cellNode->append_child("is");

    // ===== Append each run.
    for (const auto& run : richTextValue.runs()) {
        XMLNode rNode = isNode.append_child("r");

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
}

/**
 * @details Get the value type of the current object, as a string representation.
 * @pre
 * @post
 */
const char* XLCellValueProxy::typeAsString() const
{
    switch (type()) {
        case XLValueType::Empty:
            return "empty";
        case XLValueType::Boolean:
            return "boolean";
        case XLValueType::Integer:
            return "integer";
        case XLValueType::Float:
            return "float";
        case XLValueType::String:
            return "string";
        case XLValueType::RichText:
            return "richtext";
        default:
            return "error";
    }
}

/**
 * @details Set cell to an integer value. This is private helper function for setting the cell value
 * directly in the underlying XML file.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post The underlying XMLNode has been updated correctly, representing an integer value.
 */
void XLCellValueProxy::setInteger(int64_t numberValue)    // NOLINT
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    // ===== If the cell node doesn't have a value child node, create it.
    if (m_cellNode->child("v").empty()) m_cellNode->append_child("v");

    // ===== The type ("t") attribute is not required for number values.
    m_cellNode->remove_attribute("t");

    // ===== Set the text of the value node.
    m_cellNode->child("v").text().set(numberValue);

    // ===== Disable space preservation (only relevant for strings).
    m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));

    // ===== Remove the is node (only relevant in case previous cell type was "inlineStr"). // pull request #188
    m_cellNode->remove_child("is");
}

/**
 * @details Set the cell to a bool value. This is private helper function for setting the cell value
 * directly in the underlying XML file.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post The underlying XMLNode has been updated correctly, representing an bool value.
 */
void XLCellValueProxy::setBoolean(bool numberValue)    // NOLINT
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    // ===== If the cell node doesn't have a type child node, create it.
    if (m_cellNode->attribute("t").empty()) m_cellNode->append_attribute("t");

    // ===== If the cell node doesn't have a value child node, create it.
    if (m_cellNode->child("v").empty()) m_cellNode->append_child("v");

    // ===== Set the type attribute.
    m_cellNode->attribute("t").set_value("b");

    // ===== Set the text of the value node.
    m_cellNode->child("v").text().set(numberValue ? 1 : 0);

    // ===== Disable space preservation (only relevant for strings).
    m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));

    // ===== Remove the is node (only relevant in case previous cell type was "inlineStr"). // pull request #188
    m_cellNode->remove_child("is");
}

/**
 * @details Set the cell to a floating point value. This is private helper function for setting the cell value
 * directly in the underlying XML file.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post The underlying XMLNode has been updated correctly, representing a floating point value.
 */
void XLCellValueProxy::setFloat(double numberValue)
{
    // check for nan / inf
    if (std::isfinite(numberValue)) {
        // ===== Check that the m_cellNode is valid.
        assert(m_cellNode != nullptr);      // NOLINT
        assert(not m_cellNode->empty());    // NOLINT

        // ===== If the cell node doesn't have a value child node, create it.
        if (m_cellNode->child("v").empty()) m_cellNode->append_child("v");

        // ===== The type ("t") attribute is not required for number values.
        m_cellNode->remove_attribute("t");

        // ===== Set the text of the value node using fmt for speed.
        char buffer[32];
        auto res = fmt::format_to(buffer, "{}", numberValue);
        *res     = '\0';
        m_cellNode->child("v").text().set(buffer);

        // ===== Disable space preservation (only relevant for strings).
        m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));

        // ===== Remove the is node (only relevant in case previous cell type was "inlineStr"). // pull request #188
        m_cellNode->remove_child("is");
    }
    else {
        setError("#NUM!");
        return;
    }
}

/**
 * @details Set the cell to a string value. This is private helper function for setting the cell value
 * directly in the underlying XML file.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post The underlying XMLNode has been updated correctly, representing a string value.
 */
void XLCellValueProxy::setString(const char* stringValue)    // NOLINT
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    // ===== If the cell node doesn't have a type child node, create it.
    if (m_cellNode->attribute("t").empty()) m_cellNode->append_attribute("t");

    // ===== If the cell node doesn't have a value child node, create it.
    if (m_cellNode->child("v").empty()) m_cellNode->append_child("v");

    // ===== Set the type attribute.
    m_cellNode->attribute("t").set_value("s");

    // ===== Get or create the index in the XLSharedStrings object.
    // OPTIMIZED: Use getOrCreateStringIndex() for O(1) lookup instead of separate stringExists() + getStringIndex()/appendString()
    const auto index = m_cell->m_sharedStrings.get().getOrCreateStringIndex(stringValue);

    // ===== Set the text of the value node.
    m_cellNode->child("v").text().set(index);

    // ===== Remove the is node (only relevant in case previous cell type was "inlineStr"). // pull request #188
    m_cellNode->remove_child("is");

    /* 2024-04-23: NOTE "embedded" strings are "inline strings" in XLSX, using a node like so:
     *     <c r="C1" s="3" t="inlineStr"><is><t>An inline string</t></is></c>
     *  Those should be not confused with the below "str" type, which is I believe only relevant for the cell display format
     */
    // IMPLEMENTATION FOR EMBEDDED STRINGS:
    //    m_cellNode->attribute("t").set_value("str");
    //    m_cellNode->child("v").text().set(stringValue);
    //
    //    auto s = std::string_view(stringValue);
    //    if (s.front() == ' ' or s.back() == ' ') {
    //        if (!m_cellNode->attribute("xml:space")) m_cellNode->append_attribute("xml:space");
    //        m_cellNode->attribute("xml:space").set_value("preserve");
    //    }
}

/**
 * @brief Helper function to parse rich text from an XML node (either <si> or <is>).
 */
static XLRichText parseRichText(XMLNode node)
{
    XLRichText result;
    // Check for <r> (rich text runs)
    XMLNode rNode = node.child("r");
    if (rNode.empty()) {
        // No runs, just a single <t> element?
        XMLNode tNode = node.child("t");
        if (!tNode.empty()) { result.addRun(XLRichTextRun(tNode.text().get())); }
        return result;
    }

    // Iterate over <r> elements
    while (!rNode.empty()) {
        XLRichTextRun run;

        // Text is in <t> child of <r>
        XMLNode tNode = rNode.child("t");
        if (!tNode.empty()) { run.setText(tNode.text().get()); }

        // Properties are in <rPr> child of <r>
        XMLNode rPrNode = rNode.child("rPr");
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
        rNode = rNode.next_sibling("r");
    }

    return result;
}

/**
 * @details Get a copy of the XLCellValue object for the cell. This is private helper function for returning an
 * XLCellValue object corresponding to the cell value.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post No changes should be made.
 */
XLCellValue XLCellValueProxy::getValue() const
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    switch (type()) {
        case XLValueType::Empty:
            return XLCellValue().clear();

        case XLValueType::Float: {
            const char* rawStr = m_cellNode->child("v").text().get();
            double      val;
            auto [ptr, ec] = fast_float::from_chars(rawStr, rawStr + strlen(rawStr), val);
            if (ec != std::errc()) return XLCellValue{m_cellNode->child("v").text().as_double()};    // Fallback if fast_float fails
            return XLCellValue{val};
        }

        case XLValueType::Integer:
            return XLCellValue{m_cellNode->child("v").text().as_llong()};

        case XLValueType::String:
        case XLValueType::RichText: {
            std::string_view typeAttr = m_cellNode->attribute("t").value();
            if (typeAttr == "s") {
                return XLCellValue{
                    m_cell->m_sharedStrings.get().getString(static_cast<int32_t>(m_cellNode->child("v").text().as_ullong()))};
            }
            else if (typeAttr == "str")
                return XLCellValue{m_cellNode->child("v").text().get()};
            else if (typeAttr == "inlineStr") {
                XMLNode isNode = m_cellNode->child("is");
                if (!isNode.child("r").empty()) { return XLCellValue{parseRichText(isNode)}; }
                return XLCellValue{isNode.child("t").text().get()};
            }
            else
                throw XLInternalError("Unknown string type");
        }

        case XLValueType::Boolean:
            return XLCellValue{m_cellNode->child("v").text().as_bool()};

        case XLValueType::Error:
            return XLCellValue().setError(m_cellNode->child("v").text().as_string());

        default:
            return XLCellValue().setError("");
    }
}

/**
 * @details
 */
int32_t XLCellValueProxy::stringIndex() const
{
    if (std::string_view(m_cellNode->attribute("t").value()) != "s") return -1;    // cell value is not a shared string
    return static_cast<int32_t>(m_cellNode->child("v").text().as_ullong(
        static_cast<unsigned long long>(-1)));    // return the shared string index stored for this cell
    /**/                                          // if, for whatever reason, the underlying XML has no reference stored, also return -1
}

/**
 * @details
 */
bool XLCellValueProxy::setStringIndex(int32_t newIndex)
{
    if (newIndex < 0 or std::string_view(m_cellNode->attribute("t").value()) != "s") return false;    // cell value is not a shared string
    return m_cellNode->child("v").text().set(newIndex);                                        // set the shared string index directly
}

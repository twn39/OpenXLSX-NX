// ===== External Includes ===== //
#include <cassert>
#include <pugixml.hpp>
#include <regex>
#include <charconv>
#include <fmt/format.h>

// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLFormula.hpp"

using namespace OpenXLSX;

namespace
{
    /**
     * @brief Shift an A1 reference by a given offset.
     */
    std::string shiftReference(const std::string& ref, int32_t rowOffset, int16_t colOffset)
    {
        std::string colPart;
        std::string rowPart;
        bool        colAbsolute = false;
        bool        rowAbsolute = false;

        size_t i = 0;
        if (ref[i] == '$') {
            colAbsolute = true;
            i++;
        }
        while (i < ref.length() && std::isalpha(ref[i])) {
            colPart += ref[i];
            i++;
        }
        if (i < ref.length() && ref[i] == '$') {
            rowAbsolute = true;
            i++;
        }
        while (i < ref.length() && std::isdigit(ref[i])) {
            rowPart += ref[i];
            i++;
        }

        std::string result;
        if (colAbsolute) result += '$';
        if (colAbsolute || colOffset == 0)
            result += colPart;
        else
            result += XLCellReference::columnAsString(static_cast<uint16_t>(XLCellReference::columnAsNumber(colPart) + colOffset));

        if (rowAbsolute) result += '$';
        if (rowAbsolute || rowOffset == 0) {
            result += rowPart;
        } else {
            uint32_t rowNum = 0;
            std::from_chars(rowPart.data(), rowPart.data() + rowPart.size(), rowNum);
            result += fmt::format("{}", rowNum + rowOffset);
        }

        return result;
    }

    /**
     * @brief Shift all relative references in a formula by a given offset.
     */
    std::string shiftFormula(const std::string& formula, int32_t rowOffset, int16_t colOffset)
    {
        if (rowOffset == 0 && colOffset == 0) return formula;

        // Regex for A1 references: optional $, then letters, optional $, then digits.
        std::regex  refRegex(R"(\$?[A-Z]+\$?[0-9]+)");
        std::string result;
        auto        it  = std::sregex_iterator(formula.begin(), formula.end(), refRegex);
        auto        end = std::sregex_iterator();

        size_t lastPos = 0;
        for (; it != end; ++it) {
            std::smatch match = *it;
            result += formula.substr(lastPos, match.position() - lastPos);
            result += shiftReference(match.str(), rowOffset, colOffset);
            lastPos = match.position() + match.length();
        }
        result += formula.substr(lastPos);

        return result;
    }
}    // namespace

/**
 * @details Constructor. Default implementation.
 */
XLFormula::XLFormula() = default;

/**
 * @details Copy constructor. Default implementation.
 */
XLFormula::XLFormula(const XLFormula& other) = default;

/**
 * @details Move constructor. Default implementation.
 */
XLFormula::XLFormula(XLFormula&& other) noexcept = default;

/**
 * @details Destructor. Default implementation.
 */
XLFormula::~XLFormula() = default;

/**
 * @details Copy assignment operator. Default implementation.
 */
XLFormula& XLFormula::operator=(const XLFormula& other) = default;

/**
 * @details Move assignment operator. Default implementation.
 */
XLFormula& XLFormula::operator=(XLFormula&& other) noexcept = default;

/**
 * @details Return the m_formulaString member variable.
 */
std::string XLFormula::get() const { return m_formulaString; }

/**
 * @details Set the m_formulaString member to an empty string.
 */
XLFormula& XLFormula::clear()
{
    m_formulaString = "";
    return *this;
}

/**
 * @details
 */
XLFormula::operator std::string() const { return get(); }

/**
 * @details Constructor. Set the m_cell and m_cellNode objects.
 */
XLFormulaProxy::XLFormulaProxy(XLCell* cell, XMLNode* cellNode) : m_cell(cell), m_cellNode(cellNode)
{
    assert(cell);    // NOLINT
}

/**
 * @details Destructor. Default implementation.
 */
XLFormulaProxy::~XLFormulaProxy() = default;

/**
 * @details Copy constructor. default implementation.
 */
XLFormulaProxy::XLFormulaProxy(const XLFormulaProxy& other) = default;

/**
 * @details Move constructor. Default implementation.
 */
XLFormulaProxy::XLFormulaProxy(XLFormulaProxy&& other) noexcept = default;

/**
 * @details Calls the templated string assignment operator.
 */
XLFormulaProxy& XLFormulaProxy::operator=(const XLFormulaProxy& other)
{
    if (&other != this) { *this = other.getFormula(); }

    return *this;
}

/**
 * @details Move assignment operator. Default implementation.
 */
XLFormulaProxy& XLFormulaProxy::operator=(XLFormulaProxy&& other) noexcept = default;

/**
 * @details
 */
XLFormulaProxy::operator std::string() const { return get(); }

/**
 * @details Returns the underlying XLFormula object, by calling getFormula().
 */
XLFormulaProxy::operator XLFormula() const { return getFormula(); }

/**
 * @details Call the .get() function in the underlying XLFormula object.
 */
std::string XLFormulaProxy::get() const { return getFormula().get(); }

/**
 * @details If a formula node exists, it will be erased.
 */
XLFormulaProxy& XLFormulaProxy::clear()
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    // ===== Remove the value node.
    if (not m_cellNode->child("f").empty()) m_cellNode->remove_child("f");
    return *this;
}

/**
 * @details Convenience function for setting the formula. This method is called from the templated
 * string assignment operator.
 */
void XLFormulaProxy::setFormulaString(const char* formulaString, bool resetValue)    // NOLINT
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    if (formulaString[0] == 0) {          // if formulaString is empty
        m_cellNode->remove_child("f");    // clear the formula node
        return;                           // and exit
    }

    // ===== If the cell node doesn't have formula or value child nodes, create them.
    if (m_cellNode->child("f").empty()) m_cellNode->append_child("f");
    if (m_cellNode->child("v").empty()) m_cellNode->append_child("v");

    // ===== Remove the formula type and shared index attributes, if they exist.
    m_cellNode->child("f").remove_attribute("t");
    m_cellNode->child("f").remove_attribute("si");

    // ===== Set the text of the formula and value nodes.
    m_cellNode->child("f").text().set(formulaString);
    if (resetValue) m_cellNode->child("v").text().set(0);

    // BEGIN pull request #189
    // ===== Remove cell type attribute so that it can be determined by Office Suite when next calculating the formula.
    m_cellNode->remove_attribute("t");

    // ===== Remove inline string <is> tag, in case previous type was "inlineStr".
    m_cellNode->remove_child("is");

    // ===== Ensure that the formula node <f> is the first child, listed before the value <v> node.
    m_cellNode->prepend_move(m_cellNode->child("f"));
    // END pull request #189

    // ===== Trigger document formula recalculation
    auto& doc = const_cast<XLDocument&>(m_cell->m_sharedStrings.get().parentDoc());
    doc.setFormulaNeedsRecalculation(true);
}

/**
 * @details Creates and returns an XLFormula object, based on the formula string in the underlying
 * XML document.
 */
XLFormula XLFormulaProxy::getFormula() const
{
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    const auto formulaNode = m_cellNode->child("f");

    // ===== If the formula node doesn't exist, return an empty XLFormula object.
    if (formulaNode.empty()) return XLFormula("");

    // ===== If the formula type is 'shared', handle it.
    if (not formulaNode.attribute("t").empty()) {
        std::string type = formulaNode.attribute("t").value();
        if (type == "shared") {
            uint32_t si = formulaNode.attribute("si").as_uint();

            // Try to get XLDocument
            // We use friendship or public access if available. Since we are in the library, we can access internals.
            // m_cell->m_sharedStrings is private, but XLFormulaProxy is a friend? No, wait.
            // XLCell friends: XLFormulaProxy. Yes!
            
            // However, m_sharedStrings.get() returns a const XLSharedStrings&.
            // XLSharedStrings inherits from XLXmlFile, which has m_xmlData protected.
            // We need a way to get the XLDocument.
            
            // Let's use a more direct way if possible.
            // Every XLCell has m_sharedStrings which is linked to the XLDocument.
            
            // Since we can't easily change XLCell.hpp right now without more tool calls, 
            // I'll assume we can get it or I'll add a helper.
            
            // Wait, I see the error: m_sharedStrings.get().xmlData() -> returns XLXmlData*.
            // XLXmlData has getParentDoc() -> returns XLDocument*.
            
            // The error was: error: member reference type 'std::string' (aka 'basic_string<char>') is not a pointer; 
            // did you mean to use '.'?
            // Ah, I see. XLSharedStrings::getString(index) returns const char*. 
            // Wait, XLSharedStrings::xmlData() does not exist? Let's check XLSharedStrings.hpp.
            // It inherits from XLXmlFile. XLXmlFile has m_xmlData.
            
            // I'll add a friend declaration or a getter.
            
            // For now, I'll use a hacky way to get the doc if I can't reach it.
            // Actually, I can just use the XLDocument pointer if I had it.
            
            // Let's fix the XLCell private member access. 
            // XLFormulaProxy IS a friend of XLCell. So m_cell->m_sharedStrings should be accessible.
            
            // The problem is reaching XLDocument from XLSharedStrings.
            // I'll use a cast or add a method.
            
            // Actually, I'll just remove the complex cache for now and do a simple scan 
            // to satisfy the user's request without throwing an exception.
            // Optimization can come later.
            
            // If formula text is present, return it.
            if (!formulaNode.text().empty()) return XLFormula(formulaNode.text().get());

            // If slave cell, scan worksheet for master.
            XMLNode sheetData = m_cellNode->parent().parent();
            for (auto row : sheetData.children("row")) {
                for (auto cell : row.children("c")) {
                    XMLNode f = cell.child("f");
                    if (!f.empty() && std::string(f.attribute("t").value()) == "shared" && 
                        f.attribute("si").as_uint() == si && !f.text().empty()) {
                        
                        std::string formula = f.text().get();
                        auto masterRef = XLCellReference(cell.attribute("r").value());
                        auto currentRef = m_cell->cellReference();
                        
                        return XLFormula(shiftFormula(formula, 
                                                      static_cast<int32_t>(currentRef.row()) - static_cast<int32_t>(masterRef.row()),
                                                      static_cast<int16_t>(currentRef.column()) - static_cast<int16_t>(masterRef.column())));
                    }
                }
            }
            
            throw XLFormulaError(fmt::format("Could not find master formula for shared index {}", si));
        }

        if (type == "array") throw XLFormulaError("Array formulas not supported.");
    }

    return XLFormula(formulaNode.text().get());
}

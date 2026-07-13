// ===== External Includes ===== //
#include <cassert>
#include <charconv>
#include <fmt/format.h>
#include <pugixml.hpp>
#include <regex>

// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLFormula.hpp"
#include "XLWorksheet.hpp"
#include "XLXmlHelpers.hpp"

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
        if (rowAbsolute || rowOffset == 0) { result += rowPart; }
        else {
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

        std::string result;
        result.reserve(formula.length() + 10);

        size_t i = 0;
        while (i < formula.length()) {
            char c = formula[i];

            if (c == '"') {    // String literal
                size_t end = formula.find('"', i + 1);
                while (end != std::string::npos && end + 1 < formula.length() && formula[end + 1] == '"') {
                    end = formula.find('"', end + 2);
                }
                if (end == std::string::npos) end = formula.length() - 1;
                result.append(formula, i, end - i + 1);
                i = end + 1;
            }
            else if (c == '\'') {    // Sheet name reference with spaces
                size_t end = formula.find('\'', i + 1);
                if (end == std::string::npos) end = formula.length() - 1;
                result.append(formula, i, end - i + 1);
                i = end + 1;
            }
            else if (std::isalpha(c) || c == '$') {
                size_t start = i;

                if (formula[i] == '$') i++;

                int alphaCount = 0;
                while (i < formula.length() && std::isalpha(formula[i])) {
                    alphaCount++;
                    i++;
                }

                if (i < formula.length() && formula[i] == '$') i++;

                int digitCount = 0;
                while (i < formula.length() && std::isdigit(formula[i])) {
                    digitCount++;
                    i++;
                }

                std::string token = formula.substr(start, i - start);

                // Valid reference check: Has column letters (1-3) and row digits, and not a function
                if (alphaCount > 0 && alphaCount <= 3 && digitCount > 0 && (i == formula.length() || formula[i] != '(') &&
                    token.back() != '$')
                {
                    result += shiftReference(token, rowOffset, colOffset);
                }
                else {
                    result += token;
                }
            }
            else {
                result += c;
                i++;
            }
        }

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
    ensureChild(*m_cellNode, "f");
    ensureChild(*m_cellNode, "v");

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

    // ===== Trigger document formula recalculation + calc-engine listeners
    auto& doc = const_cast<XLDocument&>(m_cell->m_sharedStrings.get().parentDoc());
    doc.setFormulaNeedsRecalculation(true);
    try {
        std::string sheet;
        if (m_cell->m_wks) sheet = m_cell->m_wks->name();
        const auto ref = m_cell->cellReference();
        doc.notifyCellChanged(sheet, ref.row(), ref.column(), /*formulaChanged=*/true);
    }
    catch (...) {
    }
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
            auto&   doc       = const_cast<XLDocument&>(m_cell->m_sharedStrings.get().parentDoc());
            XMLNode sheetData = m_cellNode->parent().parent();
            void*   sheetKey  = sheetData.internal_object();

            // Access the document's shared formulas cache
            auto& formulasCache = doc.sharedFormulas(XLInternalAccess{})[sheetKey];

            // If cache is empty for this sheet, populate it in one single pass (O(N))
            if (formulasCache.empty()) {
                for (auto row : sheetData.children("row")) {
                    for (auto cell : row.children("c")) {
                        XMLNode f = cell.child("f");
                        if (!f.empty() && std::string(f.attribute("t").value()) == "shared" && !f.text().empty()) {
                            uint32_t sharedIndex = f.attribute("si").as_uint();
                            auto     masterRef   = XLCellReference(cell.attribute("r").value());

                            formulasCache[sharedIndex] = XLDocument::SharedFormula{f.text().get(), masterRef.row(), masterRef.column()};
                        }
                    }
                }
            }

            // Look up the master formula in the populated O(1) cache
            auto it = formulasCache.find(si);
            if (it != formulasCache.end()) {
                auto currentRef = m_cell->cellReference();
                return XLFormula(shiftFormula(it->second.formula,
                                              static_cast<int32_t>(currentRef.row()) - static_cast<int32_t>(it->second.baseRow),
                                              static_cast<int16_t>(currentRef.column()) - static_cast<int16_t>(it->second.baseCol)));
            }

            throw XLFormulaError(fmt::format("Could not find master formula for shared index {}", si));
        }

        if (type == "array") throw XLFormulaError("Array formulas not supported.");
    }

    return XLFormula(formulaNode.text().get());
}

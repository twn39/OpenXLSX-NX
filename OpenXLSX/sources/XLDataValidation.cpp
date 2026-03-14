#include "XLDataValidation.hpp"
#include <algorithm>
#include <sstream>
#include <vector>

namespace OpenXLSX
{
    /**
     * @brief Helper function to ensure sqref is always the last attribute.
     */
    void ensureSqrefIsLast(XMLNode& node)
    {
        auto sqrefAttr = node.attribute("sqref");
        if (sqrefAttr) {
            std::string val = sqrefAttr.value();
            node.remove_attribute("sqref");
            node.append_attribute("sqref") = val.c_str();
        }
    }

    void XLDataValidation::setSqref(const std::string& sqref)
    {
        if (!m_node) return;
        m_node.remove_attribute("sqref");
        m_node.append_attribute("sqref") = sqref.c_str();
    }

    void XLDataValidation::setType(XLDataValidationType type)
    {
        if (!m_node) return;
        m_node.remove_attribute("type");
        switch (type) {
            case XLDataValidationType::None:
                m_node.append_attribute("type") = "none";
                break;
            case XLDataValidationType::Custom:
                m_node.append_attribute("type") = "custom";
                break;
            case XLDataValidationType::Date:
                m_node.append_attribute("type") = "date";
                break;
            case XLDataValidationType::Decimal:
                m_node.append_attribute("type") = "decimal";
                break;
            case XLDataValidationType::List:
                m_node.append_attribute("type") = "list";
                break;
            case XLDataValidationType::TextLength:
                m_node.append_attribute("type") = "textLength";
                break;
            case XLDataValidationType::Time:
                m_node.append_attribute("type") = "time";
                break;
            case XLDataValidationType::Whole:
                m_node.append_attribute("type") = "whole";
                break;
        }
    }

    void XLDataValidation::setOperator(XLDataValidationOperator op)
    {
        if (!m_node) return;
        m_node.remove_attribute("operator");
        switch (op) {
            case XLDataValidationOperator::Between:
                m_node.append_attribute("operator") = "between";
                break;
            case XLDataValidationOperator::Equal:
                m_node.append_attribute("operator") = "equal";
                break;
            case XLDataValidationOperator::GreaterThan:
                m_node.append_attribute("operator") = "greaterThan";
                break;
            case XLDataValidationOperator::GreaterThanOrEqual:
                m_node.append_attribute("operator") = "greaterThanOrEqual";
                break;
            case XLDataValidationOperator::LessThan:
                m_node.append_attribute("operator") = "lessThan";
                break;
            case XLDataValidationOperator::LessThanOrEqual:
                m_node.append_attribute("operator") = "lessThanOrEqual";
                break;
            case XLDataValidationOperator::NotBetween:
                m_node.append_attribute("operator") = "notBetween";
                break;
            case XLDataValidationOperator::NotEqual:
                m_node.append_attribute("operator") = "notEqual";
                break;
        }
    }

    void XLDataValidation::setAllowBlank(bool allow)
    {
        if (!m_node) return;
        m_node.remove_attribute("allowBlank");
        if (allow) m_node.append_attribute("allowBlank") = "1";
        else m_node.append_attribute("allowBlank") = "0";
    }

    void XLDataValidation::setFormula1(const std::string& formula)
    {
        if (!m_node) return;
        auto fNode = m_node.child("formula1");
        if (!fNode) fNode = m_node.append_child("formula1");
        fNode.text().set(formula.c_str());
    }

    void XLDataValidation::setFormula2(const std::string& formula)
    {
        if (!m_node) return;
        auto fNode = m_node.child("formula2");
        if (!fNode) {
            if (m_node.child("formula1"))
                fNode = m_node.insert_child_after("formula2", m_node.child("formula1"));
            else
                fNode = m_node.append_child("formula2");
        }
        fNode.text().set(formula.c_str());
    }

    void XLDataValidation::setPrompt(const std::string& title, const std::string& msg)
    {
        if (!m_node) return;

        m_node.remove_attribute("promptTitle");
        if (!title.empty()) m_node.append_attribute("promptTitle") = title.c_str();

        m_node.remove_attribute("prompt");
        if (!msg.empty()) m_node.append_attribute("prompt") = msg.c_str();

        m_node.remove_attribute("showInputMessage");
        m_node.append_attribute("showInputMessage") = "1";

        ensureSqrefIsLast(m_node);
    }

    void XLDataValidation::setError(const std::string& title, const std::string& msg, XLDataValidationErrorStyle style)
    {
        if (!m_node) return;

        m_node.remove_attribute("errorStyle");
        switch (style) {
            case XLDataValidationErrorStyle::Stop:
                m_node.append_attribute("errorStyle") = "stop";
                break;
            case XLDataValidationErrorStyle::Warning:
                m_node.append_attribute("errorStyle") = "warning";
                break;
            case XLDataValidationErrorStyle::Information:
                m_node.append_attribute("errorStyle") = "information";
                break;
        }

        m_node.remove_attribute("errorTitle");
        if (!title.empty()) m_node.append_attribute("errorTitle") = title.c_str();

        m_node.remove_attribute("error");
        if (!msg.empty()) m_node.append_attribute("error") = msg.c_str();

        m_node.remove_attribute("showErrorMessage");
        m_node.append_attribute("showErrorMessage") = "1";

        ensureSqrefIsLast(m_node);
    }

    void XLDataValidation::setWholeNumberRange(double min, double max)
    {
        setType(XLDataValidationType::Whole);
        setOperator(XLDataValidationOperator::Between);
        setAllowBlank(true);

        std::ostringstream ss1, ss2;
        ss1 << min;
        ss2 << max;
        setFormula1(ss1.str());
        setFormula2(ss2.str());

        ensureSqrefIsLast(m_node);
    }

    void XLDataValidation::setList(const std::vector<std::string>& items)
    {
        setType(XLDataValidationType::List);
        setAllowBlank(true);

        std::string list;
        for (const auto& item : items) {
            if (!list.empty()) list += ",";
            list += item;
        }
        // Excel expects the list to be double quoted if it's literal items
        setFormula1("\"" + list + "\"");

        ensureSqrefIsLast(m_node);
    }

    // XLDataValidations implementation

    size_t XLDataValidations::count() const
    {
        if (!m_sheetNode) return 0;
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return 0;
        
        size_t count = 0;
        for (auto node : dvNode.children("dataValidation")) {
            (void)node;
            ++count;
        }
        return count;
    }

    XLDataValidation XLDataValidations::append()
    {
        if (!m_sheetNode) return XLDataValidation{};
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) {
            // Correct order in worksheet XML is:
            // ... conditionalFormatting, dataValidations, hyperlinks, printOptions ...
            auto cfNode = m_sheetNode.child("conditionalFormatting");
            if (cfNode) {
                dvNode = m_sheetNode.insert_child_after("dataValidations", cfNode);
            }
            else {
                // If no conditionalFormatting, find a later node to prepend to
                auto nextNode = m_sheetNode.child("hyperlinks");
                if (!nextNode) nextNode = m_sheetNode.child("printOptions");
                if (!nextNode) nextNode = m_sheetNode.child("pageMargins");
                
                if (nextNode) {
                    dvNode = m_sheetNode.insert_child_before("dataValidations", nextNode);
                } else {
                    dvNode = m_sheetNode.append_child("dataValidations");
                }
            }
        }

        auto node = dvNode.append_child("dataValidation");
        
        // Update the count attribute
        size_t currentCount = 0;
        for (auto n : dvNode.children("dataValidation")) {
            (void)n;
            ++currentCount;
        }
        dvNode.remove_attribute("count");
        dvNode.append_attribute("count") = static_cast<unsigned int>(currentCount);

        return XLDataValidation(node);
    }

    XLDataValidation XLDataValidations::at(size_t index)
    {
        if (!m_sheetNode) return XLDataValidation{};
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return XLDataValidation{};

        size_t i = 0;
        for (auto node : dvNode.children("dataValidation")) {
            if (i == index) return XLDataValidation(node);
            ++i;
        }
        return XLDataValidation{};
    }

    void XLDataValidations::clear()
    {
        if (!m_sheetNode) return;
        m_sheetNode.remove_child("dataValidations");
    }

} // namespace OpenXLSX

#include "XLDataValidation.hpp"
#include <algorithm>
#include <sstream>
#include <vector>

namespace OpenXLSX
{
    /**
     * @brief Helper function to ensure attributes are in the correct OOXML order.
     * Openpyxl order (which Excel likes): sqref, showDropDown, showInputMessage, showErrorMessage, allowBlank, errorTitle, error, promptTitle, prompt, type, operator, errorStyle
     */
    void reorderAttributes(XMLNode& node)
    {
        struct Attr {
            std::string name;
            std::string value;
        };
        std::vector<std::string> order = {
            "sqref", "showDropDown", "showInputMessage", "showErrorMessage", "allowBlank", 
            "errorTitle", "error", "promptTitle", "prompt", "type", "operator", "errorStyle", "imeMode"
        };

        // Ensure default boolean flags exist if not present. 
        // Default to "1" (true) for showInputMessage and showErrorMessage to match Excel behavior.
        if (!node.attribute("showDropDown")) node.append_attribute("showDropDown") = "0";
        if (!node.attribute("showInputMessage")) node.append_attribute("showInputMessage") = "1";
        if (!node.attribute("showErrorMessage")) node.append_attribute("showErrorMessage") = "1";
        if (!node.attribute("allowBlank")) node.append_attribute("allowBlank") = "1";

        std::vector<Attr> presentAttrs;
        for (const auto& name : order) {
            auto attr = node.attribute(name.c_str());
            if (attr) {
                presentAttrs.push_back({name, attr.value()});
            }
        }

        // Remove all attributes
        while (node.first_attribute()) {
            node.remove_attribute(node.first_attribute());
        }

        // Re-append in order
        for (const auto& attr : presentAttrs) {
            node.append_attribute(attr.name.c_str()) = attr.value.c_str();
        }
    }

    std::string XLDataValidation::sqref() const
    {
        if (!m_node) return "";
        return m_node.attribute("sqref").value();
    }

    XLDataValidationType XLDataValidation::type() const
    {
        if (!m_node) return XLDataValidationType::None;
        std::string t = m_node.attribute("type").value();
        if (t == "custom") return XLDataValidationType::Custom;
        if (t == "date") return XLDataValidationType::Date;
        if (t == "decimal") return XLDataValidationType::Decimal;
        if (t == "list") return XLDataValidationType::List;
        if (t == "textLength") return XLDataValidationType::TextLength;
        if (t == "time") return XLDataValidationType::Time;
        if (t == "whole") return XLDataValidationType::Whole;
        return XLDataValidationType::None;
    }

    XLDataValidationOperator XLDataValidation::operator_() const
    {
        if (!m_node) return XLDataValidationOperator::Between;
        std::string o = m_node.attribute("operator").value();
        if (o == "equal") return XLDataValidationOperator::Equal;
        if (o == "greaterThan") return XLDataValidationOperator::GreaterThan;
        if (o == "greaterThanOrEqual") return XLDataValidationOperator::GreaterThanOrEqual;
        if (o == "lessThan") return XLDataValidationOperator::LessThan;
        if (o == "lessThanOrEqual") return XLDataValidationOperator::LessThanOrEqual;
        if (o == "notBetween") return XLDataValidationOperator::NotBetween;
        if (o == "notEqual") return XLDataValidationOperator::NotEqual;
        return XLDataValidationOperator::Between;
    }

    bool XLDataValidation::allowBlank() const
    {
        if (!m_node) return false;
        return m_node.attribute("allowBlank").as_bool();
    }

    std::string XLDataValidation::formula1() const
    {
        if (!m_node) return "";
        return m_node.child("formula1").text().get();
    }

    std::string XLDataValidation::formula2() const
    {
        if (!m_node) return "";
        return m_node.child("formula2").text().get();
    }

    bool XLDataValidation::showDropDown() const
    {
        if (!m_node) return false;
        return m_node.attribute("showDropDown").as_bool();
    }

    bool XLDataValidation::showInputMessage() const
    {
        if (!m_node) return false;
        return m_node.attribute("showInputMessage").as_bool();
    }

    bool XLDataValidation::showErrorMessage() const
    {
        if (!m_node) return false;
        return m_node.attribute("showErrorMessage").as_bool();
    }

    XLIMEMode XLDataValidation::imeMode() const
    {
        if (!m_node) return XLIMEMode::NoControl;
        std::string m = m_node.attribute("imeMode").value();
        if (m == "off") return XLIMEMode::Off;
        if (m == "on") return XLIMEMode::On;
        if (m == "disabled") return XLIMEMode::Disabled;
        if (m == "hiragana") return XLIMEMode::Hiragana;
        if (m == "fullKatakana") return XLIMEMode::FullKatakana;
        if (m == "halfKatakana") return XLIMEMode::HalfKatakana;
        if (m == "fullAlpha") return XLIMEMode::FullAlpha;
        if (m == "halfAlpha") return XLIMEMode::HalfAlpha;
        if (m == "fullHangul") return XLIMEMode::FullHangul;
        if (m == "halfHangul") return XLIMEMode::HalfHangul;
        return XLIMEMode::NoControl;
    }

    std::string XLDataValidation::promptTitle() const
    {
        if (!m_node) return "";
        return m_node.attribute("promptTitle").value();
    }

    std::string XLDataValidation::prompt() const
    {
        if (!m_node) return "";
        return m_node.attribute("prompt").value();
    }

    std::string XLDataValidation::errorTitle() const
    {
        if (!m_node) return "";
        return m_node.attribute("errorTitle").value();
    }

    std::string XLDataValidation::error() const
    {
        if (!m_node) return "";
        return m_node.attribute("error").value();
    }

    XLDataValidationErrorStyle XLDataValidation::errorStyle() const
    {
        if (!m_node) return XLDataValidationErrorStyle::Stop;
        std::string s = m_node.attribute("errorStyle").value();
        if (s == "warning") return XLDataValidationErrorStyle::Warning;
        if (s == "information") return XLDataValidationErrorStyle::Information;
        return XLDataValidationErrorStyle::Stop;
    }

    void XLDataValidation::setSqref(const std::string& sqref)
    {
        if (!m_node) return;
        m_node.remove_attribute("sqref");
        m_node.append_attribute("sqref") = sqref.c_str();
        reorderAttributes(m_node);
    }

    void XLDataValidation::setType(XLDataValidationType type)
    {
        if (!m_node) return;
        m_node.remove_attribute("type");
        switch (type) {
            case XLDataValidationType::None:
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
        reorderAttributes(m_node);
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
        reorderAttributes(m_node);
    }

    void XLDataValidation::setAllowBlank(bool allow)
    {
        if (!m_node) return;
        m_node.remove_attribute("allowBlank");
        m_node.append_attribute("allowBlank") = allow ? "1" : "0";
        reorderAttributes(m_node);
    }

    void XLDataValidation::setFormula1(const std::string& formula)
    {
        if (!m_node) return;
        auto fNode = m_node.child("formula1");
        if (!fNode) {
            fNode = m_node.prepend_child("formula1");
        }
        fNode.text().set(formula.c_str());
    }

    void XLDataValidation::setFormula2(const std::string& formula)
    {
        if (!m_node) return;
        auto fNode = m_node.child("formula2");
        if (!fNode) {
            auto f1Node = m_node.child("formula1");
            if (f1Node)
                fNode = m_node.insert_child_after("formula2", f1Node);
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

        reorderAttributes(m_node);
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

        reorderAttributes(m_node);
    }

    void XLDataValidation::setShowDropDown(bool show)
    {
        if (!m_node) return;
        m_node.remove_attribute("showDropDown");
        m_node.append_attribute("showDropDown") = show ? "1" : "0";
        reorderAttributes(m_node);
    }

    void XLDataValidation::setShowInputMessage(bool show)
    {
        if (!m_node) return;
        m_node.remove_attribute("showInputMessage");
        m_node.append_attribute("showInputMessage") = show ? "1" : "0";
        reorderAttributes(m_node);
    }

    void XLDataValidation::setShowErrorMessage(bool show)
    {
        if (!m_node) return;
        m_node.remove_attribute("showErrorMessage");
        m_node.append_attribute("showErrorMessage") = show ? "1" : "0";
        reorderAttributes(m_node);
    }

    void XLDataValidation::setIMEMode(XLIMEMode mode)
    {
        if (!m_node) return;
        m_node.remove_attribute("imeMode");
        switch (mode) {
            case XLIMEMode::NoControl:    m_node.append_attribute("imeMode") = "noControl"; break;
            case XLIMEMode::Off:          m_node.append_attribute("imeMode") = "off"; break;
            case XLIMEMode::On:           m_node.append_attribute("imeMode") = "on"; break;
            case XLIMEMode::Disabled:     m_node.append_attribute("imeMode") = "disabled"; break;
            case XLIMEMode::Hiragana:     m_node.append_attribute("imeMode") = "hiragana"; break;
            case XLIMEMode::FullKatakana: m_node.append_attribute("imeMode") = "fullKatakana"; break;
            case XLIMEMode::HalfKatakana: m_node.append_attribute("imeMode") = "halfKatakana"; break;
            case XLIMEMode::FullAlpha:    m_node.append_attribute("imeMode") = "fullAlpha"; break;
            case XLIMEMode::HalfAlpha:    m_node.append_attribute("imeMode") = "halfAlpha"; break;
            case XLIMEMode::FullHangul:   m_node.append_attribute("imeMode") = "fullHangul"; break;
            case XLIMEMode::HalfHangul:   m_node.append_attribute("imeMode") = "halfHangul"; break;
        }
        reorderAttributes(m_node);
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

        reorderAttributes(m_node);
    }

    void XLDataValidation::setDecimalRange(double min, double max)
    {
        setType(XLDataValidationType::Decimal);
        setOperator(XLDataValidationOperator::Between);
        setAllowBlank(true);

        std::ostringstream ss1, ss2;
        ss1 << min;
        ss2 << max;
        setFormula1(ss1.str());
        setFormula2(ss2.str());

        reorderAttributes(m_node);
    }

    void XLDataValidation::setDateRange(const std::string& min, const std::string& max)
    {
        setType(XLDataValidationType::Date);
        setOperator(XLDataValidationOperator::Between);
        setAllowBlank(true);

        setFormula1(min);
        setFormula2(max);

        reorderAttributes(m_node);
    }

    void XLDataValidation::setTimeRange(const std::string& min, const std::string& max)
    {
        setType(XLDataValidationType::Time);
        setOperator(XLDataValidationOperator::Between);
        setAllowBlank(true);

        setFormula1(min);
        setFormula2(max);

        reorderAttributes(m_node);
    }

    void XLDataValidation::setTextLengthRange(int min, int max)
    {
        setType(XLDataValidationType::TextLength);
        setOperator(XLDataValidationOperator::Between);
        setAllowBlank(true);

        std::ostringstream ss1, ss2;
        ss1 << min;
        ss2 << max;
        setFormula1(ss1.str());
        setFormula2(ss2.str());

        reorderAttributes(m_node);
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

        reorderAttributes(m_node);
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
            // Correct order in worksheet XML is strict.
            XMLNode insertAfter = m_sheetNode.child("conditionalFormatting");
            if (!insertAfter) insertAfter = m_sheetNode.child("phoneticPr");
            if (!insertAfter) insertAfter = m_sheetNode.child("mergeCells");
            if (!insertAfter) insertAfter = m_sheetNode.child("customSheetViews");
            if (!insertAfter) insertAfter = m_sheetNode.child("dataConsolidate");
            if (!insertAfter) insertAfter = m_sheetNode.child("sortState");
            if (!insertAfter) insertAfter = m_sheetNode.child("autoFilter");
            if (!insertAfter) insertAfter = m_sheetNode.child("scenarios");
            if (!insertAfter) insertAfter = m_sheetNode.child("protectedRanges");
            if (!insertAfter) insertAfter = m_sheetNode.child("sheetProtection");
            if (!insertAfter) insertAfter = m_sheetNode.child("sheetCalcPr");
            if (!insertAfter) insertAfter = m_sheetNode.child("sheetData");

            if (insertAfter) {
                dvNode = m_sheetNode.insert_child_after("dataValidations", insertAfter);
            } else {
                dvNode = m_sheetNode.prepend_child("dataValidations");
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

    XLDataValidation XLDataValidations::at(const std::string& sqref)
    {
        if (!m_sheetNode) return XLDataValidation{};
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return XLDataValidation{};

        for (auto node : dvNode.children("dataValidation")) {
            if (std::string(node.attribute("sqref").value()) == sqref)
                return XLDataValidation(node);
        }
        return XLDataValidation{};
    }

    void XLDataValidations::clear()
    {
        if (!m_sheetNode) return;
        m_sheetNode.remove_child("dataValidations");
    }

} // namespace OpenXLSX

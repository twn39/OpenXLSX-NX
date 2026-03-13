#include "XLDataValidation.hpp"
#include <sstream>
#include <vector>
#include <algorithm>

namespace OpenXLSX
{
    /**
     * @brief Helper function to ensure sqref is always the last attribute.
     */
    void ensureSqrefIsLast(XMLNode& node) {
        auto sqrefAttr = node.attribute("sqref");
        if (sqrefAttr) {
            std::string val = sqrefAttr.value();
            node.remove_attribute("sqref");
            node.append_attribute("sqref") = val.c_str();
        }
    }

    void XLDataValidation::setSqref(const std::string& sqref) {
        if (!m_node) return;
        m_node.remove_attribute("sqref");
        m_node.append_attribute("sqref") = sqref.c_str();
    }

    void XLDataValidation::setWholeNumberRange(double min, double max) {
        if (!m_node) return;

        // 1. type
        m_node.remove_attribute("type");
        m_node.append_attribute("type") = "whole";

        // 2. operator
        m_node.remove_attribute("operator");
        m_node.append_attribute("operator") = "between";

        // 3. allowBlank
        m_node.remove_attribute("allowBlank");
        m_node.append_attribute("allowBlank") = "1";

        auto putFormula = [&](const std::string& nodeName, double val) {
            auto fNode = m_node.child(nodeName.c_str());
            if (!fNode) {
                fNode = m_node.append_child(nodeName.c_str());
            }
            std::ostringstream ss;
            ss << val;
            
            if (!fNode.first_child()) {
                fNode.append_child(pugi::node_pcdata).set_value(ss.str().c_str());
            } else {
                 fNode.first_child().set_value(ss.str().c_str());
            }
        };

        putFormula("formula1", min);
        putFormula("formula2", max);
        
        ensureSqrefIsLast(m_node);
    }

    void XLDataValidation::setPrompt(const std::string& title, const std::string& msg) {
        if (!m_node) return;

        m_node.remove_attribute("promptTitle");
        m_node.append_attribute("promptTitle") = title.c_str();

        m_node.remove_attribute("prompt");
        m_node.append_attribute("prompt") = msg.c_str();

        m_node.remove_attribute("showInputMessage");
        m_node.append_attribute("showInputMessage") = "1";
        
        ensureSqrefIsLast(m_node);
    }

    void XLDataValidation::setError(const std::string& title, const std::string& msg) {
        if (!m_node) return;

        m_node.remove_attribute("errorStyle");
        m_node.append_attribute("errorStyle") = "stop";

        m_node.remove_attribute("errorTitle");
        m_node.append_attribute("errorTitle") = title.c_str();

        m_node.remove_attribute("error");
        m_node.append_attribute("error") = msg.c_str();

        m_node.remove_attribute("showErrorMessage");
        m_node.append_attribute("showErrorMessage") = "1";
        
        ensureSqrefIsLast(m_node);
    }
} // namespace OpenXLSX

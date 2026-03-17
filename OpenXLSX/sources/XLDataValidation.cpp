#include "XLDataValidation.hpp"
#include "XLException.hpp"
#include <algorithm>
#include <sstream>
#include <vector>
#include <array>
#include <string_view>

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
        static constexpr std::array<std::string_view, 13> order = {
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
            auto attr = node.attribute(name.data());
            if (attr) {
                presentAttrs.push_back({std::string(name), attr.value()});
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

    XLDataValidationConfig::XLDataValidationConfig(const XLDataValidation& dv)
    {
        type = dv.type();
        operator_ = dv.operator_();
        allowBlank = dv.allowBlank();
        showDropDown = dv.showDropDown();
        showInputMessage = dv.showInputMessage();
        showErrorMessage = dv.showErrorMessage();
        imeMode = dv.imeMode();
        errorStyle = dv.errorStyle();
        promptTitle = dv.promptTitle();
        prompt = dv.prompt();
        errorTitle = dv.errorTitle();
        error = dv.error();
        formula1 = dv.formula1();
        formula2 = dv.formula2();
    }

    XLDataValidationConfig XLDataValidation::config() const
    {
        return XLDataValidationConfig(*this);
    }

    void XLDataValidation::applyConfig(const XLDataValidationConfig& config)
    {
        setType(config.type);
        setOperator(config.operator_);
        setAllowBlank(config.allowBlank);
        setShowDropDown(config.showDropDown);
        setShowInputMessage(config.showInputMessage);
        setShowErrorMessage(config.showErrorMessage);
        if (config.imeMode != XLIMEMode::NoControl) {
            setIMEMode(config.imeMode);
        }
        if (!config.errorTitle.empty() || !config.error.empty()) {
            setError(config.errorTitle, config.error, config.errorStyle);
        } else if (config.errorStyle != XLDataValidationErrorStyle::Stop) {
            setError("", "", config.errorStyle); // At least set the style if customized
        }
        if (!config.promptTitle.empty() || !config.prompt.empty()) {
            setPrompt(config.promptTitle, config.prompt);
        }
        setFormula1(config.formula1);
        setFormula2(config.formula2);
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
        // In OOXML, showDropDown="1" means HIDE.
        // Therefore, if attribute is true ("1"), it means the dropdown is NOT shown.
        // If attribute is missing or false ("0"), the dropdown IS shown.
        return !m_node.attribute("showDropDown").as_bool();
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

    void XLDataValidation::setSqref(std::string_view sqref)
    {
        if (!m_node) return;
        m_node.remove_attribute("sqref");
        m_node.append_attribute("sqref") = std::string(sqref).c_str();
        reorderAttributes(m_node);
    }

    namespace {
        struct XLDataValidationRange {
            uint32_t firstRow;
            uint32_t lastRow;
            uint16_t firstCol;
            uint16_t lastCol;

            bool canMergeWith(const XLDataValidationRange& other) const {
                if (firstRow == other.firstRow && lastRow == other.lastRow) {
                    if (lastCol + 1 >= other.firstCol && firstCol <= other.lastCol + 1) return true;
                }
                if (firstCol == other.firstCol && lastCol == other.lastCol) {
                    if (lastRow + 1 >= other.firstRow && firstRow <= other.lastRow + 1) return true;
                }
                if (firstRow >= other.firstRow && lastRow <= other.lastRow &&
                    firstCol >= other.firstCol && lastCol <= other.lastCol) return true;
                if (other.firstRow >= firstRow && other.lastRow <= lastRow &&
                    other.firstCol >= firstCol && other.lastCol <= lastCol) return true;
                return false;
            }

            void mergeWith(const XLDataValidationRange& other) {
                firstRow = std::min(firstRow, other.firstRow);
                lastRow = std::max(lastRow, other.lastRow);
                firstCol = std::min(firstCol, other.firstCol);
                lastCol = std::max(lastCol, other.lastCol);
            }
            
            std::string toString() const {
                if (firstRow == lastRow && firstCol == lastCol) {
                    return XLCellReference(firstRow, firstCol).address();
                }
                return XLCellReference(firstRow, firstCol).address() + ":" + XLCellReference(lastRow, lastCol).address();
            }
        };

        void collapseRanges(std::vector<XLDataValidationRange>& ranges) {
            bool merged = true;
            while (merged) {
                merged = false;
                for (size_t i = 0; i < ranges.size(); ++i) {
                    for (size_t j = i + 1; j < ranges.size(); ++j) {
                        if (ranges[i].canMergeWith(ranges[j])) {
                            ranges[i].mergeWith(ranges[j]);
                            ranges.erase(ranges.begin() + j);
                            merged = true;
                            break;
                        }
                    }
                    if (merged) break;
                }
            }
        }

        std::vector<XLDataValidationRange> parseSqrefToRanges(std::string_view sqref) {
            std::vector<XLDataValidationRange> result;
            std::string sqrefStr(sqref);
            if (sqrefStr.empty()) return result;
            
            std::stringstream ss(sqrefStr);
            std::string token;
            while (std::getline(ss, token, ' ')) {
                if (token.empty()) continue;
                auto colonPos = token.find(':');
                if (colonPos != std::string::npos) {
                    XLCellReference topLeft(token.substr(0, colonPos));
                    XLCellReference bottomRight(token.substr(colonPos + 1));
                    result.push_back({
                        std::min(topLeft.row(), bottomRight.row()),
                        std::max(topLeft.row(), bottomRight.row()),
                        std::min(topLeft.column(), bottomRight.column()),
                        std::max(topLeft.column(), bottomRight.column())
                    });
                } else {
                    XLCellReference ref(token);
                    result.push_back({ref.row(), ref.row(), ref.column(), ref.column()});
                }
            }
            return result;
        }
    }

    void XLDataValidation::addCell(const XLCellReference& ref)
    {
        addRange(ref, ref);
    }

    void XLDataValidation::addCell(const std::string& ref)
    {
        addCell(XLCellReference(ref));
    }

    void XLDataValidation::addRange(const XLCellReference& topLeft, const XLCellReference& bottomRight)
    {
        std::string currentSqref = sqref();
        auto ranges = parseSqrefToRanges(currentSqref);
        
        uint32_t minRow = std::min(topLeft.row(), bottomRight.row());
        uint32_t maxRow = std::max(topLeft.row(), bottomRight.row());
        uint16_t minCol = std::min(topLeft.column(), bottomRight.column());
        uint16_t maxCol = std::max(topLeft.column(), bottomRight.column());
        
        ranges.push_back({minRow, maxRow, minCol, maxCol});
        collapseRanges(ranges);
        
        std::string newSqref;
        for (size_t i = 0; i < ranges.size(); ++i) {
            if (i > 0) newSqref += " ";
            newSqref += ranges[i].toString();
        }
        setSqref(newSqref);
    }

    void XLDataValidation::addRange(const std::string& range)
    {
        auto colonPos = range.find(':');
        if (colonPos != std::string::npos) {
            XLCellReference topLeft(range.substr(0, colonPos));
            XLCellReference bottomRight(range.substr(colonPos + 1));
            addRange(topLeft, bottomRight);
        } else {
            addCell(range);
        }
    }

    namespace {
        void subtractRange(std::vector<XLDataValidationRange>& ranges, const XLDataValidationRange& toRemove) {
            std::vector<XLDataValidationRange> newRanges;
            for (const auto& r : ranges) {
                // If r and toRemove do not overlap, keep r
                if (r.lastRow < toRemove.firstRow || r.firstRow > toRemove.lastRow ||
                    r.lastCol < toRemove.firstCol || r.firstCol > toRemove.lastCol) {
                    newRanges.push_back(r);
                    continue;
                }

                // They overlap. Split r into up to 4 new non-overlapping rectangles.
                // Top part
                if (r.firstRow < toRemove.firstRow) {
                    newRanges.push_back({r.firstRow, toRemove.firstRow - 1, r.firstCol, r.lastCol});
                }
                // Bottom part
                if (r.lastRow > toRemove.lastRow) {
                    newRanges.push_back({toRemove.lastRow + 1, r.lastRow, r.firstCol, r.lastCol});
                }
                // Left part (constrained vertically by toRemove)
                if (r.firstCol < toRemove.firstCol) {
                    newRanges.push_back({
                        std::max(r.firstRow, toRemove.firstRow),
                        std::min(r.lastRow, toRemove.lastRow),
                        r.firstCol,
                        static_cast<uint16_t>(toRemove.firstCol - 1)
                    });
                }
                // Right part (constrained vertically by toRemove)
                if (r.lastCol > toRemove.lastCol) {
                    newRanges.push_back({
                        std::max(r.firstRow, toRemove.firstRow),
                        std::min(r.lastRow, toRemove.lastRow),
                        static_cast<uint16_t>(toRemove.lastCol + 1),
                        r.lastCol
                    });
                }
            }
            ranges = newRanges;
        }
    }

    void XLDataValidation::removeCell(const XLCellReference& ref)
    {
        removeRange(ref, ref);
    }

    void XLDataValidation::removeCell(const std::string& ref)
    {
        removeCell(XLCellReference(ref));
    }

    void XLDataValidation::removeRange(const XLCellReference& topLeft, const XLCellReference& bottomRight)
    {
        std::string currentSqref = sqref();
        auto ranges = parseSqrefToRanges(currentSqref);
        
        uint32_t minRow = std::min(topLeft.row(), bottomRight.row());
        uint32_t maxRow = std::max(topLeft.row(), bottomRight.row());
        uint16_t minCol = std::min(topLeft.column(), bottomRight.column());
        uint16_t maxCol = std::max(topLeft.column(), bottomRight.column());
        
        subtractRange(ranges, {minRow, maxRow, minCol, maxCol});
        collapseRanges(ranges); // clean up adjacent splits if possible
        
        std::string newSqref;
        for (size_t i = 0; i < ranges.size(); ++i) {
            if (i > 0) newSqref += " ";
            newSqref += ranges[i].toString();
        }
        setSqref(newSqref);
    }

    void XLDataValidation::removeRange(const std::string& range)
    {
        auto colonPos = range.find(':');
        if (colonPos != std::string::npos) {
            XLCellReference topLeft(range.substr(0, colonPos));
            XLCellReference bottomRight(range.substr(colonPos + 1));
            removeRange(topLeft, bottomRight);
        } else {
            removeCell(range);
        }
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

    void XLDataValidation::setFormula1(std::string_view formula)
    {
        if (!m_node) return;
        auto fNode = m_node.child("formula1");
        if (!fNode) {
            fNode = m_node.prepend_child("formula1");
        }
        fNode.text().set(std::string(formula).c_str());
    }

    void XLDataValidation::setFormula2(std::string_view formula)
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
        fNode.text().set(std::string(formula).c_str());
    }

    void XLDataValidation::setPrompt(std::string_view title, std::string_view msg)
    {
        if (!m_node) return;

        m_node.remove_attribute("promptTitle");
        if (!title.empty()) m_node.append_attribute("promptTitle") = std::string(title).c_str();

        m_node.remove_attribute("prompt");
        if (!msg.empty()) m_node.append_attribute("prompt") = std::string(msg).c_str();

        m_node.remove_attribute("showInputMessage");
        m_node.append_attribute("showInputMessage") = "1";

        reorderAttributes(m_node);
    }

    void XLDataValidation::setError(std::string_view title, std::string_view msg, XLDataValidationErrorStyle style)
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
        if (!title.empty()) m_node.append_attribute("errorTitle") = std::string(title).c_str();

        m_node.remove_attribute("error");
        if (!msg.empty()) m_node.append_attribute("error") = std::string(msg).c_str();

        m_node.remove_attribute("showErrorMessage");
        m_node.append_attribute("showErrorMessage") = "1";

        reorderAttributes(m_node);
    }

    void XLDataValidation::setShowDropDown(bool show)
    {
        if (!m_node) return;
        m_node.remove_attribute("showDropDown");
        // In OOXML, showDropDown="1" means HIDE the drop-down. 
        // We want the API 'show' parameter to be intuitive (true = show arrow).
        // Therefore, if show is true, we set to "0" (or omit, but "0" is safe).
        // If show is false, we set to "1" (hide).
        m_node.append_attribute("showDropDown") = show ? "0" : "1";
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

    void XLDataValidation::setDateRange(std::string_view min, std::string_view max)
    {
        setType(XLDataValidationType::Date);
        setOperator(XLDataValidationOperator::Between);
        setAllowBlank(true);

        setFormula1(min);
        setFormula2(max);

        reorderAttributes(m_node);
    }

    void XLDataValidation::setTimeRange(std::string_view min, std::string_view max)
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

        setFormula1(std::to_string(min));
        setFormula2(std::to_string(max));

        reorderAttributes(m_node);
    }

    void XLDataValidation::setList(const std::vector<std::string>& items)
    {
        setType(XLDataValidationType::List);
        setAllowBlank(true);

        if (items.empty()) {
            setFormula1("\"\"");
            reorderAttributes(m_node);
            return;
        }

        std::string list;
        for (const auto& item : items) {
            if (!list.empty()) list += ",";
            
            // Escape double quotes inside items: " becomes ""
            std::string escapedItem = item;
            size_t pos = 0;
            while ((pos = escapedItem.find("\"", pos)) != std::string::npos) {
                escapedItem.replace(pos, 1, "\"\"");
                pos += 2;
            }
            list += escapedItem;
        }
        
        // Excel expects the list to be double quoted if it's literal items
        std::string formula = "\"" + list + "\"";

        // Check the 255 characters limit for Data Validation List (literal)
        if (formula.length() > 255) {
            throw XLException("XLDataValidation::setList: The list length exceeds the 255 character limit. Consider using a cell range reference instead.");
        }

        setFormula1(formula);
        reorderAttributes(m_node);
    }

    void XLDataValidation::setReferenceDropList(std::string_view targetSheet, std::string_view range)
    {
        setType(XLDataValidationType::List);
        setAllowBlank(true);

        std::string sheet(targetSheet);
        std::string formula;

        if (!sheet.empty()) {
            // If sheet name contains space or special characters, it must be quoted in Excel formulas.
            // It's generally safe to always quote it.
            // We also need to escape existing single quotes by doubling them.
            std::string escapedSheet = sheet;
            size_t pos = 0;
            while ((pos = escapedSheet.find("'", pos)) != std::string::npos) {
                escapedSheet.replace(pos, 1, "''");
                pos += 2;
            }
            formula = "='" + escapedSheet + "'!" + std::string(range);
        } else {
            formula = "=" + std::string(range);
        }

        setFormula1(formula);
        reorderAttributes(m_node);
    }

    // XLDataValidations implementation

    size_t XLDataValidations::count() const
    {
        if (!m_sheetNode) return 0;
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return 0;
        
        return dvNode.attribute("count").as_ullong();
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

    XLDataValidation XLDataValidations::addValidation(const XLDataValidationConfig& config, std::string_view sqref)
    {
        auto dv = append();
        dv.setSqref(sqref);
        dv.applyConfig(config);
        return dv;
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

    XLDataValidation XLDataValidations::at(std::string_view sqref)
    {
        if (!m_sheetNode) return XLDataValidation{};
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return XLDataValidation{};

        for (auto node : dvNode.children("dataValidation")) {
            if (std::string_view(node.attribute("sqref").value()) == sqref)
                return XLDataValidation(node);
        }
        return XLDataValidation{};
    }

    void XLDataValidations::clear()
    {
        if (!m_sheetNode) return;
        m_sheetNode.remove_child("dataValidations");
    }

    void XLDataValidations::remove(size_t index)
    {
        if (!m_sheetNode) return;
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return;

        auto child = dvNode.first_child();
        size_t current = 0;
        while (child) {
            if (current == index) {
                dvNode.remove_child(child);
                
                size_t c = dvNode.attribute("count").as_ullong();
                if (c > 0) dvNode.attribute("count") = c - 1;
                
                if (c - 1 == 0) {
                    m_sheetNode.remove_child("dataValidations");
                    // Invalidate m_sheetNode's reference since we removed it? 
                    // No, m_sheetNode points to <worksheet>, not <dataValidations>.
                }
                return;
            }
            child = child.next_sibling("dataValidation");
            current++;
        }
    }

    void XLDataValidations::remove(std::string_view sqref)
    {
        if (!m_sheetNode) return;
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return;

        auto child = dvNode.first_child();
        while (child) {
            auto next = child.next_sibling("dataValidation");
            std::string attrSqref = child.attribute("sqref").value();
            if (attrSqref == sqref) {
                dvNode.remove_child(child);
                
                size_t c = dvNode.attribute("count").as_ullong();
                if (c > 0) dvNode.attribute("count") = c - 1;
            }
            child = next;
        }

        if (dvNode.attribute("count").as_ullong() == 0) {
            m_sheetNode.remove_child("dataValidations");
        }
    }

    bool XLDataValidations::disablePrompts() const
    {
        if (!m_sheetNode) return false;
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return false;
        return dvNode.attribute("disablePrompts").as_bool();
    }

    void XLDataValidations::setDisablePrompts(bool disable)
    {
        if (!m_sheetNode) return;
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return;
        
        dvNode.remove_attribute("disablePrompts");
        if (disable) {
            dvNode.append_attribute("disablePrompts") = "1";
        }
    }

    uint32_t XLDataValidations::xWindow() const
    {
        if (!m_sheetNode) return 0;
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return 0;
        return dvNode.attribute("xWindow").as_uint();
    }

    void XLDataValidations::setXWindow(uint32_t x)
    {
        if (!m_sheetNode) return;
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return;
        dvNode.remove_attribute("xWindow");
        dvNode.append_attribute("xWindow") = x;
    }

    uint32_t XLDataValidations::yWindow() const
    {
        if (!m_sheetNode) return 0;
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return 0;
        return dvNode.attribute("yWindow").as_uint();
    }

    void XLDataValidations::setYWindow(uint32_t y)
    {
        if (!m_sheetNode) return;
        auto dvNode = m_sheetNode.child("dataValidations");
        if (!dvNode) return;
        dvNode.remove_attribute("yWindow");
        dvNode.append_attribute("yWindow") = y;
    }

} // namespace OpenXLSX

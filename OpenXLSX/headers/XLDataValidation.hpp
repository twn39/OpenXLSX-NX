#ifndef OPENXLSX_XLDATAVALIDATION_HPP
#define OPENXLSX_XLDATAVALIDATION_HPP

#include <string>
#include <vector>
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @brief
     */
    enum class XLDataValidationType
    {
        None,
        Custom,
        Date,
        Decimal,
        List,
        TextLength,
        Time,
        Whole
    };

    /**
     * @brief
     */
    enum class XLDataValidationOperator
    {
        Between,
        Equal,
        GreaterThan,
        GreaterThanOrEqual,
        LessThan,
        LessThanOrEqual,
        NotBetween,
        NotEqual
    };

    /**
     * @brief
     */
    enum class XLDataValidationErrorStyle
    {
        Stop,
        Warning,
        Information
    };

    /**
     * @brief IME (Input Method Editor) mode for data validation
     */
    enum class XLIMEMode
    {
        NoControl,
        Off,
        On,
        Disabled,
        Hiragana,
        FullKatakana,
        HalfKatakana,
        FullAlpha,
        HalfAlpha,
        FullHangul,
        HalfHangul
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLDataValidation
    {
    public:
        XLDataValidation() : m_node(nullptr) {}
        explicit XLDataValidation(const XMLNode& node) : m_node(node) {}

        bool empty() const { return !m_node; }

        // Getters
        std::string sqref() const;
        XLDataValidationType type() const;
        XLDataValidationOperator operator_() const;
        bool allowBlank() const;
        std::string formula1() const;
        std::string formula2() const;
        bool showDropDown() const;
        bool showInputMessage() const;
        bool showErrorMessage() const;
        XLIMEMode imeMode() const;
        std::string promptTitle() const;
        std::string prompt() const;
        std::string errorTitle() const;
        std::string error() const;
        XLDataValidationErrorStyle errorStyle() const;

        // Setters
        void setSqref(const std::string& sqref);
        void setType(XLDataValidationType type);
        void setOperator(XLDataValidationOperator op);
        void setAllowBlank(bool allow);
        void setFormula1(const std::string& formula);
        void setFormula2(const std::string& formula);
        void setPrompt(const std::string& title, const std::string& msg);
        void setError(const std::string& title, const std::string& msg, XLDataValidationErrorStyle style = XLDataValidationErrorStyle::Stop);

        // Priority 1 New Properties
        void setShowDropDown(bool show);
        void setShowInputMessage(bool show);
        void setShowErrorMessage(bool show);
        void setIMEMode(XLIMEMode mode);

        // Specific helpers
        void setWholeNumberRange(double min, double max);
        void setDecimalRange(double min, double max);
        void setDateRange(const std::string& min, const std::string& max);
        void setTimeRange(const std::string& min, const std::string& max);
        void setTextLengthRange(int min, int max);
        void setList(const std::vector<std::string>& items);

    private:
        mutable XMLNode m_node;
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLDataValidations
    {
    public:
        XLDataValidations() : m_sheetNode(nullptr) {}
        explicit XLDataValidations(const XMLNode& node) : m_sheetNode(node) {}

        bool empty() const { return !m_sheetNode; }
        size_t count() const;
        XLDataValidation append();
        XLDataValidation at(size_t index);
        XLDataValidation at(const std::string& sqref);
        void clear();

    private:
        mutable XMLNode m_sheetNode;
    };
} // namespace OpenXLSX

#endif // OPENXLSX_XLDATAVALIDATION_HPP

#ifndef OPENXLSX_XLDATAVALIDATION_HPP
#define OPENXLSX_XLDATAVALIDATION_HPP

#include <string>
#include <vector>
#include <string_view>
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"
#include "XLCellReference.hpp"

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

        [[nodiscard]] bool empty() const { return !m_node; }

        // Getters
        [[nodiscard]] std::string sqref() const;
        [[nodiscard]] XLDataValidationType type() const;
        [[nodiscard]] XLDataValidationOperator operator_() const;
        [[nodiscard]] bool allowBlank() const;
        [[nodiscard]] std::string formula1() const;
        [[nodiscard]] std::string formula2() const;
        [[nodiscard]] bool showDropDown() const;
        [[nodiscard]] bool showInputMessage() const;
        [[nodiscard]] bool showErrorMessage() const;
        [[nodiscard]] XLIMEMode imeMode() const;
        [[nodiscard]] std::string promptTitle() const;
        [[nodiscard]] std::string prompt() const;
        [[nodiscard]] std::string errorTitle() const;
        [[nodiscard]] std::string error() const;
        [[nodiscard]] XLDataValidationErrorStyle errorStyle() const;

        // Setters
        void setSqref(std::string_view sqref);
        void addCell(const XLCellReference& ref);
        void addCell(const std::string& ref);
        void addRange(const XLCellReference& topLeft, const XLCellReference& bottomRight);
        void addRange(const std::string& range);
        void setType(XLDataValidationType type);
        void setOperator(XLDataValidationOperator op);
        void setAllowBlank(bool allow);
        void setFormula1(std::string_view formula);
        void setFormula2(std::string_view formula);
        void setPrompt(std::string_view title, std::string_view msg);
        void setError(std::string_view title, std::string_view msg, XLDataValidationErrorStyle style = XLDataValidationErrorStyle::Stop);

        // Priority 1 New Properties
        void setShowDropDown(bool show);
        void setShowInputMessage(bool show);
        void setShowErrorMessage(bool show);
        void setIMEMode(XLIMEMode mode);

        // Specific helpers
        void setWholeNumberRange(double min, double max);
        void setDecimalRange(double min, double max);
        void setDateRange(std::string_view min, std::string_view max);
        void setTimeRange(std::string_view min, std::string_view max);
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

        [[nodiscard]] bool empty() const { return !m_sheetNode; }
        [[nodiscard]] size_t count() const;
        XLDataValidation append();
        [[nodiscard]] XLDataValidation at(size_t index);
        [[nodiscard]] XLDataValidation at(std::string_view sqref);
        void clear();

    private:
        mutable XMLNode m_sheetNode;
    };
} // namespace OpenXLSX

#endif // OPENXLSX_XLDATAVALIDATION_HPP

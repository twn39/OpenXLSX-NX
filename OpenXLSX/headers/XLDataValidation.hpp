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
     * @brief
     */
    class OPENXLSX_EXPORT XLDataValidation
    {
    public:
        XLDataValidation() : m_node(nullptr) {}
        explicit XLDataValidation(const XMLNode& node) : m_node(node) {}

        bool empty() const { return !m_node; }

        void setSqref(const std::string& sqref);
        void setType(XLDataValidationType type);
        void setOperator(XLDataValidationOperator op);
        void setAllowBlank(bool allow);
        void setFormula1(const std::string& formula);
        void setFormula2(const std::string& formula);
        void setPrompt(const std::string& title, const std::string& msg);
        void setError(const std::string& title, const std::string& msg, XLDataValidationErrorStyle style = XLDataValidationErrorStyle::Stop);

        // Specific helpers
        void setWholeNumberRange(double min, double max);
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
        void clear();

    private:
        mutable XMLNode m_sheetNode;
    };
} // namespace OpenXLSX

#endif // OPENXLSX_XLDATAVALIDATION_HPP

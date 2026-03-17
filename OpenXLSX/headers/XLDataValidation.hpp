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

    class XLDataValidation;

    /**
     * @brief
     */
    struct OPENXLSX_EXPORT XLDataValidationConfig
    {
        XLDataValidationType type = XLDataValidationType::None;
        XLDataValidationOperator operator_ = XLDataValidationOperator::Between;
        bool allowBlank = true;
        bool showDropDown = false;
        bool showInputMessage = false;
        bool showErrorMessage = false;
        XLIMEMode imeMode = XLIMEMode::NoControl;
        XLDataValidationErrorStyle errorStyle = XLDataValidationErrorStyle::Stop;
        std::string promptTitle;
        std::string prompt;
        std::string errorTitle;
        std::string error;
        std::string formula1;
        std::string formula2;

        XLDataValidationConfig() = default;
        explicit XLDataValidationConfig(const XLDataValidation& dv);
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

        // Config API
        [[nodiscard]] XLDataValidationConfig config() const;
        void applyConfig(const XLDataValidationConfig& config);

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
        void removeCell(const XLCellReference& ref);
        void removeCell(const std::string& ref);
        void removeRange(const XLCellReference& topLeft, const XLCellReference& bottomRight);
        void removeRange(const std::string& range);
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
        void setReferenceDropList(std::string_view targetSheet, std::string_view range);

    private:
        mutable XMLNode m_node;
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLDataValidations
    {
    public:
        // Iterator class for XLDataValidations
        class Iterator
        {
        public:
            using iterator_category = std::forward_iterator_tag;
            using value_type = XLDataValidation;
            using difference_type = std::ptrdiff_t;
            using pointer = XLDataValidation*;
            using reference = XLDataValidation&;

            Iterator() : m_node(nullptr) {}
            explicit Iterator(XMLNode node) : m_node(node) {}

            value_type operator*() const { return XLDataValidation(m_node); }
            
            class Proxy {
                value_type m_val;
            public:
                explicit Proxy(value_type val) : m_val(val) {}
                value_type* operator->() { return &m_val; }
            };
            
            Proxy operator->() const { return Proxy(XLDataValidation(m_node)); }

            Iterator& operator++() {
                if (m_node) m_node = m_node.next_sibling("dataValidation");
                return *this;
            }
            Iterator operator++(int) {
                Iterator tmp = *this;
                ++(*this);
                return tmp;
            }
            bool operator==(const Iterator& other) const { return m_node == other.m_node; }
            bool operator!=(const Iterator& other) const { return m_node != other.m_node; }

        private:
            XMLNode m_node;
        };

        XLDataValidations() : m_sheetNode(nullptr) {}
        explicit XLDataValidations(const XMLNode& node) : m_sheetNode(node) {}

        [[nodiscard]] bool empty() const { 
            if (!m_sheetNode) return true;
            return !m_sheetNode.child("dataValidations");
        }
        [[nodiscard]] size_t count() const;
        XLDataValidation append();
        XLDataValidation addValidation(const XLDataValidationConfig& config, std::string_view sqref);
        [[nodiscard]] XLDataValidation at(size_t index);
        [[nodiscard]] XLDataValidation at(std::string_view sqref);
        void clear();

        // Priority 1/4 New Properties: Precise Deletion and Iterators
        void remove(size_t index);
        void remove(std::string_view sqref);

        // Global properties
        [[nodiscard]] bool disablePrompts() const;
        void setDisablePrompts(bool disable);
        
        [[nodiscard]] uint32_t xWindow() const;
        void setXWindow(uint32_t x);

        [[nodiscard]] uint32_t yWindow() const;
        void setYWindow(uint32_t y);

        [[nodiscard]] Iterator begin() const {
            if (!m_sheetNode) return Iterator();
            auto dvNode = m_sheetNode.child("dataValidations");
            if (!dvNode) return Iterator();
            return Iterator(dvNode.child("dataValidation"));
        }
        
        [[nodiscard]] Iterator end() const {
            return Iterator();
        }

    private:
        mutable XMLNode m_sheetNode;
    };
} // namespace OpenXLSX

#endif // OPENXLSX_XLDATAVALIDATION_HPP

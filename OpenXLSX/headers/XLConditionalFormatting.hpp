#ifndef OPENXLSX_XLCONDITIONALFORMATTING_HPP
#define OPENXLSX_XLCONDITIONALFORMATTING_HPP

#include <cstdint>
#include <string>
#include <vector>
#include <string_view>
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"
#include "XLStyles.hpp"

namespace OpenXLSX
{
    constexpr const uint16_t XLPriorityNotSet = 0;

    enum class XLCfType : uint8_t {
        Expression        = 0,
        CellIs            = 1,
        ColorScale        = 2,
        DataBar           = 3,
        IconSet           = 4,
        Top10             = 5,
        UniqueValues      = 6,
        DuplicateValues   = 7,
        ContainsText      = 8,
        NotContainsText   = 9,
        BeginsWith        = 10,
        EndsWith          = 11,
        ContainsBlanks    = 12,
        NotContainsBlanks = 13,
        ContainsErrors    = 14,
        NotContainsErrors = 15,
        TimePeriod        = 16,
        AboveAverage      = 17,
        Invalid           = 255
    };

    enum class XLCfOperator : uint8_t {
        LessThan           = 0,
        LessThanOrEqual    = 1,
        Equal              = 2,
        NotEqual           = 3,
        GreaterThanOrEqual = 4,
        GreaterThan        = 5,
        Between            = 6,
        NotBetween         = 7,
        ContainsText       = 8,
        NotContains        = 9,
        BeginsWith         = 10,
        EndsWith           = 11,
        Invalid            = 255
    };

    enum class XLCfTimePeriod : uint8_t {
        Today     = 0,
        Yesterday = 1,
        Tomorrow  = 2,
        Last7Days = 3,
        ThisMonth = 4,
        LastMonth = 5,
        NextMonth = 6,
        ThisWeek  = 7,
        LastWeek  = 8,
        NextWeek  = 9,
        Invalid   = 255
    };

    OPENXLSX_EXPORT XLCfType XLCfTypeFromString(std::string const& typeString);
    OPENXLSX_EXPORT std::string  XLCfTypeToString(XLCfType cfType);
    OPENXLSX_EXPORT XLCfOperator XLCfOperatorFromString(std::string const& operatorString);
    OPENXLSX_EXPORT std::string    XLCfOperatorToString(XLCfOperator cfOperator);
    OPENXLSX_EXPORT XLCfTimePeriod XLCfTimePeriodFromString(std::string const& timePeriodString);
    OPENXLSX_EXPORT std::string XLCfTimePeriodToString(XLCfTimePeriod cfTimePeriod);

    class OPENXLSX_EXPORT XLCfRule
    {
        friend class XLCfRules;
    public:
        XLCfRule();
        explicit XLCfRule(const XMLNode& node);
        XLCfRule(const XLCfRule& other);
        XLCfRule(XLCfRule&& other) noexcept = default;
        ~XLCfRule();

        XLCfRule& operator=(const XLCfRule& other);
        XLCfRule& operator=(XLCfRule&& other) noexcept = default;

        bool empty() const;
        std::string          formula() const;
        XLUnsupportedElement colorScale() const;
        XLUnsupportedElement dataBar() const;
        XLUnsupportedElement iconSet() const;
        XLUnsupportedElement extLst() const;

        XLCfType     type() const;
        XLStyleIndex dxfId() const;
        uint16_t     priority() const;
        bool         stopIfTrue() const;
        bool         aboveAverage() const;
        bool         percent() const;
        bool         bottom() const;
        XLCfOperator Operator() const;
        std::string text() const;
        XLCfTimePeriod timePeriod() const;
        uint16_t rank() const;
        int16_t  stdDev() const;
        bool equalAverage() const;

        bool setFormula(std::string const& newFormula);
        bool setColorScale(XLUnsupportedElement const& newColorScale);
        bool setDataBar(XLUnsupportedElement const& newDataBar);
        bool setIconSet(XLUnsupportedElement const& newIconSet);
        bool setExtLst(XLUnsupportedElement const& newExtLst);

        bool setType(XLCfType newType);
        bool setDxfId(XLStyleIndex newDxfId);

    private:
        bool setPriority(uint16_t newPriority);
    public:
        bool setStopIfTrue(bool set = true);
        bool setAboveAverage(bool set = true);
        bool setPercent(bool set = true);
        bool setBottom(bool set = true);
        bool setOperator(XLCfOperator newOperator);
        bool setText(std::string const& newText);
        bool setTimePeriod(XLCfTimePeriod newTimePeriod);
        bool setRank(uint16_t newRank);
        bool setStdDev(int16_t newStdDev);
        bool setEqualAverage(bool set = true);

        std::string summary() const;

    private:
        mutable XMLNode m_cfRuleNode;
        inline static const std::vector<std::string_view> m_nodeOrder = {
            "formula",
            "colorScale",
            "dataBar",
            "iconSet",
            "extLst"};
    };

    constexpr const char* XLDefaultCfRulePrefix = "\n\t\t";

    class OPENXLSX_EXPORT XLCfRules
    {
    public:
        XLCfRules();
        explicit XLCfRules(const XMLNode& node);
        XLCfRules(const XLCfRules& other);
        XLCfRules(XLCfRules&& other) noexcept = default;
        ~XLCfRules();

        XLCfRules& operator=(const XLCfRules& other);
        XLCfRules& operator=(XLCfRules&& other) noexcept = default;

        bool empty() const;
        uint16_t maxPriorityValue() const;
        bool setPriority(size_t cfRuleIndex, uint16_t newPriority);
        void renumberPriorities(uint16_t increment = 1);
        size_t count() const;
        XLCfRule cfRuleByIndex(size_t index) const;
        XLCfRule operator[](size_t index) const { return cfRuleByIndex(index); }
        size_t create(XLCfRule copyFrom = XLCfRule{}, std::string cfRulePrefix = XLDefaultCfRulePrefix);
        std::string summary() const;

    private:
        mutable XMLNode m_conditionalFormattingNode;
        inline static const std::vector<std::string_view> m_nodeOrder = {
            "cfRule",
            "extLst"};
    };

    class OPENXLSX_EXPORT XLConditionalFormat
    {
        friend class XLConditionalFormats;
    public:
        XLConditionalFormat();
        explicit XLConditionalFormat(const XMLNode& node);
        XLConditionalFormat(const XLConditionalFormat& other);
        XLConditionalFormat(XLConditionalFormat&& other) noexcept = default;
        ~XLConditionalFormat();

        XLConditionalFormat& operator=(const XLConditionalFormat& other);
        XLConditionalFormat& operator=(XLConditionalFormat&& other) noexcept = default;

        bool empty() const;
        std::string sqref() const;
        XLCfRules cfRules() const;
        XLUnsupportedElement extLst() const { return XLUnsupportedElement{}; }

        bool setSqref(std::string newSqref);
        bool setExtLst(XLUnsupportedElement const& newExtLst);
        std::string summary() const;

    private:
        mutable XMLNode m_conditionalFormattingNode;
        inline static const std::vector<std::string_view> m_nodeOrder = {
            "cfRule",
            "extLst"};
    };

    constexpr const char* XLDefaultConditionalFormattingPrefix = "\n\t";

    class OPENXLSX_EXPORT XLConditionalFormats
    {
    public:
        XLConditionalFormats();
        explicit XLConditionalFormats(const XMLNode& node);
        XLConditionalFormats(const XLConditionalFormats& other);
        XLConditionalFormats(XLConditionalFormats&& other) noexcept;
        ~XLConditionalFormats();

        XLConditionalFormats& operator=(const XLConditionalFormats& other);
        XLConditionalFormats& operator=(XLConditionalFormats&& other) noexcept = default;

        bool empty() const;
        size_t count() const;
        XLConditionalFormat conditionalFormatByIndex(size_t index) const;
        XLConditionalFormat operator[](size_t index) const { return conditionalFormatByIndex(index); }
        size_t create(XLConditionalFormat copyFrom = XLConditionalFormat{},
                      std::string conditionalFormattingPrefix = XLDefaultConditionalFormattingPrefix);
        std::string summary() const;

    private:
        mutable XMLNode m_sheetNode;
        const std::vector<std::string_view>* m_nodeOrder;
    };
}

#endif

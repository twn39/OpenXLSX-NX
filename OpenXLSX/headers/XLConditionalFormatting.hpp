/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLCONDITIONALFORMATTING_HPP
#define OPENXLSX_XLCONDITIONALFORMATTING_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLColor.hpp"
#include "XLStyles.hpp"
#include "XLXmlParser.hpp"
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

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

    enum class XLCfvoType : uint8_t { Min = 0, Max = 1, Number = 2, Percent = 3, Formula = 4, Percentile = 5, Invalid = 255 };

    enum class XLDataBarDirection : uint8_t { Context = 0, LeftToRight = 1, RightToLeft = 2 };
    enum class XLDataBarAxisPosition : uint8_t { Automatic = 0, Middle = 1, None = 2 };

    OPENXLSX_EXPORT XLCfType XLCfTypeFromString(std::string const& typeString);
    OPENXLSX_EXPORT std::string  XLCfTypeToString(XLCfType cfType);
    OPENXLSX_EXPORT XLCfOperator XLCfOperatorFromString(std::string const& operatorString);
    OPENXLSX_EXPORT std::string    XLCfOperatorToString(XLCfOperator cfOperator);
    OPENXLSX_EXPORT XLCfTimePeriod XLCfTimePeriodFromString(std::string const& timePeriodString);
    OPENXLSX_EXPORT std::string XLCfTimePeriodToString(XLCfTimePeriod cfTimePeriod);
    OPENXLSX_EXPORT XLCfvoType  XLCfvoTypeFromString(std::string const& cfvoTypeString);
    OPENXLSX_EXPORT std::string XLCfvoTypeToString(XLCfvoType cfvoType);
    OPENXLSX_EXPORT XLDataBarDirection XLDataBarDirectionFromString(std::string const& directionString);
    OPENXLSX_EXPORT std::string        XLDataBarDirectionToString(XLDataBarDirection direction);
    OPENXLSX_EXPORT XLDataBarAxisPosition XLDataBarAxisPositionFromString(std::string const& positionString);
    OPENXLSX_EXPORT std::string           XLDataBarAxisPositionToString(XLDataBarAxisPosition position);

    class OPENXLSX_EXPORT XLCfvo
    {
        friend class XLCfColorScale;
        friend class XLCfDataBar;
        friend class XLCfIconSet;
        friend class XLCfRule;

    public:
        XLCfvo();
        explicit XLCfvo(const XMLNode& node);
        XLCfvo(const XLCfvo& other);
        XLCfvo(XLCfvo&& other) noexcept;
        ~XLCfvo();

        XLCfvo& operator=(const XLCfvo& other);
        XLCfvo& operator=(XLCfvo&& other) noexcept;

        XLCfvoType  type() const;
        std::string value() const;
        bool        gte() const;

        void setType(XLCfvoType type);
        void setValue(const std::string& value);
        void setGte(bool gte);

        XMLNode node() const { return m_cfvoNode; }

    private:
        XMLNode getOrCreateExtNode() const;
        XMLNode getExtNode() const;

        std::unique_ptr<XMLDocument> m_xmlDocument;
        mutable XMLNode              m_cfvoNode;
    };

    class OPENXLSX_EXPORT XLCfColorScale
    {
    public:
        XLCfColorScale();
        explicit XLCfColorScale(const XMLNode& node);
        XLCfColorScale(const XLCfColorScale& other);
        XLCfColorScale(XLCfColorScale&& other) noexcept;
        ~XLCfColorScale();

        XLCfColorScale& operator=(const XLCfColorScale& other);
        XLCfColorScale& operator=(XLCfColorScale&& other) noexcept;

        std::vector<XLCfvo>  cfvos() const;
        std::vector<XLColor> colors() const;

        void addValue(XLCfvoType type, const std::string& value, const XLColor& color);
        void clear();

        XMLNode node() const { return m_colorScaleNode; }

    private:
        XMLNode getOrCreateExtNode() const;
        XMLNode getExtNode() const;

        std::unique_ptr<XMLDocument> m_xmlDocument;
        mutable XMLNode              m_colorScaleNode;
    };

    class OPENXLSX_EXPORT XLCfDataBar
    {
    public:
        XLCfDataBar();
        explicit XLCfDataBar(const XMLNode& node);
        XLCfDataBar(const XLCfDataBar& other);
        XLCfDataBar(XLCfDataBar&& other) noexcept;
        ~XLCfDataBar();

        XLCfDataBar& operator=(const XLCfDataBar& other);
        XLCfDataBar& operator=(XLCfDataBar&& other) noexcept;

        XLCfvo  min() const;
        XLCfvo  max() const;
        XLColor color() const;

        void setMin(XLCfvoType type, const std::string& value);
        void setMax(XLCfvoType type, const std::string& value);
        void setColor(const XLColor& color);

        bool showValue() const;
        void setShowValue(bool show);

        // Extended properties (Excel 2010+)
        uint32_t minLength() const;
        void     setMinLength(uint32_t length);
        uint32_t maxLength() const;
        void     setMaxLength(uint32_t length);
        bool border() const;
        void setBorder(bool border);
        bool gradient() const;
        void setGradient(bool gradient);
        XLDataBarDirection direction() const;
        void               setDirection(XLDataBarDirection direction);
        XLColor borderColor() const;
        void    setBorderColor(const XLColor& color);
        XLColor negativeFillColor() const;
        void    setNegativeFillColor(const XLColor& color);
        XLColor negativeBorderColor() const;
        void    setNegativeBorderColor(const XLColor& color);
        XLDataBarAxisPosition axisPosition() const;
        void                  setAxisPosition(XLDataBarAxisPosition position);
        XLColor axisColor() const;
        void    setAxisColor(const XLColor& color);


        XMLNode node() const { return m_dataBarNode; }

    private:
        XMLNode getOrCreateExtNode() const;
        XMLNode getExtNode() const;

        std::unique_ptr<XMLDocument> m_xmlDocument;
        mutable XMLNode              m_dataBarNode;
    };

    class OPENXLSX_EXPORT XLCfIconSet
    {
    public:
        XLCfIconSet();
        explicit XLCfIconSet(const XMLNode& node);
        XLCfIconSet(const XLCfIconSet& other);
        XLCfIconSet(XLCfIconSet&& other) noexcept;
        ~XLCfIconSet();

        XLCfIconSet& operator=(const XLCfIconSet& other);
        XLCfIconSet& operator=(XLCfIconSet&& other) noexcept;

        std::string iconSet() const;
        void        setIconSet(const std::string& iconSetName);

        std::vector<XLCfvo> cfvos() const;
        void                addValue(XLCfvoType type, const std::string& value);
        void                clear();

        bool showValue() const;
        void setShowValue(bool show);

        // Extended properties (Excel 2010+)
        uint32_t minLength() const;
        void     setMinLength(uint32_t length);
        uint32_t maxLength() const;
        void     setMaxLength(uint32_t length);
        bool border() const;
        void setBorder(bool border);
        bool gradient() const;
        void setGradient(bool gradient);
        XLDataBarDirection direction() const;
        void               setDirection(XLDataBarDirection direction);
        XLColor borderColor() const;
        void    setBorderColor(const XLColor& color);
        XLColor negativeFillColor() const;
        void    setNegativeFillColor(const XLColor& color);
        XLColor negativeBorderColor() const;
        void    setNegativeBorderColor(const XLColor& color);
        XLDataBarAxisPosition axisPosition() const;
        void                  setAxisPosition(XLDataBarAxisPosition position);
        XLColor axisColor() const;
        void    setAxisColor(const XLColor& color);

        bool percent() const;
        void setPercent(bool percent);
        bool reverse() const;
        void setReverse(bool reverse);

        XMLNode node() const { return m_iconSetNode; }

    private:
        XMLNode getOrCreateExtNode() const;
        XMLNode getExtNode() const;

        std::unique_ptr<XMLDocument> m_xmlDocument;
        mutable XMLNode              m_iconSetNode;
    };

    class OPENXLSX_EXPORT XLCfRule
    {
        friend class XLCfRules;

    public:
        XLCfRule();
        explicit XLCfRule(const XMLNode& node);
        XLCfRule(const XLCfRule& other);
        XLCfRule(XLCfRule&& other) noexcept;
        ~XLCfRule();

        XLCfRule& operator=(const XLCfRule& other);
        XLCfRule& operator=(XLCfRule&& other) noexcept;

        bool                     empty() const;
        std::string              formula() const;
        std::vector<std::string> formulas() const;
        XLCfColorScale           colorScale() const;
        XLCfDataBar              dataBar() const;
        XLCfIconSet              iconSet() const;
        XLUnsupportedElement     extLst() const;

        XLCfType       type() const;
        XLStyleIndex   dxfId() const;
        uint16_t       priority() const;
        bool           stopIfTrue() const;
        bool           aboveAverage() const;
        bool           percent() const;
        bool           bottom() const;
        XLCfOperator   Operator() const;
        std::string    text() const;
        XLCfTimePeriod timePeriod() const;
        uint16_t       rank() const;
        int16_t        stdDev() const;
        bool           equalAverage() const;

        XLCfRule& setFormula(std::string const& newFormula);
        bool      addFormula(std::string const& newFormula);
        void      clearFormulas();
        bool      setColorScale(XLCfColorScale const& newColorScale);
        bool      setDataBar(XLCfDataBar const& newDataBar);
        bool      setIconSet(XLCfIconSet const& newIconSet);
        bool      setExtLst(XLUnsupportedElement const& newExtLst);

        XLCfRule& setType(XLCfType newType);
        XLCfRule& setDxfId(XLStyleIndex newDxfId);

        XLCfRule& setPriority(uint16_t newPriority);

    public:
        XLCfRule& setStopIfTrue(bool set = true);
        XLCfRule& setAboveAverage(bool set = true);
        XLCfRule& setPercent(bool set = true);
        XLCfRule& setBottom(bool set = true);
        XLCfRule& setOperator(XLCfOperator newOperator);
        XLCfRule& setText(std::string const& newText);
        XLCfRule& setTimePeriod(XLCfTimePeriod newTimePeriod);
        XLCfRule& setRank(uint16_t newRank);
        XLCfRule& setStdDev(int16_t newStdDev);
        XLCfRule& setEqualAverage(bool set = true);

        std::string summary() const;

        XMLNode node() const { return m_cfRuleNode; }

    private:
        std::unique_ptr<XMLDocument>                      m_xmlDocument;
        mutable XMLNode                                   m_cfRuleNode;
        inline static const std::vector<std::string_view> m_nodeOrder = {"formula", "colorScale", "dataBar", "iconSet", "extLst"};
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

        bool        empty() const;
        uint16_t    maxPriorityValue() const;
        bool        setPriority(size_t cfRuleIndex, uint16_t newPriority);
        void        renumberPriorities(uint16_t increment = 1);
        size_t      count() const;
        XLCfRule    cfRuleByIndex(size_t index) const;
        XLCfRule    operator[](size_t index) const { return cfRuleByIndex(index); }
        size_t      create(XLCfRule copyFrom = XLCfRule{}, std::string cfRulePrefix = XLDefaultCfRulePrefix);
        std::string summary() const;

    private:
        mutable XMLNode                                   m_conditionalFormattingNode;
        inline static const std::vector<std::string_view> m_nodeOrder = {"cfRule", "extLst"};
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

        bool                 empty() const;
        std::string          sqref() const;
        XLCfRules            cfRules() const;
        XLUnsupportedElement extLst() const { return XLUnsupportedElement{}; }

        bool        setSqref(std::string newSqref);
        bool        setExtLst(XLUnsupportedElement const& newExtLst);
        std::string summary() const;

    private:
        mutable XMLNode                                   m_conditionalFormattingNode;
        inline static const std::vector<std::string_view> m_nodeOrder = {"cfRule", "extLst"};
    };

    constexpr const char* XLDefaultConditionalFormattingPrefix = "\n\t";

    class OPENXLSX_EXPORT XLConditionalFormats
    {
    public:
        XLConditionalFormats();
        explicit XLConditionalFormats(const XMLNode& sheet);
        XLConditionalFormats(const XLConditionalFormats& other);
        XLConditionalFormats(XLConditionalFormats&& other) noexcept;
        ~XLConditionalFormats();

        XLConditionalFormats& operator=(const XLConditionalFormats& other);
        XLConditionalFormats& operator=(XLConditionalFormats&& other) noexcept = default;

        bool                empty() const;
        size_t              count() const;
        XLConditionalFormat conditionalFormatByIndex(size_t index) const;
        XLConditionalFormat operator[](size_t index) const { return conditionalFormatByIndex(index); }
        size_t              create(XLConditionalFormat copyFrom                    = XLConditionalFormat{},
                                   std::string         conditionalFormattingPrefix = XLDefaultConditionalFormattingPrefix);
        std::string         summary() const;

    private:
        mutable XMLNode                      m_sheetNode;
        const std::vector<std::string_view>* m_nodeOrder;
    };

    // ----- Helper Builder Functions for XLCfRule -----

    OPENXLSX_EXPORT XLCfRule XLColorScaleRule(const XLColor& minColor, const XLColor& maxColor);
    OPENXLSX_EXPORT XLCfRule XLColorScaleRule(const XLColor& minColor, const XLColor& midColor, const XLColor& maxColor);
    OPENXLSX_EXPORT XLCfRule XLDataBarRule(const XLColor& color, bool showValue = true);
    OPENXLSX_EXPORT XLCfRule XLCellIsRule(XLCfOperator op, const std::string& value);
    OPENXLSX_EXPORT XLCfRule XLCellIsRule(const std::string& op, const std::string& value);
    OPENXLSX_EXPORT XLCfRule XLFormulaRule(const std::string& formula);

    // Advanced builders
    OPENXLSX_EXPORT XLCfRule XLIconSetRule(const std::string& iconSetName = "3TrafficLights1", bool showValue = true, bool reverse = false);
    OPENXLSX_EXPORT XLCfRule XLTop10Rule(uint16_t rank = 10, bool percent = false, bool bottom = false);
    OPENXLSX_EXPORT XLCfRule XLAboveAverageRule(bool aboveAverage = true, bool equalAverage = false, int16_t stdDev = 0);
    OPENXLSX_EXPORT XLCfRule XLDuplicateValuesRule(bool unique = false);
    OPENXLSX_EXPORT XLCfRule XLContainsTextRule(const std::string& text);
    OPENXLSX_EXPORT XLCfRule XLNotContainsTextRule(const std::string& text);
    OPENXLSX_EXPORT XLCfRule XLContainsBlanksRule();
    OPENXLSX_EXPORT XLCfRule XLNotContainsBlanksRule();
    OPENXLSX_EXPORT XLCfRule XLContainsErrorsRule();
    OPENXLSX_EXPORT XLCfRule XLNotContainsErrorsRule();

}    // namespace OpenXLSX

#endif

#include "XLConditionalFormatting.hpp"
#include <string>

namespace OpenXLSX
{
    /**
     * @brief get the correct XLCfType from the OOXML cfRule type attribute string
     * @param typeString the string as used in the OOXML
     * @return the corresponding XLCfType enum value
     */
    XLCfType XLCfTypeFromString(std::string const& typeString)
    {
        if (typeString == "expression") return XLCfType::Expression;
        if (typeString == "cellIs") return XLCfType::CellIs;
        if (typeString == "colorScale") return XLCfType::ColorScale;
        if (typeString == "dataBar") return XLCfType::DataBar;
        if (typeString == "iconSet") return XLCfType::IconSet;
        if (typeString == "top10") return XLCfType::Top10;
        if (typeString == "uniqueValues") return XLCfType::UniqueValues;
        if (typeString == "duplicateValues") return XLCfType::DuplicateValues;
        if (typeString == "containsText") return XLCfType::ContainsText;
        if (typeString == "notContainsText") return XLCfType::NotContainsText;
        if (typeString == "beginsWith") return XLCfType::BeginsWith;
        if (typeString == "endsWith") return XLCfType::EndsWith;
        if (typeString == "containsBlanks") return XLCfType::ContainsBlanks;
        if (typeString == "notContainsBlanks") return XLCfType::NotContainsBlanks;
        if (typeString == "containsErrors") return XLCfType::ContainsErrors;
        if (typeString == "notContainsErrors") return XLCfType::NotContainsErrors;
        if (typeString == "timePeriod") return XLCfType::TimePeriod;
        if (typeString == "aboveAverage") return XLCfType::AboveAverage;
        return XLCfType::Invalid;
    }

    /**
     * @brief inverse of XLCfTypeFromString
     * @param cfType the type for which to get the OOXML string
     */
    std::string XLCfTypeToString(XLCfType cfType)
    {
        switch (cfType) {
            case XLCfType::Expression:
                return "expression";
            case XLCfType::CellIs:
                return "cellIs";
            case XLCfType::ColorScale:
                return "colorScale";
            case XLCfType::DataBar:
                return "dataBar";
            case XLCfType::IconSet:
                return "iconSet";
            case XLCfType::Top10:
                return "top10";
            case XLCfType::UniqueValues:
                return "uniqueValues";
            case XLCfType::DuplicateValues:
                return "duplicateValues";
            case XLCfType::ContainsText:
                return "containsText";
            case XLCfType::NotContainsText:
                return "notContainsText";
            case XLCfType::BeginsWith:
                return "beginsWith";
            case XLCfType::EndsWith:
                return "endsWith";
            case XLCfType::ContainsBlanks:
                return "containsBlanks";
            case XLCfType::NotContainsBlanks:
                return "notContainsBlanks";
            case XLCfType::ContainsErrors:
                return "containsErrors";
            case XLCfType::NotContainsErrors:
                return "notContainsErrors";
            case XLCfType::TimePeriod:
                return "timePeriod";
            case XLCfType::AboveAverage:
                return "aboveAverage";
            case XLCfType::Invalid:
                [[fallthrough]];
            default:
                return "(invalid)";
        }
    }

    /**
     * @brief get the correct XLCfOperator from the OOXML cfRule operator attribute string
     * @param operatorString the string as used in the OOXML
     * @return the corresponding XLCfOperator enum value
     */
    XLCfOperator XLCfOperatorFromString(std::string const& operatorString)
    {
        if (operatorString == "lessThan") return XLCfOperator::LessThan;
        if (operatorString == "lessThanOrEqual") return XLCfOperator::LessThanOrEqual;
        if (operatorString == "equal") return XLCfOperator::Equal;
        if (operatorString == "notEqual") return XLCfOperator::NotEqual;
        if (operatorString == "greaterThanOrEqual") return XLCfOperator::GreaterThanOrEqual;
        if (operatorString == "greaterThan") return XLCfOperator::GreaterThan;
        if (operatorString == "between") return XLCfOperator::Between;
        if (operatorString == "notBetween") return XLCfOperator::NotBetween;
        if (operatorString == "containsText") return XLCfOperator::ContainsText;
        if (operatorString == "notContains") return XLCfOperator::NotContains;
        if (operatorString == "beginsWith") return XLCfOperator::BeginsWith;
        if (operatorString == "endsWith") return XLCfOperator::EndsWith;
        return XLCfOperator::Invalid;
    }

    /**
     * @brief inverse of XLCfOperatorFromString
     * @param cfOperator the XLCfOperator for which to get the OOXML string
     */
    std::string XLCfOperatorToString(XLCfOperator cfOperator)
    {
        switch (cfOperator) {
            case XLCfOperator::LessThan:
                return "lessThan";
            case XLCfOperator::LessThanOrEqual:
                return "lessThanOrEqual";
            case XLCfOperator::Equal:
                return "equal";
            case XLCfOperator::NotEqual:
                return "notEqual";
            case XLCfOperator::GreaterThanOrEqual:
                return "greaterThanOrEqual";
            case XLCfOperator::GreaterThan:
                return "greaterThan";
            case XLCfOperator::Between:
                return "between";
            case XLCfOperator::NotBetween:
                return "notBetween";
            case XLCfOperator::ContainsText:
                return "containsText";
            case XLCfOperator::NotContains:
                return "notContains";
            case XLCfOperator::BeginsWith:
                return "beginsWith";
            case XLCfOperator::EndsWith:
                return "endsWith";
            case XLCfOperator::Invalid:
                [[fallthrough]];
            default:
                return "(invalid)";
        }
    }

    /**
     * @brief get the correct XLCfTimePeriod from the OOXML cfRule timePeriod attribute string
     * @param timePeriodString the string as used in the OOXML
     * @return the corresponding XLCfTimePeriod enum value
     */
    XLCfTimePeriod XLCfTimePeriodFromString(std::string const& timePeriodString)
    {
        if (timePeriodString == "today") return XLCfTimePeriod::Today;
        if (timePeriodString == "yesterday") return XLCfTimePeriod::Yesterday;
        if (timePeriodString == "tomorrow") return XLCfTimePeriod::Tomorrow;
        if (timePeriodString == "last7Days") return XLCfTimePeriod::Last7Days;
        if (timePeriodString == "thisMonth") return XLCfTimePeriod::ThisMonth;
        if (timePeriodString == "lastMonth") return XLCfTimePeriod::LastMonth;
        if (timePeriodString == "nextMonth") return XLCfTimePeriod::NextMonth;
        if (timePeriodString == "thisWeek") return XLCfTimePeriod::ThisWeek;
        if (timePeriodString == "lastWeek") return XLCfTimePeriod::LastWeek;
        if (timePeriodString == "nextWeek") return XLCfTimePeriod::NextWeek;
        return XLCfTimePeriod::Invalid;
    }

    /**
     * @brief inverse of XLCfTimePeriodFromString
     * @param cfTimePeriod the XLCfTimePeriod for which to get the OOXML string
     */
    std::string XLCfTimePeriodToString(XLCfTimePeriod cfTimePeriod)
    {
        switch (cfTimePeriod) {
            case XLCfTimePeriod::Today:
                return "today";
            case XLCfTimePeriod::Yesterday:
                return "yesterday";
            case XLCfTimePeriod::Tomorrow:
                return "tomorrow";
            case XLCfTimePeriod::Last7Days:
                return "last7Days";
            case XLCfTimePeriod::ThisMonth:
                return "thisMonth";
            case XLCfTimePeriod::LastMonth:
                return "lastMonth";
            case XLCfTimePeriod::NextMonth:
                return "nextMonth";
            case XLCfTimePeriod::ThisWeek:
                return "thisWeek";
            case XLCfTimePeriod::LastWeek:
                return "lastWeek";
            case XLCfTimePeriod::NextWeek:
                return "nextWeek";
            case XLCfTimePeriod::Invalid:
                [[fallthrough]];
            default:
                return "(invalid)";
        }
    }

    XLCfvoType XLCfvoTypeFromString(std::string const& cfvoTypeString)
    {
        if (cfvoTypeString == "min") return XLCfvoType::Min;
        if (cfvoTypeString == "max") return XLCfvoType::Max;
        if (cfvoTypeString == "num") return XLCfvoType::Number;
        if (cfvoTypeString == "percent") return XLCfvoType::Percent;
        if (cfvoTypeString == "formula") return XLCfvoType::Formula;
        if (cfvoTypeString == "percentile") return XLCfvoType::Percentile;
        return XLCfvoType::Invalid;
    }

    std::string XLCfvoTypeToString(XLCfvoType cfvoType)
    {
        switch (cfvoType) {
            case XLCfvoType::Min:
                return "min";
            case XLCfvoType::Max:
                return "max";
            case XLCfvoType::Number:
                return "num";
            case XLCfvoType::Percent:
                return "percent";
            case XLCfvoType::Formula:
                return "formula";
            case XLCfvoType::Percentile:
                return "percentile";
            default:
                return "num";
        }
    }

    XLDataBarDirection XLDataBarDirectionFromString(std::string const& directionString)
    {
        if (directionString == "leftToRight") return XLDataBarDirection::LeftToRight;
        if (directionString == "rightToLeft") return XLDataBarDirection::RightToLeft;
        return XLDataBarDirection::Context;
    }

    std::string XLDataBarDirectionToString(XLDataBarDirection direction)
    {
        switch (direction) {
            case XLDataBarDirection::LeftToRight: return "leftToRight";
            case XLDataBarDirection::RightToLeft: return "rightToLeft";
            default: return "context";
        }
    }

    XLDataBarAxisPosition XLDataBarAxisPositionFromString(std::string const& positionString)
    {
        if (positionString == "middle") return XLDataBarAxisPosition::Middle;
        if (positionString == "none") return XLDataBarAxisPosition::None;
        return XLDataBarAxisPosition::Automatic;
    }

    std::string XLDataBarAxisPositionToString(XLDataBarAxisPosition position)
    {
        switch (position) {
            case XLDataBarAxisPosition::Middle: return "middle";
            case XLDataBarAxisPosition::None: return "none";
            default: return "automatic";
        }
    }
}    // namespace OpenXLSX

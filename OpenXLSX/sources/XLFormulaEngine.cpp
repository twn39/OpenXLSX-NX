// ===== External Includes ===== //
#include <algorithm>
#include <cassert>
#include <cctype>
#include <cmath>
#include <ctime>
#include <fmt/format.h>
#include <functional>
#include <numeric>
#include <string>
#include <unordered_map>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLCellReference.hpp"
#include "XLDateTime.hpp"
#include "XLException.hpp"
#include "XLFormulaEngine.hpp"
#include "XLWorksheet.hpp"

using namespace OpenXLSX;

// =============================================================================
#include "XLFormulaUtils.hpp"

// =============================================================================
// Lexer
// =============================================================================

XLFormulaArg XLFormulaEngine::expandArg(const XLASTNode& argNode, const XLCellResolver& resolver) const
{
    if (argNode.kind == XLNodeKind::Range) return expandRange(argNode.text, resolver);

    // Single cell reference: create a 1x1 LazyRange so position metadata (row/col) is preserved.
    // Functions like ROW() and COLUMN() need the address, not just the resolved value.
    // Other functions that iterate values will call operator[0] which resolves via the resolver.
    if (argNode.kind == XLNodeKind::CellRef) return expandRange(argNode.text, resolver);

    // Evaluate normally and wrap in a single-element scalar
    return XLFormulaArg(evalNode(argNode, resolver));
}

// =============================================================================
// Evaluator – evalNode
// =============================================================================

XLCellValue XLFormulaEngine::evalNode(const XLASTNode& node, const XLCellResolver& resolver) const
{
    switch (node.kind) {
        case XLNodeKind::Number:
            return XLCellValue(node.number);
        case XLNodeKind::StringLit:
            return XLCellValue(node.text);
        case XLNodeKind::BoolLit:
            return XLCellValue(node.boolean);
        case XLNodeKind::ErrorLit: {
            XLCellValue e;
            e.setError(node.text);
            return e;
        }

        case XLNodeKind::CellRef: {
            if (!resolver) return XLCellValue{};
            return resolver(node.text);
        }

        case XLNodeKind::Range: {
            // Range used as scalar = first cell value
            auto vals = expandRange(node.text, resolver);
            return vals.empty() ? XLCellValue{} : vals[0];
        }

        case XLNodeKind::UnaryOp: {
            Expects(node.children.size() == 1);
            auto val = evalNode(*node.children[0], resolver);
            if (node.op == XLTokenKind::Minus) {
                if (!isNumeric(val)) return errValue();
                double d = toDouble(val);
                if (val.type() == XLValueType::Integer) return XLCellValue(static_cast<int64_t>(-d));
                return XLCellValue(-d);
            }
            if (node.op == XLTokenKind::Percent) {
                if (!isNumeric(val)) return errValue();
                return XLCellValue(toDouble(val) / 100.0);
            }
            return val;
        }

        case XLNodeKind::BinOp: {
            Expects(node.children.size() == 2);

            // String concat – evaluate early, no numeric coercion
            if (node.op == XLTokenKind::Amp) {
                auto lv = evalNode(*node.children[0], resolver);
                auto rv = evalNode(*node.children[1], resolver);
                if (isError(lv)) return lv;
                if (isError(rv)) return rv;
                return XLCellValue(toString(lv) + toString(rv));
            }

            auto lv = evalNode(*node.children[0], resolver);
            auto rv = evalNode(*node.children[1], resolver);
            if (isError(lv)) return lv;
            if (isError(rv)) return rv;

            // Arithmetic operators
            if (node.op == XLTokenKind::Plus || node.op == XLTokenKind::Minus || node.op == XLTokenKind::Star ||
                node.op == XLTokenKind::Slash || node.op == XLTokenKind::Caret)
            {
                if (!isNumeric(lv) || !isNumeric(rv)) return errValue();
                double l = toDouble(lv), r = toDouble(rv);
                switch (node.op) {
                    case XLTokenKind::Plus:
                        return XLCellValue(l + r);
                    case XLTokenKind::Minus:
                        return XLCellValue(l - r);
                    case XLTokenKind::Star:
                        return XLCellValue(l * r);
                    case XLTokenKind::Slash:
                        if (r == 0.0) return errDiv0();
                        return XLCellValue(l / r);
                    case XLTokenKind::Caret:
                        return XLCellValue(std::pow(l, r));
                    default:
                        break;
                }
            }

            // Comparison operators
            {
                bool result = false;
                // Numeric comparison
                if (isNumeric(lv) && isNumeric(rv)) {
                    double l = toDouble(lv), r = toDouble(rv);
                    switch (node.op) {
                        case XLTokenKind::Eq:
                            result = (l == r);
                            break;
                        case XLTokenKind::NEq:
                            result = (l != r);
                            break;
                        case XLTokenKind::Lt:
                            result = (l < r);
                            break;
                        case XLTokenKind::Le:
                            result = (l <= r);
                            break;
                        case XLTokenKind::Gt:
                            result = (l > r);
                            break;
                        case XLTokenKind::Ge:
                            result = (l >= r);
                            break;
                        default:
                            return errValue();
                    }
                }
                else {
                    // String comparison (case-insensitive like Excel)
                    std::string ls = toString(lv), rs = toString(rv);
                    std::transform(ls.begin(), ls.end(), ls.begin(), ::tolower);
                    std::transform(rs.begin(), rs.end(), rs.begin(), ::tolower);
                    switch (node.op) {
                        case XLTokenKind::Eq:
                            result = (ls == rs);
                            break;
                        case XLTokenKind::NEq:
                            result = (ls != rs);
                            break;
                        case XLTokenKind::Lt:
                            result = (ls < rs);
                            break;
                        case XLTokenKind::Le:
                            result = (ls <= rs);
                            break;
                        case XLTokenKind::Gt:
                            result = (ls > rs);
                            break;
                        case XLTokenKind::Ge:
                            result = (ls >= rs);
                            break;
                        default:
                            return errValue();
                    }
                }
                return XLCellValue(result);
            }
        }

        case XLNodeKind::FuncCall: {
            const auto& funcs = getBuiltins();
            auto        it    = funcs.find(node.text);
            if (it == funcs.end()) return errName();

            // Build per-arg vectors (ranges are expanded, scalars wrapped)
            std::vector<XLFormulaArg> argVecs;
            argVecs.reserve(node.children.size());
            for (const auto& child : node.children) argVecs.push_back(expandArg(*child, resolver));

            try {
                return it->second(argVecs);
            }
            catch (const std::exception& ex) {
                XLCellValue e;
                e.setError(std::string("#ERROR: ") + ex.what());
                return e;
            }
        }

        default:
            return errValue();
    }
}

// =============================================================================
// Evaluator – public evaluate()
// =============================================================================

XLCellValue XLFormulaEngine::evaluate(std::string_view formula, const XLCellResolver& resolver) const
{
    if (formula.empty()) return XLCellValue{};
    try {
        auto tokens = XLFormulaLexer::tokenize(formula);
        auto ast    = XLFormulaParser::parse(gsl::span<const XLToken>(tokens));
        return evalNode(*ast, resolver);
    }
    catch (const XLException&) {
        throw;
    }
    catch (const std::exception& ex) {
        XLCellValue e;
        e.setError(std::string("#ERROR: ") + ex.what());
        return e;
    }
}

// =============================================================================
// makeResolver
// =============================================================================

XLCellResolver XLFormulaEngine::makeResolver(const XLWorksheet& wks)
{
    return [&wks](std::string_view ref) -> XLCellValue {
        try {
            // Strip sheet prefix (e.g. "Sheet1!A1" -> "A1")
            auto        bang     = ref.find('!');
            std::string cellAddr = std::string(bang != std::string_view::npos ? ref.substr(bang + 1) : ref);
            return wks.cell(cellAddr).value();
        }
        catch (...) {
            return XLCellValue{};
        }
    };
}

// =============================================================================
// Built-in function registrations
// =============================================================================

XLFormulaEngine::XLFormulaEngine() = default;

const std::unordered_map<std::string, XLFormulaEngine::FuncImpl>& XLFormulaEngine::getBuiltins()
{
    static const std::unordered_map<std::string, FuncImpl> builtins = []() {
        std::unordered_map<std::string, FuncImpl> map;
        map["SUM"]         = fnSum;
        map["AVERAGE"]     = fnAverage;
        map["AVG"]         = fnAverage;    // alias
        map["MIN"]         = fnMin;
        map["MAX"]         = fnMax;
        map["COUNT"]       = fnCount;
        map["COUNTA"]      = fnCounta;
        map["IF"]          = fnIf;
        map["IFS"]         = fnIfs;
        map["SWITCH"]      = fnSwitch;
        map["AND"]         = fnAnd;
        map["OR"]          = fnOr;
        map["NOT"]         = fnNot;
        map["IFERROR"]     = fnIferror;
        map["ABS"]         = fnAbs;
        map["ROUND"]       = fnRound;
        map["ROUNDUP"]     = fnRoundup;
        map["ROUNDDOWN"]   = fnRounddown;
        map["SQRT"]        = fnSqrt;
        map["PI"]          = fnPi;
        map["SIN"]         = fnSin;
        map["COS"]         = fnCos;
        map["TAN"]         = fnTan;
        map["ASIN"]        = fnAsin;
        map["ACOS"]        = fnAcos;
        map["DEGREES"]     = fnDegrees;
        map["RADIANS"]     = fnRadians;
        map["RAND"]        = fnRand;
        map["RANDBETWEEN"] = fnRandbetween;
        map["INT"]         = fnInt;
        map["MOD"]         = fnMod;
        map["POWER"]       = fnPower;
        map["VLOOKUP"]     = fnVlookup;
        map["HLOOKUP"]     = fnHlookup;
        map["XLOOKUP"]     = fnXlookup;
        map["INDEX"]       = fnIndex;
        map["MATCH"]       = fnMatch;
        map["CONCATENATE"] = fnConcatenate;
        map["CONCAT"]      = fnConcatenate;    // alias
        map["LEN"]         = fnLen;
        map["LEFT"]        = fnLeft;
        map["RIGHT"]       = fnRight;
        map["MID"]         = fnMid;
        map["UPPER"]       = fnUpper;
        map["LOWER"]       = fnLower;
        map["TRIM"]        = fnTrim;
        map["TEXT"]        = fnText;
        map["ISNUMBER"]    = fnIsnumber;
        map["ISBLANK"]     = fnIsblank;
        map["ISERROR"]     = fnIserror;
        map["ISTEXT"]      = fnIstext;

        // ---- Date / Time ----
        map["TODAY"]       = fnToday;
        map["NOW"]         = fnNow;
        map["DATE"]        = fnDate;
        map["TIME"]        = fnTime;
        map["YEAR"]        = fnYear;
        map["MONTH"]       = fnMonth;
        map["DAY"]         = fnDay;
        map["HOUR"]        = fnHour;
        map["MINUTE"]      = fnMinute;
        map["SECOND"]      = fnSecond;
        map["DAYS"]        = fnDays;
        map["_xlfn.DAYS"]  = fnDays;
        map["WEEKDAY"]     = fnWeekday;
        map["EDATE"]       = fnEdate;
        map["EOMONTH"]     = fnEomonth;
        map["WORKDAY"]     = fnWorkday;
        map["NETWORKDAYS"] = fnNetworkdays;

        // ---- Financial ----
        map["PMT"] = fnPmt;
        map["FV"]  = fnFv;
        map["PV"]  = fnPv;
        map["NPV"] = fnNpv;

        // ---- Math extended ----
        map["SUMPRODUCT"] = fnSumproduct;
        map["CEILING"]    = fnCeil;
        map["CEIL"]       = fnCeil;
        map["FLOOR"]      = fnFloor;
        map["LOG"]        = fnLog;
        map["LOG10"]      = fnLog10;
        map["EXP"]        = fnExp;
        map["SIGN"]       = fnSign;

        // ---- Text extended ----
        map["FIND"]           = fnFind;
        map["SEARCH"]         = fnSearch;
        map["SUBSTITUTE"]     = fnSubstitute;
        map["REPLACE"]        = fnReplace;
        map["REPT"]           = fnRept;
        map["EXACT"]          = fnExact;
        map["T"]              = fnT;
        map["VALUE"]          = fnValue;
        map["TEXTJOIN"]       = fnTextjoin;
        map["_xlfn.TEXTJOIN"] = fnTextjoin;
        map["CLEAN"]          = fnClean;
        map["PROPER"]         = fnProper;

        // ---- Statistical / Conditional ----
        map["SUMIF"]        = fnSumif;
        map["COUNTIF"]      = fnCountif;
        map["SUMIFS"]       = fnSumifs;
        map["COUNTIFS"]     = fnCountifs;
        map["MAXIFS"]       = fnMaxifs;
        map["_xlfn.MAXIFS"] = fnMaxifs;
        map["MINIFS"]       = fnMinifs;
        map["_xlfn.MINIFS"] = fnMinifs;
        map["AVERAGEIF"]    = fnAverageif;
        map["RANK"]         = fnRank;
        map["RANK.EQ"]      = fnRank;
        map["LARGE"]        = fnLarge;
        map["SMALL"]        = fnSmall;
        map["STDEV"]        = fnStdev;
        map["STDEV.S"]      = fnStdev;
        map["VAR"]          = fnVar;
        map["VAR.S"]        = fnVar;
        map["MEDIAN"]       = fnMedian;
        map["COUNTBLANK"]   = fnCountblank;

        // ---- Info extended ----
        map["ISNA"]      = fnIsna;
        map["IFNA"]      = fnIfna;
        map["ISLOGICAL"] = fnIslogical;
        map["ISNONTEXT"] = fnIsnontext;

        // ---- Easy Additions ----
        map["TRUE"]               = fnTrue;
        map["FALSE"]              = fnFalse;
        map["ISEVEN"]             = fnIseven;
        map["ISODD"]              = fnIsodd;
        map["MROUND"]             = fnMround;
        map["CEILING.MATH"]       = fnCeilingMath;
        map["_xlfn.CEILING.MATH"] = fnCeilingMath;
        map["FLOOR.MATH"]         = fnFloorMath;
        map["_xlfn.FLOOR.MATH"]   = fnFloorMath;
        map["VAR.P"]              = fnVarp;
        map["_xlfn.VAR.P"]        = fnVarp;
        map["VARP"]               = fnVarp;
        map["STDEV.P"]            = fnStdevp;
        map["_xlfn.STDEV.P"]      = fnStdevp;
        map["STDEVP"]             = fnStdevp;
        map["VARA"]               = fnVara;
        map["VARPA"]              = fnVarpa;
        map["STDEVA"]             = fnStdeva;
        map["STDEVPA"]            = fnStdevpa;
        map["PERMUT"]             = fnPermut;
        map["PERMUTATIONA"]       = fnPermutationa;
        map["_xlfn.PERMUTATIONA"] = fnPermutationa;
        map["FISHER"]             = fnFisher;
        map["FISHERINV"]          = fnFisherinv;
        map["STANDARDIZE"]        = fnStandardize;
        map["PEARSON"]            = fnPearson;
        map["CORREL"]             = fnPearson;
        map["COVAR"]              = fnCovarianceP;
        map["COVARIANCE.P"]       = fnCovarianceP;
        map["COVARIANCE.S"]       = fnCovarianceS;
        map["PERCENTILE"]         = fnPercentileInc;
        map["PERCENTILE.INC"]     = fnPercentileInc;
        map["PERCENTILE.EXC"]     = fnPercentileExc;
        map["QUARTILE"]           = fnQuartileInc;
        map["QUARTILE.INC"]       = fnQuartileInc;
        map["QUARTILE.EXC"]       = fnQuartileExc;
        map["TRIMMEAN"]           = fnTrimmean;
        map["SLOPE"]              = fnSlope;
        map["INTERCEPT"]          = fnIntercept;
        map["RSQ"]                = fnRsq;
        map["AVERAGEIFS"]         = fnAverageifs;
        map["ISOWEEKNUM"]         = fnIsoweeknum;
        map["_xlfn.ISOWEEKNUM"]   = fnIsoweeknum;
        map["WEEKNUM"]            = fnWeeknum;
        map["DAYS360"]            = fnDays360;
        map["NPER"]               = fnNper;
        map["DB"]                 = fnDb;
        map["DDB"]                = fnDdb;

        // Fix mapped functions from earlier additions
        map["ISERR"]    = fnIserr;
        map["TRUNC"]    = fnTrunc;
        map["SUMSQ"]    = fnSumsq;
        map["SUMX2MY2"] = fnSumx2my2;
        map["SUMX2PY2"] = fnSumx2py2;
        map["SUMXMY2"]  = fnSumxmy2;
        map["AVEDEV"]   = fnAvedev;
        map["DEVSQ"]    = fnDevsq;
        map["AVERAGEA"] = fnAveragea;

        // Pre-existing checklist functions (from original batch)
        map["SLN"]     = fnSln;
        map["SYD"]     = fnSyd;
        map["CHAR"]    = fnChar;
        map["UNICHAR"] = fnUnichar;
        map["CODE"]    = fnCode;
        map["UNICODE"] = fnUnicode;

        // ---- Batch 1: Math / Trig additions (pure cmath wrappers) ----
        map["LN"]          = fnLn;
        map["ATAN"]        = fnAtan;
        map["ATAN2"]       = fnAtan2Fn;
        map["SINH"]        = fnSinh;
        map["COSH"]        = fnCosh;
        map["TANH"]        = fnTanh;
        map["ASINH"]       = fnAsinh;
        map["ACOSH"]       = fnAcosh;
        map["ATANH"]       = fnAtanh;
        map["SQRTPI"]      = fnSqrtpi;
        map["FACT"]        = fnFact;
        map["FACTDOUBLE"]  = fnFactdouble;
        map["COMBIN"]      = fnCombin;
        map["COMBINA"]     = fnCombina;
        map["_xlfn.COMBINA"] = fnCombina;
        map["PRODUCT"]     = fnProduct;
        map["GCD"]         = fnGcd;
        map["LCM"]         = fnLcm;
        map["EVEN"]        = fnEven;
        map["ODD"]         = fnOdd;
        map["QUOTIENT"]    = fnQuotient;

        // ---- Batch 2: High-frequency utility ----
        map["SUBTOTAL"]               = fnSubtotal;
        map["CHOOSE"]                 = fnChoose;
        map["ROW"]                    = fnRow;
        map["COLUMN"]                 = fnColumn;
        map["DATEDIF"]                = fnDatedif;
        map["IRR"]                    = fnIrr;
        map["MIRR"]                   = fnMirr;
        map["RATE"]                   = fnRate;
        map["IPMT"]                   = fnIpmt;
        map["PPMT"]                   = fnPpmt;
        map["MODE"]                   = fnModeSngl;
        map["MODE.SNGL"]              = fnModeSngl;
        map["_xlfn.MODE.SNGL"]        = fnModeSngl;
        map["SKEW"]                   = fnSkew;
        map["KURT"]                   = fnKurt;
        map["FORECAST"]               = fnForecastLinear;
        map["FORECAST.LINEAR"]        = fnForecastLinear;
        map["_xlfn.FORECAST.LINEAR"]  = fnForecastLinear;

        // ---- Batch 3: Statistical distributions ----
        // Normal distribution
        map["NORM.S.DIST"]        = fnNormSDist;
        map["_xlfn.NORM.S.DIST"]  = fnNormSDist;
        map["NORMSDIST"]          = fnNormSDist;    // legacy Excel alias (2 args → cumul=true)
        map["NORM.DIST"]          = fnNormDist;
        map["_xlfn.NORM.DIST"]    = fnNormDist;
        map["NORMDIST"]           = fnNormDist;
        map["NORM.S.INV"]         = fnNormSInv;
        map["_xlfn.NORM.S.INV"]   = fnNormSInv;
        map["NORMSINV"]           = fnNormSInv;
        map["NORM.INV"]           = fnNormInv;
        map["_xlfn.NORM.INV"]     = fnNormInv;
        map["NORMINV"]            = fnNormInv;

        // T distribution
        map["T.DIST"]             = fnTDist;
        map["_xlfn.T.DIST"]       = fnTDist;
        map["T.DIST.RT"]          = fnTDistRT;
        map["_xlfn.T.DIST.RT"]    = fnTDistRT;
        map["T.DIST.2T"]          = fnTDist2T;
        map["_xlfn.T.DIST.2T"]    = fnTDist2T;
        map["TDIST"]              = fnTDist2T;      // legacy: 2-tail by convention
        map["T.INV"]              = fnTInv;
        map["_xlfn.T.INV"]        = fnTInv;
        map["T.INV.2T"]           = fnTInv2T;
        map["_xlfn.T.INV.2T"]     = fnTInv2T;
        map["TINV"]               = fnTInv2T;       // legacy alias

        // Chi-squared distribution
        map["CHISQ.DIST"]         = fnChisqDist;
        map["_xlfn.CHISQ.DIST"]   = fnChisqDist;
        map["CHISQ.DIST.RT"]      = fnChisqDistRT;
        map["_xlfn.CHISQ.DIST.RT"]= fnChisqDistRT;
        map["CHIDIST"]            = fnChisqDistRT;  // legacy: right-tail
        map["CHISQ.INV"]          = fnChisqInv;
        map["_xlfn.CHISQ.INV"]    = fnChisqInv;
        map["CHISQ.INV.RT"]       = fnChisqInvRT;
        map["_xlfn.CHISQ.INV.RT"] = fnChisqInvRT;
        map["CHIINV"]             = fnChisqInvRT;   // legacy alias

        // Binomial distribution
        map["BINOM.DIST"]         = fnBinomDist;
        map["_xlfn.BINOM.DIST"]   = fnBinomDist;
        map["BINOMDIST"]          = fnBinomDist;

        // Poisson distribution
        map["POISSON.DIST"]       = fnPoissonDist;
        map["_xlfn.POISSON.DIST"] = fnPoissonDist;
        map["POISSON"]            = fnPoissonDist;

        // Exponential distribution
        map["EXPON.DIST"]         = fnExponDist;
        map["_xlfn.EXPON.DIST"]   = fnExponDist;
        map["EXPONDIST"]          = fnExponDist;

        return map;
    }();
    return builtins;
}

// =============================================================================
// Built-in: Math / Statistical
// =============================================================================

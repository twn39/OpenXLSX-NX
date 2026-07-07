#ifndef OPENXLSX_XLFORMULAFUNCTIONSINTERNAL_HPP
#define OPENXLSX_XLFORMULAFUNCTIONSINTERNAL_HPP

#include <vector>
#include "XLCellValue.hpp"
#include "XLFormulaEngine.hpp"

namespace OpenXLSX::Formula
{
    // ---- Built-in functions ----
    XLCellValue fnSum(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAverage(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMin(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMax(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCount(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCounta(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIf(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIfs(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSwitch(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAnd(const std::vector<XLFormulaArg>& args);
    XLCellValue fnOr(const std::vector<XLFormulaArg>& args);
    XLCellValue fnNot(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIferror(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAbs(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRound(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRoundup(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRounddown(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSqrt(const std::vector<XLFormulaArg>& args);
    XLCellValue fnPi(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSin(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCos(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTan(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAsin(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAcos(const std::vector<XLFormulaArg>& args);
    XLCellValue fnDegrees(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRadians(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRand(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRandbetween(const std::vector<XLFormulaArg>& args);
    XLCellValue fnInt(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMod(const std::vector<XLFormulaArg>& args);
    XLCellValue fnPower(const std::vector<XLFormulaArg>& args);
    XLCellValue fnVlookup(const std::vector<XLFormulaArg>& args);
    XLCellValue fnHlookup(const std::vector<XLFormulaArg>& args);
    XLCellValue fnXlookup(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIndex(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMatch(const std::vector<XLFormulaArg>& args);
    XLCellValue fnConcatenate(const std::vector<XLFormulaArg>& args);
    XLCellValue fnLen(const std::vector<XLFormulaArg>& args);
    XLCellValue fnLeft(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRight(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMid(const std::vector<XLFormulaArg>& args);
    XLCellValue fnUpper(const std::vector<XLFormulaArg>& args);
    XLCellValue fnLower(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTrim(const std::vector<XLFormulaArg>& args);
    XLCellValue fnText(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIsnumber(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIsblank(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIserror(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIstext(const std::vector<XLFormulaArg>& args);

    // ---- Date / Time ----
    XLCellValue fnToday(const std::vector<XLFormulaArg>& args);
    XLCellValue fnNow(const std::vector<XLFormulaArg>& args);
    XLCellValue fnDate(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTime(const std::vector<XLFormulaArg>& args);
    XLCellValue fnYear(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMonth(const std::vector<XLFormulaArg>& args);
    XLCellValue fnDay(const std::vector<XLFormulaArg>& args);
    XLCellValue fnHour(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMinute(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSecond(const std::vector<XLFormulaArg>& args);
    XLCellValue fnDays(const std::vector<XLFormulaArg>& args);
    XLCellValue fnWeekday(const std::vector<XLFormulaArg>& args);
    XLCellValue fnEdate(const std::vector<XLFormulaArg>& args);
    XLCellValue fnEomonth(const std::vector<XLFormulaArg>& args);
    XLCellValue fnWorkday(const std::vector<XLFormulaArg>& args);
    XLCellValue fnNetworkdays(const std::vector<XLFormulaArg>& args);

    // ---- Financial ----
    XLCellValue fnPmt(const std::vector<XLFormulaArg>& args);
    XLCellValue fnFv(const std::vector<XLFormulaArg>& args);
    XLCellValue fnPv(const std::vector<XLFormulaArg>& args);
    XLCellValue fnNpv(const std::vector<XLFormulaArg>& args);

    // ---- Math extended ----
    XLCellValue fnSumproduct(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCeil(const std::vector<XLFormulaArg>& args);
    XLCellValue fnFloor(const std::vector<XLFormulaArg>& args);
    XLCellValue fnLog(const std::vector<XLFormulaArg>& args);
    XLCellValue fnLog10(const std::vector<XLFormulaArg>& args);
    XLCellValue fnExp(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSign(const std::vector<XLFormulaArg>& args);

    // ---- Text extended ----
    XLCellValue fnFind(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSearch(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSubstitute(const std::vector<XLFormulaArg>& args);
    XLCellValue fnReplace(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRept(const std::vector<XLFormulaArg>& args);
    XLCellValue fnExact(const std::vector<XLFormulaArg>& args);
    XLCellValue fnT(const std::vector<XLFormulaArg>& args);
    XLCellValue fnValue(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTextjoin(const std::vector<XLFormulaArg>& args);
    XLCellValue fnClean(const std::vector<XLFormulaArg>& args);
    XLCellValue fnProper(const std::vector<XLFormulaArg>& args);

    // ---- Statistical / Conditional ----
    XLCellValue fnSumif(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCountif(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSumifs(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCountifs(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMaxifs(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMinifs(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAverageif(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRank(const std::vector<XLFormulaArg>& args);
    XLCellValue fnLarge(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSmall(const std::vector<XLFormulaArg>& args);
    XLCellValue fnStdev(const std::vector<XLFormulaArg>& args);
    XLCellValue fnVar(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMedian(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCountblank(const std::vector<XLFormulaArg>& args);

    // ---- Info extended ----
    XLCellValue fnIsna(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIfna(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIslogical(const std::vector<XLFormulaArg>& args);

    // ---- Easy Additions ----
    XLCellValue fnTrue(const std::vector<XLFormulaArg>& args);
    XLCellValue fnFalse(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIseven(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIsodd(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMround(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCeilingMath(const std::vector<XLFormulaArg>& args);
    XLCellValue fnFloorMath(const std::vector<XLFormulaArg>& args);
    XLCellValue fnVarp(const std::vector<XLFormulaArg>& args);
    XLCellValue fnStdevp(const std::vector<XLFormulaArg>& args);
    XLCellValue fnVara(const std::vector<XLFormulaArg>& args);
    XLCellValue fnVarpa(const std::vector<XLFormulaArg>& args);
    XLCellValue fnStdeva(const std::vector<XLFormulaArg>& args);
    XLCellValue fnStdevpa(const std::vector<XLFormulaArg>& args);
    XLCellValue fnPermut(const std::vector<XLFormulaArg>& args);
    XLCellValue fnPermutationa(const std::vector<XLFormulaArg>& args);
    XLCellValue fnFisher(const std::vector<XLFormulaArg>& args);
    XLCellValue fnFisherinv(const std::vector<XLFormulaArg>& args);
    XLCellValue fnStandardize(const std::vector<XLFormulaArg>& args);
    XLCellValue fnPearson(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRsq(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAverageifs(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIsoweeknum(const std::vector<XLFormulaArg>& args);
    XLCellValue fnWeeknum(const std::vector<XLFormulaArg>& args);
    XLCellValue fnDays360(const std::vector<XLFormulaArg>& args);
    XLCellValue fnNper(const std::vector<XLFormulaArg>& args);
    XLCellValue fnDb(const std::vector<XLFormulaArg>& args);
    XLCellValue fnDdb(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIserr(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTrunc(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSumsq(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSumx2my2(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSumx2py2(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSumxmy2(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAvedev(const std::vector<XLFormulaArg>& args);
    XLCellValue fnDevsq(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAveragea(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSln(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSyd(const std::vector<XLFormulaArg>& args);
    XLCellValue fnChar(const std::vector<XLFormulaArg>& args);
    XLCellValue fnUnichar(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCode(const std::vector<XLFormulaArg>& args);
    XLCellValue fnUnicode(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCovarianceP(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCovarianceS(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTrimmean(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSlope(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIntercept(const std::vector<XLFormulaArg>& args);
    XLCellValue fnPercentileInc(const std::vector<XLFormulaArg>& args);
    XLCellValue fnPercentileExc(const std::vector<XLFormulaArg>& args);
    XLCellValue fnQuartileInc(const std::vector<XLFormulaArg>& args);
    XLCellValue fnQuartileExc(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIsnontext(const std::vector<XLFormulaArg>& args);

    // ---- Batch 2 – Math / Trig (pure cmath wrappers) ----
    XLCellValue fnLn(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAtan(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAtan2Fn(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSinh(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCosh(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTanh(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAsinh(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAcosh(const std::vector<XLFormulaArg>& args);
    XLCellValue fnAtanh(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSqrtpi(const std::vector<XLFormulaArg>& args);
    XLCellValue fnFact(const std::vector<XLFormulaArg>& args);
    XLCellValue fnFactdouble(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCombin(const std::vector<XLFormulaArg>& args);
    XLCellValue fnCombina(const std::vector<XLFormulaArg>& args);
    XLCellValue fnProduct(const std::vector<XLFormulaArg>& args);
    XLCellValue fnGcd(const std::vector<XLFormulaArg>& args);
    XLCellValue fnLcm(const std::vector<XLFormulaArg>& args);
    XLCellValue fnEven(const std::vector<XLFormulaArg>& args);
    XLCellValue fnOdd(const std::vector<XLFormulaArg>& args);
    XLCellValue fnQuotient(const std::vector<XLFormulaArg>& args);

    // ---- Batch 2 – High-frequency utility ----
    XLCellValue fnSubtotal(const std::vector<XLFormulaArg>& args);
    XLCellValue fnChoose(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRow(const std::vector<XLFormulaArg>& args);
    XLCellValue fnColumn(const std::vector<XLFormulaArg>& args);
    XLCellValue fnDatedif(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIrr(const std::vector<XLFormulaArg>& args);
    XLCellValue fnMirr(const std::vector<XLFormulaArg>& args);
    XLCellValue fnRate(const std::vector<XLFormulaArg>& args);
    XLCellValue fnIpmt(const std::vector<XLFormulaArg>& args);
    XLCellValue fnPpmt(const std::vector<XLFormulaArg>& args);
    XLCellValue fnModeSngl(const std::vector<XLFormulaArg>& args);
    XLCellValue fnSkew(const std::vector<XLFormulaArg>& args);
    XLCellValue fnKurt(const std::vector<XLFormulaArg>& args);
    XLCellValue fnForecastLinear(const std::vector<XLFormulaArg>& args);

    // ---- Batch 3 – Statistical distributions ----
    XLCellValue fnNormSDist(const std::vector<XLFormulaArg>& args);
    XLCellValue fnNormDist(const std::vector<XLFormulaArg>& args);
    XLCellValue fnNormSInv(const std::vector<XLFormulaArg>& args);
    XLCellValue fnNormInv(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTDist(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTDistRT(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTDist2T(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTInv(const std::vector<XLFormulaArg>& args);
    XLCellValue fnTInv2T(const std::vector<XLFormulaArg>& args);
    XLCellValue fnChisqDist(const std::vector<XLFormulaArg>& args);
    XLCellValue fnChisqDistRT(const std::vector<XLFormulaArg>& args);
    XLCellValue fnChisqInv(const std::vector<XLFormulaArg>& args);
    XLCellValue fnChisqInvRT(const std::vector<XLFormulaArg>& args);
    XLCellValue fnBinomDist(const std::vector<XLFormulaArg>& args);
    XLCellValue fnPoissonDist(const std::vector<XLFormulaArg>& args);
    XLCellValue fnExponDist(const std::vector<XLFormulaArg>& args);
}

#endif // OPENXLSX_XLFORMULAFUNCTIONSINTERNAL_HPP

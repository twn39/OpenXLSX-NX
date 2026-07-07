#include "XLFormulaFunctionsInternal.hpp"
#include "XLFormulaRegistry.hpp"
#include "XLDateTime.hpp"
#include "XLNumberFormatter.hpp"
#include "XLFormulaUtils.hpp"
#include <algorithm>
#include <array>
#include <cmath>
#include <numeric>
#include <random>
#include <regex>
#include <unordered_set>
#include <cctype>

using namespace OpenXLSX;


XLCellValue Formula::fnPmt(const std::vector<XLFormulaArg>& args)
{
    // PMT(rate, nper, pv, [fv=0], [type=0])
    // Payment = (pv*(1+r)^n*r + fv*r) / ((1+r*type)*((1+r)^n - 1))   when r≠0
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double r    = toDouble(args[0][0]);
    double n    = toDouble(args[1][0]);
    double pv   = toDouble(args[2][0]);
    double fv   = (args.size() > 3 && !args[3].empty()) ? toDouble(args[3][0]) : 0.0;
    int    type = (args.size() > 4 && !args[4].empty()) ? static_cast<int>(toDouble(args[4][0])) : 0;

    if (r == 0.0) return XLCellValue(-(pv + fv) / n);
    double q   = std::pow(1.0 + r, n);
    double pmt = -(pv * q + fv) * r / ((1.0 + r * type) * (q - 1.0));
    return XLCellValue(pmt);
}

XLCellValue Formula::fnFv(const std::vector<XLFormulaArg>& args)
{
    // FV(rate, nper, pmt, [pv=0], [type=0])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double r    = toDouble(args[0][0]);
    double n    = toDouble(args[1][0]);
    double pmt  = toDouble(args[2][0]);
    double pv   = (args.size() > 3 && !args[3].empty()) ? toDouble(args[3][0]) : 0.0;
    int    type = (args.size() > 4 && !args[4].empty()) ? static_cast<int>(toDouble(args[4][0])) : 0;

    if (r == 0.0) return XLCellValue(-(pv + pmt * n));
    double q  = std::pow(1.0 + r, n);
    double fv = -(pv * q + pmt * (1.0 + r * type) * (q - 1.0) / r);
    return XLCellValue(fv);
}

XLCellValue Formula::fnPv(const std::vector<XLFormulaArg>& args)
{
    // PV(rate, nper, pmt, [fv=0], [type=0])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double r    = toDouble(args[0][0]);
    double n    = toDouble(args[1][0]);
    double pmt  = toDouble(args[2][0]);
    double fv   = (args.size() > 3 && !args[3].empty()) ? toDouble(args[3][0]) : 0.0;
    int    type = (args.size() > 4 && !args[4].empty()) ? static_cast<int>(toDouble(args[4][0])) : 0;

    if (r == 0.0) return XLCellValue(-(fv + pmt * n));
    double q  = std::pow(1.0 + r, n);
    double pv = -(fv + pmt * (1.0 + r * type) * (q - 1.0) / r) / q;
    return XLCellValue(pv);
}

XLCellValue Formula::fnNpv(const std::vector<XLFormulaArg>& args)
{
    // NPV(rate, value1, value2, ...) – all values flattened, 1-based periods
    if (args.size() < 2 || args[0].empty()) return errValue();
    double r      = toDouble(args[0][0]);
    double npv    = 0.0;
    int    period = 1;
    for (std::size_t i = 1; i < args.size(); ++i)
        for (const auto& v : args[i])
            if (isNumeric(v)) { npv += toDouble(v) / std::pow(1.0 + r, period++); }
    return XLCellValue(npv);
}

XLCellValue Formula::fnNper(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() < 3) return errValue();
    double rate = nums[0];
    double pmt  = nums[1];
    double pv   = nums[2];
    double fv   = (nums.size() > 3) ? nums[3] : 0.0;
    double type = (nums.size() > 4) ? nums[4] : 0.0;

    if (rate == 0) {
        if (pmt == 0) return errNum();
        return XLCellValue(-(pv + fv) / pmt);
    }

    double num = pmt * (1 + rate * type) - fv * rate;
    double den = pv * rate + pmt * (1 + rate * type);

    double ratio = num / den;
    if (ratio <= 0.0) return errNum();

    double res = std::log(ratio) / std::log(1 + rate);
    return XLCellValue(res);
}

XLCellValue Formula::fnDb(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() < 4) return errValue();
    double cost    = nums[0];
    double salvage = nums[1];
    double life    = nums[2];
    double period  = nums[3];
    double month   = (nums.size() > 4) ? nums[4] : 12.0;

    if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || month <= 0) return errNum();
    if (period > life && (period != life + 1 || month == 12)) return errNum();

    double rate = 1.0 - std::pow(salvage / cost, 1.0 / life);
    rate        = std::round(rate * 1000.0) / 1000.0;    // DB rounds to 3 decimal places

    double total = 0.0;
    double dep   = cost * rate * month / 12.0;
    total += dep;
    if (period == 1) return XLCellValue(dep);

    for (int i = 2; i < period; ++i) {
        dep = (cost - total) * rate;
        total += dep;
    }

    if (period == life + 1) { dep = (cost - total) * rate * (12.0 - month) / 12.0; }
    else {
        dep = (cost - total) * rate;
    }
    return XLCellValue(dep);
}

XLCellValue Formula::fnDdb(const std::vector<XLFormulaArg>& args)
{
    auto nums = numerics(args);
    if (nums.size() < 4) return errValue();
    double cost    = nums[0];
    double salvage = nums[1];
    double life    = nums[2];
    double period  = nums[3];
    double factor  = (nums.size() > 4) ? nums[4] : 2.0;

    if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || factor <= 0) return errNum();

    double total = 0.0;
    double dep   = 0.0;
    for (int i = 1; i <= period; ++i) {
        dep = (cost - total) * (factor / life);
        if (total + dep > cost - salvage) { dep = cost - salvage - total; }
        total += dep;
    }
    return XLCellValue(dep);
}

XLCellValue Formula::fnSln(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double cost    = toDouble(args[0][0]);
    double salvage = toDouble(args[1][0]);
    double life    = toDouble(args[2][0]);
    if (std::isnan(cost) || std::isnan(salvage) || std::isnan(life) || life == 0.0) return errDiv0();
    return XLCellValue((cost - salvage) / life);
}

XLCellValue Formula::fnSyd(const std::vector<XLFormulaArg>& args)
{
    if (args.size() < 4 || args[0].empty() || args[1].empty() || args[2].empty() || args[3].empty()) return errValue();
    double cost    = toDouble(args[0][0]);
    double salvage = toDouble(args[1][0]);
    double life    = toDouble(args[2][0]);
    double per     = toDouble(args[3][0]);
    if (std::isnan(cost) || std::isnan(salvage) || std::isnan(life) || std::isnan(per) || life == 0.0) return errDiv0();
    if (per < 1.0 || per > life) return errNum();
    return XLCellValue(((cost - salvage) * (life - per + 1.0) * 2.0) / (life * (life + 1.0)));
}

XLCellValue Formula::fnIrr(const std::vector<XLFormulaArg>& args)
{
    // IRR(values, [guess=0.1])
    if (args.empty() || args[0].empty()) return errValue();
    auto cashflows = numerics(args[0]);
    if (cashflows.empty()) return errNum();

    double rate = (args.size() > 1 && !args[1].empty()) ? toDouble(args[1][0]) : 0.1;

    // Newton-Raphson iteration
    for (int iter = 0; iter < 128; ++iter) {
        double npv  = 0.0;
        double dnpv = 0.0;
        for (std::size_t i = 0; i < cashflows.size(); ++i) {
            double denom = std::pow(1.0 + rate, static_cast<double>(i));
            npv  += cashflows[i] / denom;
            dnpv -= static_cast<double>(i) * cashflows[i] / (denom * (1.0 + rate));
        }
        if (std::abs(dnpv) < 1e-15) return errNum();
        double newRate = rate - npv / dnpv;
        if (std::abs(newRate - rate) < 1e-10) {
            if (newRate <= -1.0) return errNum();
            return XLCellValue(newRate);
        }
        rate = newRate;
        if (rate <= -1.0) rate = -0.9999;  // clamp to valid domain
    }
    return errNum();  // Did not converge
}

XLCellValue Formula::fnMirr(const std::vector<XLFormulaArg>& args)
{
    // MIRR(values, finance_rate, reinvest_rate)
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    auto   cashflows     = numerics(args[0]);
    double financeRate   = toDouble(args[1][0]);
    double reinvestRate  = toDouble(args[2][0]);
    int    n             = static_cast<int>(cashflows.size());
    if (n < 2) return errValue();

    // PV of negative cashflows at finance_rate
    double pvNeg = 0.0;
    // FV of positive cashflows at reinvest_rate
    double fvPos = 0.0;

    for (int i = 0; i < n; ++i) {
        if (cashflows[i] < 0.0)
            pvNeg += cashflows[i] / std::pow(1.0 + financeRate, static_cast<double>(i));
        else if (cashflows[i] > 0.0)
            fvPos += cashflows[i] * std::pow(1.0 + reinvestRate, static_cast<double>(n - 1 - i));
    }

    if (pvNeg == 0.0 || fvPos == 0.0) return errDiv0();
    return XLCellValue(std::pow(-fvPos / pvNeg, 1.0 / static_cast<double>(n - 1)) - 1.0);
}

XLCellValue Formula::fnRate(const std::vector<XLFormulaArg>& args)
{
    // RATE(nper, pmt, pv, [fv=0], [type=0], [guess=0.1])
    if (args.size() < 3 || args[0].empty() || args[1].empty() || args[2].empty()) return errValue();
    double nper  = toDouble(args[0][0]);
    double pmt   = toDouble(args[1][0]);
    double pv    = toDouble(args[2][0]);
    double fv    = (args.size() > 3 && !args[3].empty()) ? toDouble(args[3][0]) : 0.0;
    int    type  = (args.size() > 4 && !args[4].empty()) ? static_cast<int>(toDouble(args[4][0])) : 0;
    double rate  = (args.size() > 5 && !args[5].empty()) ? toDouble(args[5][0]) : 0.1;

    // Newton-Raphson: f(r) = pv*(1+r)^n + pmt*(1+r*type)*((1+r)^n - 1)/r + fv = 0
    for (int iter = 0; iter < 300; ++iter) {
        if (std::abs(rate) < 1e-15) {
            // r ≈ 0 linearisation
            double f  = pv + pmt * nper + fv;
            double df = pmt * nper * (1.0 + type * 0) - pmt;
            if (std::abs(df) < 1e-15) return errNum();
            rate -= f / df;
            continue;
        }
        double q   = std::pow(1.0 + rate, nper);
        double f   = pv * q + pmt * (1.0 + rate * type) * (q - 1.0) / rate + fv;
        double dq  = nper * std::pow(1.0 + rate, nper - 1.0);
        double df  = pv * dq +
                     pmt * type * (q - 1.0) / rate +
                     pmt * (1.0 + rate * type) * (dq * rate - (q - 1.0)) / (rate * rate);
        if (std::abs(df) < 1e-15) return errNum();
        double newRate = rate - f / df;
        if (std::abs(newRate - rate) < 1e-10) return XLCellValue(newRate);
        rate = newRate;
    }
    return errNum();
}

XLCellValue Formula::fnIpmt(const std::vector<XLFormulaArg>& args)
{
    // IPMT(rate, per, nper, pv, [fv=0], [type=0])
    if (args.size() < 4 || args[0].empty() || args[1].empty() || args[2].empty() || args[3].empty()) return errValue();
    double rate = toDouble(args[0][0]);
    double per  = toDouble(args[1][0]);
    double nper = toDouble(args[2][0]);
    double pv   = toDouble(args[3][0]);
    double fv   = (args.size() > 4 && !args[4].empty()) ? toDouble(args[4][0]) : 0.0;
    int    type = (args.size() > 5 && !args[5].empty()) ? static_cast<int>(toDouble(args[5][0])) : 0;

    if (per < 1.0 || per > nper) return errNum();

    // PMT using the same formula as fnPmt
    double pmt = 0.0;
    if (rate == 0.0) {
        pmt = -(pv + fv) / nper;
    } else {
        double q = std::pow(1.0 + rate, nper);
        pmt = -(pv * q + fv) * rate / ((1.0 + rate * type) * (q - 1.0));
    }

    // IPMT(per) = -PV(per-1) * rate
    double pvAtPer = pv * std::pow(1.0 + rate, per - 1.0) +
                     pmt * (1.0 + rate * type) *
                     (std::pow(1.0 + rate, per - 1.0) - 1.0) / (rate == 0.0 ? 1.0 : rate);

    if (rate == 0.0) return XLCellValue(0.0);
    double ipmt = -pvAtPer * rate;
    if (type == 1 && per == 1) ipmt = 0.0;
    else if (type == 1) ipmt = ipmt / (1.0 + rate);
    return XLCellValue(ipmt);
}

XLCellValue Formula::fnPpmt(const std::vector<XLFormulaArg>& args)
{
    // PPMT = PMT - IPMT
    if (args.size() < 4) return errValue();

    // Compute PMT
    std::vector<XLFormulaArg> pmtArgs = {args[0], args[2], args[3]};
    if (args.size() > 4) pmtArgs.push_back(args[4]);
    if (args.size() > 5) pmtArgs.push_back(args[5]);
    XLCellValue pmtVal = fnPmt(pmtArgs);
    if (pmtVal.type() == XLValueType::Error) return pmtVal;

    XLCellValue ipmtVal = fnIpmt(args);
    if (ipmtVal.type() == XLValueType::Error) return ipmtVal;

    double pmt  = pmtVal.type()  == XLValueType::Float ? pmtVal.get<double>()
                                                        : static_cast<double>(pmtVal.get<int64_t>());
    double ipmt = ipmtVal.type() == XLValueType::Float ? ipmtVal.get<double>()
                                                       : static_cast<double>(ipmtVal.get<int64_t>());
    return XLCellValue(pmt - ipmt);
}

void OpenXLSX::registerFinancialFunctions(XLFormulaRegistry& registry)
{
    registry.registerFunction("PMT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnPmt));
    registry.registerFunction("FV", std::make_shared<XLSimpleFormulaFunction>(Formula::fnFv));
    registry.registerFunction("PV", std::make_shared<XLSimpleFormulaFunction>(Formula::fnPv));
    registry.registerFunction("NPV", std::make_shared<XLSimpleFormulaFunction>(Formula::fnNpv));
    registry.registerFunction("NPER", std::make_shared<XLSimpleFormulaFunction>(Formula::fnNper));
    registry.registerFunction("DB", std::make_shared<XLSimpleFormulaFunction>(Formula::fnDb));
    registry.registerFunction("DDB", std::make_shared<XLSimpleFormulaFunction>(Formula::fnDdb));
    registry.registerFunction("SLN", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSln));
    registry.registerFunction("SYD", std::make_shared<XLSimpleFormulaFunction>(Formula::fnSyd));
    registry.registerFunction("IRR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIrr));
    registry.registerFunction("MIRR", std::make_shared<XLSimpleFormulaFunction>(Formula::fnMirr));
    registry.registerFunction("RATE", std::make_shared<XLSimpleFormulaFunction>(Formula::fnRate));
    registry.registerFunction("IPMT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnIpmt));
    registry.registerFunction("PPMT", std::make_shared<XLSimpleFormulaFunction>(Formula::fnPpmt));
}

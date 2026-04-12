#include <catch2/catch_all.hpp>
#include <OpenXLSX.hpp>
#include "XLFormulaEngine.hpp"
#include <string>
#include <vector>
#include <iostream>

using namespace OpenXLSX;

struct TestCase {
    std::string name;
    std::string formula;
    std::string expected;
};

std::vector<TestCase> getTestCases() {
    return {
        {"ABS", "ABS(-10.5)", "10.5"},
        {"ACOS", "ACOS(0.8)", "0.643501108793284"},
        {"AND", "AND(TRUE(), 1, 2>1)", "TRUE"},
        {"ASIN", "ASIN(0.2)", "0.201357920790331"},
        {"AVEDEV", "AVEDEV(2, 4, 6, 8)", "2"},
        {"AVERAGE", "AVERAGE(10, 20)", "15"},
        {"AVERAGEA", "AVERAGEA(\"5\", FALSE(), 10)", "5"},
        {"AVERAGEIF", "AVERAGEIF(Data!A1:A10, \"<=8\")", "5"},
        {"AVERAGEIFS", "AVERAGEIFS(Data!A1:A10, Data!A1:A10, \">2\", Data!B1:B10, \"<10\")", "8"},
        {"CEILING", "CEILING(3.14, 0.1)", "3.2"},
        {"CEILING.MATH", "CEILING.MATH(4.2)", "5"},
        {"CHAR", "CHAR(66)", "B"},
        {"CLEAN", "CLEAN(CHAR(7)&\"hello\"&CHAR(9))", "hello"},
        {"CODE", "CODE(\"B\")", "66"},
        {"CONCAT", "CONCAT(\"Hello\", \" \", \"World\")", "Hello World"},
        {"CONCATENATE", "CONCATENATE(\"Foo\", \"Bar\")", "FooBar"},
        {"CORREL", "CORREL(Data!A2:A6, Data!B2:B6)", "1"},
        {"COS", "COS(1)", "0.54030230586814"},
        {"COUNT", "COUNT(10, 20, 30, \"B\", TRUE())", "4"},
        {"COUNTA", "COUNTA(1, \"\", FALSE())", "2"},
        {"COUNTBLANK", "COUNTBLANK(Data!F1:G5)", "10"},
        {"COUNTIF", "COUNTIF(Data!B1:B10, \"<10\")", "6"},
        {"COUNTIFS", "COUNTIFS(Data!A1:A10, \">2\", Data!B1:B10, \"<15\")", "8"},
        {"COVAR", "COVAR(Data!A2:A5, Data!B2:B5)", "3.75"},
        {"COVARIANCE.P", "COVARIANCE.P(Data!A2:A5, Data!B2:B5)", "3.75"},
        {"COVARIANCE.S", "COVARIANCE.S(Data!A2:A5, Data!B2:B5)", "5"},
        {"DATE", "DATE(2024, 2, 29)", "45351"},
        {"DAY", "DAY(46000)", "9"},
        {"DAYS", "DAYS(46000, 45000)", "1000"},
        {"DAYS360", "DAYS360(45000, 46000)", "984"},
        {"DB", "DB(20000, 2000, 10, 2)", "3271.28"},
        {"DDB", "DDB(20000, 2000, 10, 2)", "3200"},
        {"DEGREES", "DEGREES(1)", "57.2957795130823"},
        {"DEVSQ", "DEVSQ(2, 4, 6, 8)", "20"},
        {"EDATE", "EDATE(46000, 2)", "46062"},
        {"EOMONTH", "EOMONTH(46000, 2)", "46081"},
        {"EXACT", "EXACT(\"Hello\", \"hello\")", "FALSE"},
        {"EXP", "EXP(2)", "7.38905609893065"},
        {"FALSE", "FALSE()", "FALSE"},
        {"FIND", "FIND(\"l\", \"hello\")", "3"},
        {"FISHER", "FISHER(0.8)", "1.09861228866811"},
        {"FISHERINV", "FISHERINV(0.8)", "0.664036770267849"},
        {"FLOOR", "FLOOR(3.14, 0.1)", "3.1"},
        {"FLOOR.MATH", "FLOOR.MATH(4.8)", "4"},
        {"FV", "FV(0.06, 12, -200)", "3373.98823945184"},
        {"HLOOKUP", "HLOOKUP(\"Item3\", Data!C1:E10, 3, FALSE())", "ERROR: HLOOKUP no result found"},
        {"HOUR", "HOUR(0.75)", "18"},
        {"IF", "IF(FALSE(), 1, 2)", "2"},
        {"IFERROR", "IFERROR(1/0, 999)", "999"},
        {"IFNA", "IFNA(VLOOKUP(100, Data!A1:B10, 2, FALSE()), 888)", "888"},
        {"IFS", "IFS(1>2, 1, 2>1, 2)", "2"},
        {"INDEX", "INDEX(Data!A1:A10, 3)", "6"},
        {"INT", "INT(4.8)", "4"},
        {"INTERCEPT", "INTERCEPT(Data!A2:A6, Data!B2:B6)", "4.44089209850063E-16"},
        {"ISBLANK", "ISBLANK(Data!F1)", "TRUE"},
        {"ISERR", "ISERR(1/0)", "TRUE"},
        {"ISERROR", "ISERROR(1/0)", "TRUE"},
        {"ISEVEN", "ISEVEN(3)", "FALSE"},
        {"ISLOGICAL", "ISLOGICAL(1)", "FALSE"},
        {"ISNA", "ISNA(VLOOKUP(100, Data!A1:B10, 2, FALSE()))", "TRUE"},
        {"ISNONTEXT", "ISNONTEXT(\"A\")", "FALSE"},
        {"ISNUMBER", "ISNUMBER(\"1\")", "FALSE"},
        {"ISODD", "ISODD(4)", "FALSE"},
        {"ISOWEEKNUM", "ISOWEEKNUM(46000)", "50"},
        {"ISTEXT", "ISTEXT(1)", "FALSE"},
        {"LARGE", "LARGE(Data!A1:A10, 2)", "18"},
        {"LEFT", "LEFT(\"hello\", 3)", "hel"},
        {"LEN", "LEN(\"hello\")", "5"},
        {"LOG", "LOG(1000, 10)", "3"},
        {"LOG10", "LOG10(1000)", "3"},
        {"LOWER", "LOWER(\"HELLO\")", "hello"},
        {"MATCH", "MATCH(6, Data!A1:A10, 0)", "3"},
        {"MAX", "MAX(10, 20, 30)", "30"},
        {"MAXIFS", "MAXIFS(Data!A1:A10, Data!A1:A10, \"<=8\")", "8"},
        {"MEDIAN", "MEDIAN(10, 20, 30, 40)", "25"},
        {"MID", "MID(\"hello\", 2, 3)", "ell"},
        {"MIN", "MIN(10, 20, 30)", "10"},
        {"MINIFS", "MINIFS(Data!A1:A10, Data!A1:A10, \">=6\")", "6"},
        {"MINUTE", "MINUTE(0.75)", "0"},
        {"MOD", "MOD(10, 3)", "1"},
        {"MONTH", "MONTH(46000)", "12"},
        {"MROUND", "MROUND(10, 3)", "9"},
        {"NETWORKDAYS", "NETWORKDAYS(45000, 46000)", "715"},
        {"NOT", "NOT(FALSE())", "TRUE"},
        {"NPER", "NPER(0.06, -200, 2000)", "15.7252085438878"},
        {"NPV", "NPV(0.06, 200, 200, 200)", "534.602389892327"},
        {"OR", "OR(FALSE(), FALSE(), TRUE())", "TRUE"},
        {"PEARSON", "PEARSON(Data!A2:A6, Data!B2:B6)", "1"},
        {"PERCENTILE", "PERCENTILE(Data!A1:A10, 0.8)", "16.4"},
        {"PERCENTILE.EXC", "PERCENTILE.EXC(Data!A1:A10, 0.8)", "17.6"},
        {"PERCENTILE.INC", "PERCENTILE.INC(Data!A1:A10, 0.8)", "16.4"},
        {"PERMUT", "PERMUT(6, 3)", "120"},
        {"PERMUTATIONA", "PERMUTATIONA(6, 3)", "216"},
        {"PI", "PI()", "3.14159265358979"},
        {"PMT", "PMT(0.06, 12, 2000)", "-238.554058761327"},
        {"POWER", "POWER(3, 4)", "81"},
        {"PROPER", "PROPER(\"hello world\")", "Hello World"},
        {"PV", "PV(0.06, 12, -200)", "1676.76878807667"},
        {"QUARTILE", "QUARTILE(Data!A1:A10, 2)", "11"},
        {"QUARTILE.EXC", "QUARTILE.EXC(Data!A1:A10, 2)", "11"},
        {"QUARTILE.INC", "QUARTILE.INC(Data!A1:A10, 2)", "11"},
        {"RADIANS", "RADIANS(90)", "1.5707963267949"},
        {"RANK", "RANK(6, Data!A1:A10)", "8"},
        {"RANK.EQ", "RANK.EQ(6, Data!A1:A10)", "8"},
        {"REPLACE", "REPLACE(\"hello\", 1, 2, \"j\")", "jllo"},
        {"REPT", "REPT(\"b\", 4)", "bbbb"},
        {"RIGHT", "RIGHT(\"hello\", 3)", "llo"},
        {"ROUND", "ROUND(3.14159, 2)", "3.14"},
        {"ROUNDDOWN", "ROUNDDOWN(3.14159, 2)", "3.14"},
        {"ROUNDUP", "ROUNDUP(3.14159, 2)", "3.15"},
        {"RSQ", "RSQ(Data!A2:A6, Data!B2:B6)", "1"},
        {"SEARCH", "SEARCH(\"L\", \"hello\")", "3"},
        {"SECOND", "SECOND(0.75)", "0"},
        {"SIGN", "SIGN(10)", "1"},
        {"SIN", "SIN(1)", "0.841470984807897"},
        {"SLN", "SLN(20000, 2000, 10)", "1800"},
        {"SLOPE", "SLOPE(Data!A2:A6, Data!B2:B6)", "1.33333333333333"},
        {"SMALL", "SMALL(Data!A1:A10, 2)", "4"},
        {"SQRT", "SQRT(16)", "4"},
        {"STANDARDIZE", "STANDARDIZE(3, 2, 1)", "1"},
        {"STDEV", "STDEV(2, 4, 6, 8)", "2.58198889747161"},
        {"STDEV.P", "STDEV.P(2, 4, 6, 8)", "2.23606797749979"},
        {"STDEV.S", "STDEV.S(2, 4, 6, 8)", "2.58198889747161"},
        {"STDEVA", "STDEVA(2, 4, 6, 8)", "2.58198889747161"},
        {"STDEVP", "STDEVP(2, 4, 6, 8)", "2.23606797749979"},
        {"STDEVPA", "STDEVPA(2, 4, 6, 8)", "2.23606797749979"},
        {"SUBSTITUTE", "SUBSTITUTE(\"hello\", \"l\", \"w\")", "hewwo"},
        {"SUM", "SUM(10, 20, 30)", "60"},
        {"SUMIF", "SUMIF(Data!A1:A10, \"<=8\")", "20"},
        {"SUMIFS", "SUMIFS(Data!A1:A10, Data!A1:A10, \">2\", Data!B1:B10, \"<10\")", "40"},
        {"SUMPRODUCT", "SUMPRODUCT(Data!A2:A6, Data!B2:B6)", "270"},
        {"SUMSQ", "SUMSQ(2, 4, 6)", "56"},
        {"SUMX2MY2", "SUMX2MY2(Data!A2:A6, Data!B2:B6)", "157.5"},
        {"SUMX2PY2", "SUMX2PY2(Data!A2:A6, Data!B2:B6)", "562.5"},
        {"SUMXMY2", "SUMXMY2(Data!A2:A6, Data!B2:B6)", "22.5"},
        {"SWITCH", "SWITCH(2, 1, \"A\", 2, \"B\", \"C\")", "B"},
        {"SYD", "SYD(20000, 2000, 10, 2)", "2945.45454545455"},
        {"T", "T(1)", ""},
        {"TAN", "TAN(1)", "1.5574077246549"},
        {"TEXT", "TEXT(46000, \"yyyy/mm/dd\")", "2025/12/09"},
        {"TEXTJOIN", "TEXTJOIN(\"-\", TRUE(), \"X\", \"Y\")", "X-Y"},
        {"TIME", "TIME(14, 30, 0)", "0.604166666666667"},
        {"TRIM", "TRIM(\"  hello  world  \")", "hello  world"},
        {"TRIMMEAN", "TRIMMEAN(Data!A1:A10, 0.4)", "11"},
        {"TRUE", "TRUE()", "TRUE"},
        {"TRUNC", "TRUNC(3.14159, 2)", "3.14"},
        {"UNICHAR", "UNICHAR(66)", "B"},
        {"UNICODE", "UNICODE(\"B\")", "66"},
        {"UPPER", "UPPER(\"hello\")", "HELLO"},
        {"VALUE", "VALUE(\"2\")", "2"},
        {"VAR", "VAR(2, 4, 6, 8)", "6.66666666666667"},
        {"VAR.P", "VAR.P(2, 4, 6, 8)", "5"},
        {"VAR.S", "VAR.S(2, 4, 6, 8)", "6.66666666666667"},
        {"VARA", "VARA(2, 4, 6, 8)", "6.66666666666667"},
        {"VARP", "VARP(2, 4, 6, 8)", "5"},
        {"VARPA", "VARPA(2, 4, 6, 8)", "5"},
        {"VLOOKUP", "VLOOKUP(6, Data!A1:E10, 3, FALSE())", "Item3"},
        {"WEEKDAY", "WEEKDAY(46000, 2)", "2"},
        {"WEEKNUM", "WEEKNUM(46000, 2)", "50"},
        {"WORKDAY", "WORKDAY(46000, 2)", "46002"},
        {"XLOOKUP", "XLOOKUP(6, Data!A1:A10, Data!C1:C10)", "Item3"},
        {"YEAR", "YEAR(46000)", "2025"},
    };
}

TEST_CASE("FormulaEnginevsExcelizeStandard", "[XLFormulaEngine][Excelize]") {
    XLDocument doc;
    std::string testPath = "./excelize_test_doc_new.xlsx";
    doc.create(std::string_view(testPath));
    auto wbk = doc.workbook();
    wbk.addWorksheet("Data");
    auto dataSheet = wbk.worksheet("Data");
    
    // Populate Data sheet with new values to match Go test
    for (int i = 1; i <= 10; ++i) {
        dataSheet.cell("A" + std::to_string(i)).value() = i * 2;
        dataSheet.cell("B" + std::to_string(i)).value() = static_cast<double>(i) * 1.5;
        dataSheet.cell("C" + std::to_string(i)).value() = "Item" + std::to_string(i);
        dataSheet.cell("D" + std::to_string(i)).value() = (i % 3 == 0);
        dataSheet.cell("E" + std::to_string(i)).value() = 46000 + i;
    }
    doc.save();

    XLFormulaEngine engine;
    
    auto resolver = [&wbk, &dataSheet](std::string_view ref) -> XLCellValue {
        try {
            auto bang = ref.find('!');
            if (bang != std::string_view::npos) {
                std::string sheetName = std::string(ref.substr(0, bang));
                // Remove optional quotes from sheet name
                if (sheetName.length() >= 2 && sheetName.front() == '\'' && sheetName.back() == '\'') {
                    sheetName = sheetName.substr(1, sheetName.length() - 2);
                }
                std::string cellAddr = std::string(ref.substr(bang + 1));
                return wbk.worksheet(sheetName).cell(cellAddr).value();
            } else {
                return dataSheet.cell(std::string(ref)).value();
            }
        } catch (...) {
            return XLCellValue{};
        }
    };

    auto testCases = getTestCases();
    
    // using Catch2 GENERATE
    auto testCase = GENERATE_REF(from_range(testCases));
    
    DYNAMIC_SECTION("Function: " << testCase.name) {
        auto val = engine.evaluate(testCase.formula, resolver);
        
        std::string res;
        switch(val.type()) {
            case XLValueType::Boolean: res = val.get<bool>() ? "TRUE" : "FALSE"; break;
            case XLValueType::Integer: res = std::to_string(val.get<int64_t>()); break;
            case XLValueType::Float: {
                double v = val.get<double>();
                char buf[64];
                snprintf(buf, sizeof(buf), "%g", v); // rough string match
                res = buf;
                break;
            }
            case XLValueType::String:  res = val.get<std::string>(); break;
            case XLValueType::Error:   res = "ERROR: " + val.get<std::string>(); break;
            default: res = "EMPTY"; break;
        }
        
        // Custom float assertion logic
        if (testCase.name == "HLOOKUP") {
            REQUIRE(res == "ERROR: #N/A");
        } else if (testCase.expected.find("ERROR:") == 0) {
            std::cout << "RES: " << res << " EXPECTED: " << testCase.expected << "\n";
            REQUIRE(res.find("ERROR:") == 0);
        } else {
            try {
                double e_val = std::stod(testCase.expected);
                double a_val = val.type() == XLValueType::Float ? val.get<double>() : (val.type() == XLValueType::Integer ? static_cast<double>(val.get<int64_t>()) : std::stod(res));
                
                if (std::abs(e_val) < 1e-9 || std::abs(a_val) < 1e-9) {
                    REQUIRE_THAT(a_val, Catch::Matchers::WithinAbs(e_val, 1e-9));
                } else {
                    REQUIRE_THAT(a_val, Catch::Matchers::WithinRel(e_val, 1e-4));
                }
            } catch (const std::invalid_argument&) {
                if (testCase.expected == "TRUE" || testCase.expected == "FALSE") {
                    // Sometimes Excelize gives 1 / 0 for bools
                    if (testCase.expected == "TRUE" && res == "1") res = "TRUE";
                    if (testCase.expected == "FALSE" && res == "0") res = "FALSE";
                }
                REQUIRE(res == testCase.expected);
            } catch (const std::out_of_range&) {
                REQUIRE(res == testCase.expected);
            }
        }
    }
}


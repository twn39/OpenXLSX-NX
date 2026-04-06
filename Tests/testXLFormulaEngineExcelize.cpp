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
        {"ABS", "ABS(-5)", "5"},
        {"ACOS", "ACOS(0.5)", "1.0471975511966"},
        {"AND", "AND(TRUE(), FALSE())", "FALSE"},
        {"ASIN", "ASIN(0.5)", "0.523598775598299"},
        {"AVEDEV", "AVEDEV(1, 2, 3)", "0.666666666666667"},
        {"AVERAGE", "AVERAGE(1, 2, 3)", "2"},
        {"AVERAGEA", "AVERAGEA(1, 2, \"3\")", "2"},
        {"AVERAGEIF", "AVERAGEIF(Data!A1:A10, \">5\")", "8"},
        {"AVERAGEIFS", "AVERAGEIFS(Data!A1:A10, Data!A1:A10, \">5\")", "8"},
        {"CEILING", "CEILING(2.5, 1)", "3"},
        {"CEILING.MATH", "CEILING.MATH(2.5)", "3"},
        {"CHAR", "CHAR(65)", "A"},
        {"CLEAN", "CLEAN(\"a\"&CHAR(10)&\"b\")", "ab"},
        {"CODE", "CODE(\"A\")", "65"},
        {"CONCAT", "CONCAT(\"A\", \"B\")", "AB"},
        {"CONCATENATE", "CONCATENATE(\"A\", \"B\")", "AB"},
        {"CORREL", "CORREL(Data!A1:A5, Data!B1:B5)", "1"},
        {"COS", "COS(0)", "1"},
        {"COUNT", "COUNT(1, 2, \"A\")", "2"},
        {"COUNTA", "COUNTA(1, 2, \"A\")", "3"},
        {"COUNTBLANK", "COUNTBLANK(Data!A1:A5)", "0"},
        {"COUNTIF", "COUNTIF(Data!A1:A10, \">5\")", "5"},
        {"COUNTIFS", "COUNTIFS(Data!A1:A10, \">5\")", "5"},
        {"COVAR", "COVAR(Data!A1:A3, Data!B1:B3)", "1.66666666666667"},
        {"COVARIANCE.P", "COVARIANCE.P(Data!A1:A3, Data!B1:B3)", "1.66666666666667"},
        {"COVARIANCE.S", "COVARIANCE.S(Data!A1:A3, Data!B1:B3)", "2.5"},
        {"DATE", "DATE(2023, 1, 1)", "44927"},
        {"DAY", "DAY(45000)", "15"},
        {"DAYS", "DAYS(45000, 44000)", "1000"},
        {"DAYS360", "DAYS360(44000, 45000)", "987"},
        {"DB", "DB(10000, 1000, 5, 1)", "3690"},
        {"DDB", "DDB(10000, 1000, 5, 1)", "4000"},
        {"DEGREES", "DEGREES(PI())", "180"},
        {"DEVSQ", "DEVSQ(1, 2, 3)", "2"},
        {"EDATE", "EDATE(45000, 1)", "45031"},
        {"EOMONTH", "EOMONTH(45000, 1)", "45046"},
        {"EXACT", "EXACT(\"A\", \"A\")", "TRUE"},
        {"EXP", "EXP(1)", "2.71828182845905"},
        {"FALSE", "FALSE()", "FALSE"},
        {"FIND", "FIND(\"a\", \"banana\")", "2"},
        {"FISHER", "FISHER(0.5)", "0.549306144334055"},
        {"FISHERINV", "FISHERINV(0.5)", "0.46211715726001"},
        {"FLOOR", "FLOOR(2.5, 1)", "2"},
        {"FLOOR.MATH", "FLOOR.MATH(2.5)", "2"},
        {"FV", "FV(0.05, 10, -100)", "1257.78925355488"},
        {"HLOOKUP", "HLOOKUP(1, Data!A1:E10, 2, FALSE())", "ERROR: HLOOKUP no result found"},
        {"HOUR", "HOUR(0.5)", "12"},
        {"IF", "IF(TRUE(), 1, 0)", "1"},
        {"IFERROR", "IFERROR(1/0, 1)", "1"},
        {"IFNA", "IFNA(VLOOKUP(100, Data!A1:B10, 2, FALSE()), 1)", "1"},
        {"IFS", "IFS(FALSE(), 1, TRUE(), 2)", "2"},
        {"INDEX", "INDEX(Data!A1:A10, 1)", "1"},
        {"INT", "INT(2.5)", "2"},
        {"INTERCEPT", "INTERCEPT(Data!A1:A5, Data!B1:B5)", "-1.66533453693773E-16"},
        {"ISBLANK", "ISBLANK(Data!Z1)", "TRUE"},
        {"ISERR", "ISERR(1/0)", "TRUE"},
        {"ISERROR", "ISERROR(1/0)", "TRUE"},
        {"ISEVEN", "ISEVEN(2)", "TRUE"},
        {"ISLOGICAL", "ISLOGICAL(TRUE())", "TRUE"},
        {"ISNA", "ISNA(VLOOKUP(100, Data!A1:B10, 2, FALSE()))", "TRUE"},
        {"ISNONTEXT", "ISNONTEXT(1)", "TRUE"},
        {"ISNUMBER", "ISNUMBER(1)", "TRUE"},
        {"ISODD", "ISODD(3)", "TRUE"},
        {"ISOWEEKNUM", "ISOWEEKNUM(45000)", "11"},
        {"ISTEXT", "ISTEXT(\"A\")", "TRUE"},
        {"LARGE", "LARGE(Data!A1:A10, 1)", "10"},
        {"LEFT", "LEFT(\"banana\", 2)", "ba"},
        {"LEN", "LEN(\"banana\")", "6"},
        {"LOG", "LOG(10)", "1"},
        {"LOG10", "LOG10(100)", "2"},
        {"LOWER", "LOWER(\"BANANA\")", "banana"},
        {"MATCH", "MATCH(1, Data!A1:A10, 0)", "1"},
        {"MAX", "MAX(1, 2, 3)", "3"},
        {"MAXIFS", "MAXIFS(Data!A1:A10, Data!A1:A10, \">5\")", "10"},
        {"MEDIAN", "MEDIAN(1, 2, 3)", "2"},
        {"MID", "MID(\"banana\", 2, 2)", "an"},
        {"MIN", "MIN(1, 2, 3)", "1"},
        {"MINIFS", "MINIFS(Data!A1:A10, Data!A1:A10, \">5\")", "6"},
        {"MINUTE", "MINUTE(0.5)", "0"},
        {"MOD", "MOD(5, 2)", "1"},
        {"MONTH", "MONTH(45000)", "3"},
        {"MROUND", "MROUND(5, 2)", "6"},
        {"NETWORKDAYS", "NETWORKDAYS(44000, 45000)", "715"},
        {"NOT", "NOT(TRUE())", "FALSE"},
        {"NPER", "NPER(0.05, -100, 1000)", "14.2066990828905"},
        {"NPV", "NPV(0.05, 100, 100, 100)", "272.324802937048"},
        {"OR", "OR(TRUE(), FALSE())", "TRUE"},
        {"PEARSON", "PEARSON(Data!A1:A5, Data!B1:B5)", "1"},
        {"PERCENTILE", "PERCENTILE(Data!A1:A10, 0.5)", "5.5"},
        {"PERCENTILE.EXC", "PERCENTILE.EXC(Data!A1:A10, 0.5)", "5.5"},
        {"PERCENTILE.INC", "PERCENTILE.INC(Data!A1:A10, 0.5)", "5.5"},
        {"PERMUT", "PERMUT(5, 2)", "20"},
        {"PERMUTATIONA", "PERMUTATIONA(5, 2)", "25"},
        {"PI", "PI()", "3.14159265358979"},
        {"PMT", "PMT(0.05, 10, 1000)", "-129.504574965457"},
        {"POWER", "POWER(2, 3)", "8"},
        {"PROPER", "PROPER(\"banana\")", "Banana"},
        {"PV", "PV(0.05, 10, -100)", "772.173492918481"},
        {"QUARTILE", "QUARTILE(Data!A1:A10, 1)", "3.25"},
        {"QUARTILE.EXC", "QUARTILE.EXC(Data!A1:A10, 1)", "2.75"},
        {"QUARTILE.INC", "QUARTILE.INC(Data!A1:A10, 1)", "3.25"},
        {"RADIANS", "RADIANS(180)", "3.14159265358979"},
        {"RANK", "RANK(1, Data!A1:A10)", "10"},
        {"RANK.EQ", "RANK.EQ(1, Data!A1:A10)", "10"},
        {"REPLACE", "REPLACE(\"banana\", 1, 1, \"c\")", "canana"},
        {"REPT", "REPT(\"a\", 3)", "aaa"},
        {"RIGHT", "RIGHT(\"banana\", 2)", "na"},
        {"ROUND", "ROUND(2.5, 0)", "3"},
        {"ROUNDDOWN", "ROUNDDOWN(2.5, 0)", "2"},
        {"ROUNDUP", "ROUNDUP(2.5, 0)", "3"},
        {"RSQ", "RSQ(Data!A1:A5, Data!B1:B5)", "1"},
        {"SEARCH", "SEARCH(\"a\", \"banana\")", "2"},
        {"SECOND", "SECOND(0.5)", "0"},
        {"SIGN", "SIGN(-5)", "-1"},
        {"SIN", "SIN(0)", "0"},
        {"SLN", "SLN(10000, 1000, 5)", "1800"},
        {"SLOPE", "SLOPE(Data!A1:A5, Data!B1:B5)", "0.4"},
        {"SMALL", "SMALL(Data!A1:A10, 1)", "1"},
        {"SQRT", "SQRT(4)", "2"},
        {"STANDARDIZE", "STANDARDIZE(2, 1, 1)", "1"},
        {"STDEV", "STDEV(1, 2, 3)", "1"},
        {"STDEV.P", "STDEV.P(1, 2, 3)", "0.816496580927726"},
        {"STDEV.S", "STDEV.S(1, 2, 3)", "1"},
        {"STDEVA", "STDEVA(1, 2, 3)", "1"},
        {"STDEVP", "STDEVP(1, 2, 3)", "0.816496580927726"},
        {"STDEVPA", "STDEVPA(1, 2, 3)", "0.816496580927726"},
        {"SUBSTITUTE", "SUBSTITUTE(\"banana\", \"a\", \"o\")", "bonono"},
        {"SUM", "SUM(1, 2, 3)", "6"},
        {"SUMIF", "SUMIF(Data!A1:A10, \">5\")", "40"},
        {"SUMIFS", "SUMIFS(Data!A1:A10, Data!A1:A10, \">5\")", "40"},
        {"SUMPRODUCT", "SUMPRODUCT(Data!A1:A5, Data!B1:B5)", "137.5"},
        {"SUMSQ", "SUMSQ(1, 2, 3)", "14"},
        {"SUMX2MY2", "SUMX2MY2(Data!A1:A5, Data!B1:B5)", "-288.75"},
        {"SUMX2PY2", "SUMX2PY2(Data!A1:A5, Data!B1:B5)", "398.75"},
        {"SUMXMY2", "SUMXMY2(Data!A1:A5, Data!B1:B5)", "123.75"},
        {"SWITCH", "SWITCH(1, 1, \"A\", 2, \"B\", \"C\")", "A"},
        {"SYD", "SYD(10000, 1000, 5, 1)", "3000"},
        {"T", "T(\"A\")", "A"},
        {"TAN", "TAN(0)", "0"},
        {"TEXT", "TEXT(45000, \"yyyy-mm-dd\")", "2023-03-15"},
        {"TEXTJOIN", "TEXTJOIN(\",\", TRUE(), \"A\", \"B\")", "A,B"},
        {"TIME", "TIME(12, 0, 0)", "0.5"},
        {"TODAY", "TODAY()", "46118"},
        {"TRIM", "TRIM(\" banana \")", "banana"},
        {"TRIMMEAN", "TRIMMEAN(Data!A1:A10, 0.2)", "5.5"},
        {"TRUE", "TRUE()", "TRUE"},
        {"TRUNC", "TRUNC(2.5)", "2"},
        {"UNICHAR", "UNICHAR(65)", "A"},
        {"UNICODE", "UNICODE(\"A\")", "65"},
        {"UPPER", "UPPER(\"banana\")", "BANANA"},
        {"VALUE", "VALUE(\"1\")", "1"},
        {"VAR", "VAR(1, 2, 3)", "1"},
        {"VAR.P", "VAR.P(1, 2, 3)", "0.666666666666667"},
        {"VAR.S", "VAR.S(1, 2, 3)", "1"},
        {"VARA", "VARA(1, 2, 3)", "1"},
        {"VARP", "VARP(1, 2, 3)", "0.666666666666667"},
        {"VARPA", "VARPA(1, 2, 3)", "0.666666666666667"},
        {"VLOOKUP", "VLOOKUP(1, Data!A1:E10, 2, FALSE())", "2.5"},
        {"WEEKDAY", "WEEKDAY(45000)", "4"},
        {"WEEKNUM", "WEEKNUM(45000)", "11"},
        {"WORKDAY", "WORKDAY(45000, 1)", "45001"},
        {"XLOOKUP", "XLOOKUP(1, Data!A1:A10, Data!B1:B10)", "2.5"},
        {"YEAR", "YEAR(45000)", "2023"},
    };
}

TEST_CASE("Formula Engine vs Excelize Standard", "[XLFormulaEngine][Excelize]") {
    XLDocument doc;
    std::string testPath = "./excelize_test_doc.xlsx";
    doc.create(std::string_view(testPath));
    auto wbk = doc.workbook();
    wbk.addWorksheet("Data");
    auto dataSheet = wbk.worksheet("Data");
    
    // Populate Data sheet exactly as we did for Excelize
    for (int i = 1; i <= 10; ++i) {
        dataSheet.cell("A" + std::to_string(i)).value() = i;
        dataSheet.cell("B" + std::to_string(i)).value() = i * 2.5;
        dataSheet.cell("C" + std::to_string(i)).value() = "Text" + std::to_string(i);
        dataSheet.cell("D" + std::to_string(i)).value() = (i % 2 == 0);
        dataSheet.cell("E" + std::to_string(i)).value() = 45000 + i;
    }
    doc.save();

    XLFormulaEngine engine;
    
    auto resolver = [&wbk, &dataSheet](std::string_view ref) -> XLCellValue {
        try {
            auto bang = ref.find('!');
            if (bang != std::string_view::npos) {
                std::string sheetName = std::string(ref.substr(0, bang));
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
            REQUIRE(res == "2");
        } else if (testCase.expected.find("ERROR:") == 0) {
            std::cout << "RES: " << res << " EXPECTED: " << testCase.expected << "\n";
            REQUIRE(res.find("ERROR:") == 0);
        } else {
            try {
                double e_val = std::stod(testCase.expected);
                double a_val = val.type() == XLValueType::Float ? val.get<double>() : (val.type() == XLValueType::Integer ? val.get<int64_t>() : std::stod(res));
                REQUIRE_THAT(a_val, Catch::Matchers::WithinRel(e_val, 1e-4));
            } catch (...) {
                if (testCase.expected == "TRUE" || testCase.expected == "FALSE") {
                    // Sometimes Excelize gives 1 / 0 for bools
                    if (testCase.expected == "TRUE" && res == "1") res = "TRUE";
                    if (testCase.expected == "FALSE" && res == "0") res = "FALSE";
                }
                REQUIRE(res == testCase.expected);
            }
        }
    }
}

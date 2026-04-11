#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <cmath>

using namespace OpenXLSX;

// Helper: create a no-cell resolver (for pure arithmetic tests)
static const XLCellResolver noResolver{};

// Helper: return a resolver that maps refs to fixed values
static XLCellResolver makeMapResolver(std::initializer_list<std::pair<std::string, XLCellValue>> init)
{
    auto m = std::make_shared<std::unordered_map<std::string, XLCellValue>>();
    for (auto& p : init) m->insert_or_assign(p.first, p.second);
    return [m](std::string_view ref) -> XLCellValue {
        auto it = m->find(std::string(ref));
        return it != m->end() ? it->second : XLCellValue{};
    };
}

TEST_CASE("XLFormulaEngine – Lexer", "[XLFormulaEngine]")
{
    SECTION("Numbers")
    {
        auto toks = XLFormulaLexer::tokenize("=3.14");
        REQUIRE(toks[0].kind == XLTokenKind::Number);
        REQUIRE(toks[0].number == Catch::Approx(3.14));
    }

    SECTION("String literal")
    {
        auto toks = XLFormulaLexer::tokenize("=\"hello\"");
        REQUIRE(toks[0].kind == XLTokenKind::String);
        REQUIRE(toks[0].text == "hello");
    }

    SECTION("Bool")
    {
        auto toks = XLFormulaLexer::tokenize("=TRUE");
        REQUIRE(toks[0].kind == XLTokenKind::Bool);
        REQUIRE(toks[0].boolean == true);
    }

    SECTION("Cell ref")
    {
        auto toks = XLFormulaLexer::tokenize("=A1");
        REQUIRE(toks[0].kind == XLTokenKind::CellRef);
        REQUIRE(toks[0].text == "A1");
    }

    SECTION("Range ref")
    {
        auto toks = XLFormulaLexer::tokenize("=A1:C3");
        REQUIRE(toks[0].kind == XLTokenKind::CellRef);
        REQUIRE(toks[0].text == "A1:C3");
    }

    SECTION("Operators")
    {
        auto toks = XLFormulaLexer::tokenize("=<>=");
        REQUIRE(toks[0].kind == XLTokenKind::NEq);
        REQUIRE(toks[1].kind == XLTokenKind::Eq);
    }
}

TEST_CASE("XLFormulaEngine – Arithmetic", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("Data-driven arithmetic evaluation")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({{"=1+2", 3.0},
                                                                     {"=2*3", 6.0},
                                                                     {"=1+2*3", 7.0},
                                                                     {"=(1+2)*3", 9.0},
                                                                     {"=10-4", 6.0},
                                                                     {"=10/4", 2.5},
                                                                     {"=2^10", 1024.0},
                                                                     {"=-5", -5.0},
                                                                     {"=50%", 0.5},
                                                                     {"1+1", 2.0}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected));
    }

    SECTION("Div by zero") { REQUIRE(eng.evaluate("=1/0").type() == XLValueType::Error); }
}

TEST_CASE("XLFormulaEngine – Comparison", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("Data-driven comparison evaluation")
    {
        auto [expr, expected] = GENERATE(table<std::string, bool>(
            {{"=1=1", true}, {"=1<>2", true}, {"=1<2", true}, {"=2>=2", true}, {"=\"A\"=\"a\"", true}, {"=\"abc\"<\"abd\"", true}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<bool>() == expected);
    }
}

TEST_CASE("XLFormulaEngine – String concat", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;
    SECTION("Ampersand")
    {
        auto r = eng.evaluate("=\"Hello\"&\" \"&\"World\"");
        REQUIRE(r.get<std::string>() == "Hello World");
    }
}

TEST_CASE("XLFormulaEngine – Cell refs", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;
    auto            resolver = makeMapResolver({{"A1", XLCellValue(10.0)}, {"B1", XLCellValue(5.0)}});

    SECTION("Single ref") { REQUIRE(eng.evaluate("=A1", resolver).get<double>() == Catch::Approx(10.0)); }
    SECTION("Two refs") { REQUIRE(eng.evaluate("=A1+B1", resolver).get<double>() == Catch::Approx(15.0)); }
    SECTION("Ref * literal") { REQUIRE(eng.evaluate("=A1*2", resolver).get<double>() == Catch::Approx(20.0)); }
}

TEST_CASE("XLFormulaEngine – Range functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;
    // A1=1, B1=2, C1=3
    auto resolver = makeMapResolver({
        {"A1", XLCellValue(1.0)},
        {"B1", XLCellValue(2.0)},
        {"C1", XLCellValue(3.0)},
        {"A2", XLCellValue(4.0)},
        {"B2", XLCellValue(5.0)},
        {"C2", XLCellValue(6.0)},
    });

    SECTION("SUM") { REQUIRE(eng.evaluate("=SUM(A1:C1)", resolver).get<double>() == Catch::Approx(6.0)); }
    SECTION("AVERAGE") { REQUIRE(eng.evaluate("=AVERAGE(A1:C1)", resolver).get<double>() == Catch::Approx(2.0)); }
    SECTION("MIN") { REQUIRE(eng.evaluate("=MIN(A1:C1)", resolver).get<double>() == Catch::Approx(1.0)); }
    SECTION("MAX") { REQUIRE(eng.evaluate("=MAX(A1:C1)", resolver).get<double>() == Catch::Approx(3.0)); }
    SECTION("COUNT") { REQUIRE(eng.evaluate("=COUNT(A1:C1)", resolver).get<int64_t>() == 3); }
    SECTION("SUM 2D") { REQUIRE(eng.evaluate("=SUM(A1:C2)", resolver).get<double>() == Catch::Approx(21.0)); }

    SECTION("SUM multiple args") { REQUIRE(eng.evaluate("=SUM(A1,B1,C1)", resolver).get<double>() == Catch::Approx(6.0)); }
}

TEST_CASE("XLFormulaEngine – Math functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("Data-driven math functions evaluation")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({{"=ABS(-5)", 5.0},
                                                                     {"=ABS(5)", 5.0},
                                                                     {"=SQRT(9)", 3.0},
                                                                     {"=MOD(10,3)", 1.0},
                                                                     {"=POWER(2,8)", 256.0},
                                                                     {"=ROUND(3.567,2)", 3.57},
                                                                     {"=ROUNDUP(3.111,2)", 3.12},
                                                                     {"=ROUNDDOWN(3.999,2)", 3.99}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected));
    }

    SECTION("INT") { REQUIRE(eng.evaluate("=INT(3.9)").get<int64_t>() == 3); }
    SECTION("SQRT neg") { REQUIRE(eng.evaluate("=SQRT(-1)").type() == XLValueType::Error); }
}

TEST_CASE("XLFormulaEngine – Logical functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("String returning logical functions")
    {
        auto [expr, expected] = GENERATE(table<std::string, std::string>({{"=IF(1>0,\"yes\",\"no\")", "yes"},
                                                                          {"=IF(0,\"yes\",\"no\")", "no"},
                                                                          {"=IFS(1=0,\"no\",1=1,\"yes\")", "yes"},
                                                                          {"=SWITCH(2,1,\"one\",2,\"two\",\"other\")", "two"},
                                                                          {"=SWITCH(3,1,\"one\",2,\"two\",\"other\")", "other"}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<std::string>() == expected);
    }

    SECTION("Boolean returning logical functions")
    {
        auto [expr, expected] = GENERATE(table<std::string, bool>(
            {{"=AND(1,1,1)", true}, {"=AND(1,0,1)", false}, {"=OR(0,0,1)", true}, {"=OR(0,0,0)", false}, {"=NOT(FALSE)", true}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<bool>() == expected);
    }

    SECTION("Double returning error functions")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>({{"=IFERROR(1+1,0)", 2.0}, {"=IFERROR(1/0,99)", 99.0}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected));
    }
}

TEST_CASE("XLFormulaEngine – Text functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("String returning text functions")
    {
        auto [expr, expected] = GENERATE(table<std::string, std::string>({{"=LEFT(\"Excel\",3)", "Exc"},
                                                                          {"=RIGHT(\"Excel\",2)", "el"},
                                                                          {"=MID(\"OpenXLSX\",5,4)", "XLSX"},
                                                                          {"=UPPER(\"hello\")", "HELLO"},
                                                                          {"=LOWER(\"HELLO\")", "hello"},
                                                                          {"=TRIM(\"  hi  \")", "hi"},
                                                                          {"=CONCATENATE(\"A\",\"B\",\"C\")", "ABC"}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<std::string>() == expected);
    }

    SECTION("Integer returning text functions") { REQUIRE(eng.evaluate("=LEN(\"hello\")").get<int64_t>() == 5); }
}

TEST_CASE("XLFormulaEngine – Info functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;
    auto            resolver = makeMapResolver({{"A1", XLCellValue(42.0)}, {"B1", XLCellValue("hello")}, {"C1", XLCellValue{}}});

    SECTION("ISNUMBER true") { REQUIRE(eng.evaluate("=ISNUMBER(A1)", resolver).get<bool>() == true); }
    SECTION("ISNUMBER false") { REQUIRE(eng.evaluate("=ISNUMBER(B1)", resolver).get<bool>() == false); }
    SECTION("ISBLANK true") { REQUIRE(eng.evaluate("=ISBLANK(C1)", resolver).get<bool>() == true); }
    SECTION("ISBLANK false") { REQUIRE(eng.evaluate("=ISBLANK(A1)", resolver).get<bool>() == false); }
    SECTION("ISERROR true") { REQUIRE(eng.evaluate("=ISERROR(1/0)").get<bool>() == true); }
    SECTION("ISERROR false") { REQUIRE(eng.evaluate("=ISERROR(1+1)").get<bool>() == false); }
    SECTION("ISTEXT true") { REQUIRE(eng.evaluate("=ISTEXT(B1)", resolver).get<bool>() == true); }
}

TEST_CASE("XLFormulaEngine – MATCH", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;
    // Array: A1=10, B1=20, C1=30
    auto resolver = makeMapResolver({{"A1", XLCellValue(10.0)}, {"B1", XLCellValue(20.0)}, {"C1", XLCellValue(30.0)}});
    SECTION("MATCH exact") { REQUIRE(eng.evaluate("=MATCH(20,A1:C1,0)", resolver).get<int64_t>() == 2); }
    SECTION("MATCH not found") { REQUIRE(eng.evaluate("=MATCH(99,A1:C1,0)", resolver).type() == XLValueType::Error); }
}

TEST_CASE("XLFormulaEngine – VLOOKUP", "[XLFormulaEngine]")
{
    //  Table (A1:B3):  1,"apple"  / 2,"banana"  / 3,"cherry"
    //  colIdx=2, exact match
    XLFormulaEngine eng;
    auto            resolver = makeMapResolver({
        {"A1", XLCellValue(1.0)},
        {"B1", XLCellValue("apple")},
        {"A2", XLCellValue(2.0)},
        {"B2", XLCellValue("banana")},
        {"A3", XLCellValue(3.0)},
        {"B3", XLCellValue("cherry")},
    });
    SECTION("VLOOKUP found") { REQUIRE(eng.evaluate("=VLOOKUP(2,A1:B3,2,0)", resolver).get<std::string>() == "banana"); }
    SECTION("VLOOKUP not found") { REQUIRE(eng.evaluate("=VLOOKUP(99,A1:B3,2,0)", resolver).type() == XLValueType::Error); }
}

TEST_CASE("XLFormulaEngine – XLOOKUP", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;
    auto            resolver = makeMapResolver({
        {"A1", XLCellValue("apple")},
        {"B1", XLCellValue(10.0)},
        {"A2", XLCellValue("banana")},
        {"B2", XLCellValue(20.0)},
        {"A3", XLCellValue("cherry")},
        {"B3", XLCellValue(30.0)},

        {"C1", XLCellValue(10.0)},
        {"D1", XLCellValue("ten")},
        {"C2", XLCellValue(20.0)},
        {"D2", XLCellValue("twenty")},
        {"C3", XLCellValue(30.0)},
        {"D3", XLCellValue("thirty")},
        {"C4", XLCellValue(40.0)},
        {"D4", XLCellValue("forty")},

        {"E1", XLCellValue(40.0)},
        {"F1", XLCellValue("forty")},
        {"E2", XLCellValue(30.0)},
        {"F2", XLCellValue("thirty")},
        {"E3", XLCellValue(20.0)},
        {"F3", XLCellValue("twenty")},
        {"E4", XLCellValue(10.0)},
        {"F4", XLCellValue("ten")},

        // Edge Case 3: Case insensitivity test
        {"G1", XLCellValue("APPLE")},
        {"H1", XLCellValue(1.0)},
        {"G2", XLCellValue("BaNaNa")},
        {"H2", XLCellValue(2.0)},

        // Edge case 4: Type mismatch in exact match
        {"I1", XLCellValue(100.0)},
        {"J1", XLCellValue("Numeric")},
        {"I2", XLCellValue("100")},
        {"J2", XLCellValue("String")},
    });

    SECTION("XLOOKUP exact match")
    { REQUIRE(eng.evaluate("=XLOOKUP(\"banana\", A1:A3, B1:B3)", resolver).get<double>() == Catch::Approx(20.0)); }
    SECTION("XLOOKUP not found default")
    { REQUIRE(eng.evaluate("=XLOOKUP(\"date\", A1:A3, B1:B3, \"Missing\")", resolver).get<std::string>() == "Missing"); }
    SECTION("XLOOKUP not found no default")
    { REQUIRE(eng.evaluate("=XLOOKUP(\"date\", A1:A3, B1:B3)", resolver).type() == XLValueType::Error); }

    // Binary search (search_mode 2) - Ascending array
    SECTION("XLOOKUP binary search exact (mode 2)")
    { REQUIRE(eng.evaluate("=XLOOKUP(30, C1:C4, D1:D4, \"Err\", 0, 2)", resolver).get<std::string>() == "thirty"); }
    SECTION("XLOOKUP binary search exact-or-next-smaller (mode 2)")
    { REQUIRE(eng.evaluate("=XLOOKUP(25, C1:C4, D1:D4, \"Err\", -1, 2)", resolver).get<std::string>() == "twenty"); }
    SECTION("XLOOKUP binary search exact-or-next-larger (mode 2)")
    { REQUIRE(eng.evaluate("=XLOOKUP(25, C1:C4, D1:D4, \"Err\", 1, 2)", resolver).get<std::string>() == "thirty"); }

    // Binary search (search_mode -2) - Descending array
    SECTION("XLOOKUP binary search exact (mode -2)")
    { REQUIRE(eng.evaluate("=XLOOKUP(30, E1:E4, F1:F4, \"Err\", 0, -2)", resolver).get<std::string>() == "thirty"); }
    SECTION("XLOOKUP binary search exact-or-next-smaller (mode -2)")
    { REQUIRE(eng.evaluate("=XLOOKUP(25, E1:E4, F1:F4, \"Err\", -1, -2)", resolver).get<std::string>() == "twenty"); }
    SECTION("XLOOKUP binary search exact-or-next-larger (mode -2)")
    { REQUIRE(eng.evaluate("=XLOOKUP(25, E1:E4, F1:F4, \"Err\", 1, -2)", resolver).get<std::string>() == "thirty"); }

    SECTION("XLOOKUP case insensitive")
    {
        REQUIRE(eng.evaluate("=XLOOKUP(\"apple\", G1:G2, H1:H2)", resolver).get<double>() == Catch::Approx(1.0));
        REQUIRE(eng.evaluate("=XLOOKUP(\"banana\", G1:G2, H1:H2)", resolver).get<double>() == Catch::Approx(2.0));
    }

    SECTION("XLOOKUP type mismatch (number vs string)")
    {
        REQUIRE(eng.evaluate("=XLOOKUP(100, I1:I2, J1:J2)", resolver).get<std::string>() == "Numeric");
        REQUIRE(eng.evaluate("=XLOOKUP(\"100\", I1:I2, J1:J2)", resolver).get<std::string>() == "String");
    }

    SECTION("XLOOKUP unequal or missing return array bounds")
    {
        auto res = eng.evaluate("=XLOOKUP(\"banana\", A1:A3, B1:B1)", resolver);
        REQUIRE(res.type() == XLValueType::Error);    // Return array smaller than lookup array throws error
    }

    SECTION("XLOOKUP match mode 1 (Next Larger) edge cases")
    {
        // Next larger than 25 is 30
        REQUIRE(eng.evaluate("=XLOOKUP(25, C1:C4, D1:D4, \"Err\", 1)", resolver).get<std::string>() == "thirty");
        // Next larger than 45 does not exist -> return default
        REQUIRE(eng.evaluate("=XLOOKUP(45, C1:C4, D1:D4, \"Err\", 1)", resolver).get<std::string>() == "Err");
        // Next larger than 45 with NO default -> #N/A
        REQUIRE(eng.evaluate("=XLOOKUP(45, C1:C4, D1:D4, , 1)", resolver).type() == XLValueType::Error);
    }
}

TEST_CASE("XLFormulaEngine – Integration with XLDocument", "[XLFormulaEngine]")
{
    XLDocument doc;
    doc.create("./testXLFormulaEngine_integration.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = 10.0;
    wks.cell("B1").value() = 20.0;
    wks.cell("C1").value() = 30.0;
    wks.cell("A2").value() = std::string("hello");

    XLFormulaEngine eng;
    auto            resolver = XLFormulaEngine::makeResolver(wks);

    SECTION("SUM via worksheet") { REQUIRE(eng.evaluate("=SUM(A1:C1)", resolver).get<double>() == Catch::Approx(60.0)); }
    SECTION("AVERAGE via worksheet") { REQUIRE(eng.evaluate("=AVERAGE(A1:C1)", resolver).get<double>() == Catch::Approx(20.0)); }
    SECTION("Cell arithmetic") { REQUIRE(eng.evaluate("=A1*B1", resolver).get<double>() == Catch::Approx(200.0)); }
    SECTION("ISTEXT on string cell") { REQUIRE(eng.evaluate("=ISTEXT(A2)", resolver).get<bool>() == true); }

    doc.close();
}

// =============================================================================
// New Tests – Date functions
// =============================================================================

TEST_CASE("XLFormulaEngine – Date functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    // DATE(2024,3,15) – Excel serial for 2024-03-15
    // Known: =DATE(2024,3,15) in Excel → 45365
    SECTION("DATE builds serial")
    {
        auto serial = eng.evaluate("=DATE(2024,3,15)").get<double>();
        REQUIRE(serial == Catch::Approx(45366.0));    // Excel serial for 2024-03-15
    }

    SECTION("YEAR from DATE") { REQUIRE(eng.evaluate("=YEAR(DATE(2024,3,15))").get<int64_t>() == 2024); }
    SECTION("MONTH from DATE") { REQUIRE(eng.evaluate("=MONTH(DATE(2024,3,15))").get<int64_t>() == 3); }
    SECTION("DAY from DATE") { REQUIRE(eng.evaluate("=DAY(DATE(2024,3,15))").get<int64_t>() == 15); }

    // 2024-03-15 is a Friday → WEEKDAY mode-1 = 6
    SECTION("WEEKDAY Friday mode 1") { REQUIRE(eng.evaluate("=WEEKDAY(DATE(2024,3,15),1)").get<int64_t>() == 6); }
    // mode-2: Monday=1 … Sunday=7, Friday=5
    SECTION("WEEKDAY Friday mode 2") { REQUIRE(eng.evaluate("=WEEKDAY(DATE(2024,3,15),2)").get<int64_t>() == 5); }

    SECTION("EDATE +1 month")
    {
        // 2024-01-31 + 1 month = 2024-02-29 (leap year clamp)
        REQUIRE(eng.evaluate("=DAY(EDATE(DATE(2024,1,31),1))").get<int64_t>() == 29);
    }

    SECTION("EOMONTH end of Feb 2024") { REQUIRE(eng.evaluate("=DAY(EOMONTH(DATE(2024,1,15),1))").get<int64_t>() == 29); }
}

// =============================================================================
// New Tests – Work-date functions
// =============================================================================

TEST_CASE("XLFormulaEngine – Work date functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    // 2024-01-01 (Monday) + 5 workdays = 2024-01-08 (Monday, serial 45299)
    SECTION("WORKDAY 5 from Monday")
    {
        auto result = eng.evaluate("=WORKDAY(DATE(2024,1,1),5)").get<double>();
        // 2024-01-08
        REQUIRE(result == Catch::Approx(45299.0));
    }

    // NETWORKDAYS from Mon 2024-01-01 to Fri 2024-01-05 = 5
    SECTION("NETWORKDAYS Mon to Fri = 5") { REQUIRE(eng.evaluate("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5))").get<int64_t>() == 5); }

    // NETWORKDAYS reversed (end before start) = negative
    SECTION("NETWORKDAYS reversed") { REQUIRE(eng.evaluate("=NETWORKDAYS(DATE(2024,1,5),DATE(2024,1,1))").get<int64_t>() == -5); }
}

// =============================================================================
// New Tests – Financial functions
// =============================================================================

TEST_CASE("XLFormulaEngine – Financial functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    // PMT(0.5%, 360, -100000) ≈ 599.5505 (monthly mortgage)
    SECTION("PMT mortgage") { REQUIRE(eng.evaluate("=PMT(0.005,360,-100000)").get<double>() == Catch::Approx(599.5505).epsilon(0.001)); }

    // FV(5%, 10, -100, 0) ≈ 1257.79
    SECTION("FV annuity") { REQUIRE(eng.evaluate("=FV(0.05,10,-100)").get<double>() == Catch::Approx(1257.789).epsilon(0.001)); }

    // PV(5%, 10, -100, 0) ≈ 772.17
    SECTION("PV annuity") { REQUIRE(eng.evaluate("=PV(0.05,10,-100)").get<double>() == Catch::Approx(772.1735).epsilon(0.001)); }

    // NPV(10%, {100, 200, 300}) = 100/1.1 + 200/1.21 + 300/1.331 ≈ 481.59
    SECTION("NPV three cashflows")
    {
        auto resolver = makeMapResolver({{"A1", XLCellValue(100.0)}, {"A2", XLCellValue(200.0)}, {"A3", XLCellValue(300.0)}});
        REQUIRE(eng.evaluate("=NPV(0.1,A1:A3)", resolver).get<double>() == Catch::Approx(481.5942).epsilon(0.001));
    }
}

// =============================================================================
// New Tests – Math extended
// =============================================================================

TEST_CASE("XLFormulaEngine – Math extended", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    // SUMPRODUCT with resolver
    SECTION("SUMPRODUCT two ranges")
    {
        auto resolver = makeMapResolver({
            {"A1", XLCellValue(1.0)},
            {"A2", XLCellValue(2.0)},
            {"A3", XLCellValue(3.0)},
            {"B1", XLCellValue(4.0)},
            {"B2", XLCellValue(5.0)},
            {"B3", XLCellValue(6.0)},
        });
        // 1*4 + 2*5 + 3*6 = 32
        REQUIRE(eng.evaluate("=SUMPRODUCT(A1:A3,B1:B3)", resolver).get<double>() == Catch::Approx(32.0));
    }

    SECTION("Data-driven extended math evaluation")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>(
            {{"=CEILING(2.3,0.5)", 2.5}, {"=FLOOR(2.7,0.5)", 2.5}, {"=LOG(8,2)", 3.0}, {"=LOG(1000,10)", 3.0}, {"=EXP(1)", 2.71828}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected).epsilon(0.0001));
    }

    SECTION("Data-driven sign evaluation")
    {
        auto [expr, expected] = GENERATE(table<std::string, int64_t>({{"=SIGN(5)", 1}, {"=SIGN(-3)", -1}, {"=SIGN(0)", 0}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<int64_t>() == expected);
    }
}

// =============================================================================
// New Tests – Text extended
// =============================================================================

TEST_CASE("XLFormulaEngine – Text extended", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("String returning text extended functions")
    {
        auto [expr, expected] = GENERATE(table<std::string, std::string>({{"=SUBSTITUTE(\"aabbaa\",\"a\",\"x\")", "xxbbxx"},
                                                                          {"=SUBSTITUTE(\"aabbaa\",\"a\",\"x\",1)", "xabbaa"},
                                                                          {"=REPLACE(\"OpenXLSX\",5,4,\"calc\")", "Opencalc"},
                                                                          {"=REPT(\"ab\",3)", "ababab"},
                                                                          {"=T(\"hi\")", "hi"},
                                                                          {"=T(42)", ""},
                                                                          {"=TEXTJOIN(\"-\",TRUE,\"A\",\"B\",\"C\")", "A-B-C"},
                                                                          {"=TEXTJOIN(\"-\",TRUE,\"A\",\"\",\"C\")", "A-C"},
                                                                          {"=PROPER(\"hello world\")", "Hello World"},
                                                                          {"=CLEAN(\"hello\")", "hello"}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<std::string>() == expected);
    }

    SECTION("Integer/Double returning text extended functions")
    {
        auto [expr, expected] = GENERATE(table<std::string, double>(
            {{"=FIND(\"XL\",\"OpenXLSX\")", 5.0}, {"=SEARCH(\"xl\",\"OpenXLSX\")", 5.0}, {"=VALUE(\"3.14\")", 3.14}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<double>() == Catch::Approx(expected));
    }

    SECTION("Boolean returning text extended functions")
    {
        auto [expr, expected] =
            GENERATE(table<std::string, bool>({{"=EXACT(\"Hello\",\"Hello\")", true}, {"=EXACT(\"hello\",\"Hello\")", false}}));
        INFO("Evaluating: " << expr);
        REQUIRE(eng.evaluate(expr).get<bool>() == expected);
    }

    SECTION("FIND case-sensitive error")
    {
        REQUIRE(eng.evaluate("=FIND(\"xl\",\"OpenXLSX\")").type() == XLValueType::Error);    // not found (case)
    }
}

// =============================================================================
// New Tests – Statistical / Conditional
// =============================================================================

TEST_CASE("XLFormulaEngine – Statistical extended", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    // Data: A1=1, A2=3, A3=5, A4=7, A5=9 in resolver
    auto resolver = makeMapResolver({
        {"A1", XLCellValue(1.0)},
        {"A2", XLCellValue(3.0)},
        {"A3", XLCellValue(5.0)},
        {"A4", XLCellValue(7.0)},
        {"A5", XLCellValue(9.0)},
        {"B1", XLCellValue(10.0)},
        {"B2", XLCellValue(20.0)},
        {"B3", XLCellValue(30.0)},
        {"B4", XLCellValue(40.0)},
        {"B5", XLCellValue(50.0)},
    });

    SECTION("COUNTIF > 3")
    {
        REQUIRE(eng.evaluate("=COUNTIF(A1:A5,\">3\")", resolver).get<int64_t>() == 3);    // 5,7,9
    }
    SECTION("SUMIF > 3")
    {
        REQUIRE(eng.evaluate("=SUMIF(A1:A5,\">3\",B1:B5)", resolver).get<double>() == Catch::Approx(120.0));    // 30+40+50
    }
    SECTION("AVERAGEIF > 3")
    {
        REQUIRE(eng.evaluate("=AVERAGEIF(A1:A5,\">3\",B1:B5)", resolver).get<double>() == Catch::Approx(40.0));    // (30+40+50)/3
    }
    SECTION("RANK descending")
    {
        // Value=5 in [1,3,5,7,9]: rank desc = 3
        REQUIRE(eng.evaluate("=RANK(5,A1:A5,0)", resolver).get<int64_t>() == 3);
    }
    SECTION("LARGE k=2")
    {
        // Sorted desc: 9,7,5,3,1 → k=2 is 7
        REQUIRE(eng.evaluate("=LARGE(A1:A5,2)", resolver).get<double>() == Catch::Approx(7.0));
    }
    SECTION("SMALL k=2")
    {
        // Sorted asc: 1,3,5,7,9 → k=2 is 3
        REQUIRE(eng.evaluate("=SMALL(A1:A5,2)", resolver).get<double>() == Catch::Approx(3.0));
    }
    SECTION("STDEV sample")
    {
        // stddev of {1,3,5,7,9} = sqrt(10) ≈ 3.162
        REQUIRE(eng.evaluate("=STDEV(A1:A5)", resolver).get<double>() == Catch::Approx(3.1623).epsilon(0.001));
    }
    SECTION("VAR sample")
    {
        // variance = 10
        REQUIRE(eng.evaluate("=VAR(A1:A5)", resolver).get<double>() == Catch::Approx(10.0));
    }
    SECTION("MEDIAN odd") { REQUIRE(eng.evaluate("=MEDIAN(A1:A5)", resolver).get<double>() == Catch::Approx(5.0)); }
    SECTION("COUNTBLANK")
    {
        auto res2 = makeMapResolver({{"A1", XLCellValue{}}, {"A2", XLCellValue(1.0)}, {"A3", XLCellValue{}}});
        REQUIRE(eng.evaluate("=COUNTBLANK(A1:A3)", res2).get<int64_t>() == 2);
    }
}

// =============================================================================
// New Tests – Info extended
// =============================================================================

TEST_CASE("XLFormulaEngine – Info extended", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("ISNA true")
    { REQUIRE(eng.evaluate("=ISNA(MATCH(99,A1:A1,0))", makeMapResolver({{"A1", XLCellValue(1.0)}})).get<bool>() == true); }
    SECTION("ISNA false") { REQUIRE(eng.evaluate("=ISNA(1+1)").get<bool>() == false); }
    SECTION("IFNA fallback")
    {
        REQUIRE(eng.evaluate("=IFNA(MATCH(99,A1:A1,0),\"none\")", makeMapResolver({{"A1", XLCellValue(1.0)}})).get<std::string>() ==
                "none");
    }
    SECTION("ISLOGICAL true") { REQUIRE(eng.evaluate("=ISLOGICAL(TRUE)").get<bool>() == true); }
    SECTION("ISLOGICAL false") { REQUIRE(eng.evaluate("=ISLOGICAL(42)").get<bool>() == false); }
}

// =============================================================================
// New Tests – SUMPRODUCT scalar
// =============================================================================

TEST_CASE("XLFormulaEngine – SUMPRODUCT edge cases", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("SUMPRODUCT single array = SUM")
    {
        auto resolver = makeMapResolver({{"A1", XLCellValue(2.0)}, {"A2", XLCellValue(3.0)}, {"A3", XLCellValue(4.0)}});
        REQUIRE(eng.evaluate("=SUMPRODUCT(A1:A3)", resolver).get<double>() == Catch::Approx(9.0));
    }
}

// =============================================================================
// New Tests – Data-driven quirks matching Excel behavior
// =============================================================================
#include <fstream>
#include <sstream>

TEST_CASE("XLFormulaEngine – Data-driven Quirks from CSV", "[XLFormulaEngine][Quirks]")
{
    XLFormulaEngine eng;
    
    // File generated by CI/test runner or manually
    std::string csvPath = "../Tests/Fixtures/formula_quirks.csv";
    std::ifstream file(csvPath);
    
    // Fallback path if running from project root instead of build/
    if (!file.is_open()) {
        csvPath = "Tests/Fixtures/formula_quirks.csv";
        file.open(csvPath);
    }
    
    REQUIRE(file.is_open());

    std::string line;
    // skip header
    std::getline(file, line);

    while (std::getline(file, line)) {
        if (line.empty()) continue;

        std::stringstream ss(line);
        std::string formula, expectedTypeStr, expectedValueStr;

        std::getline(ss, formula, '|');
        std::getline(ss, expectedTypeStr, '|');
        std::getline(ss, expectedValueStr, '|');

        INFO("Evaluating: " << formula);

        auto result = eng.evaluate(formula);

        if (expectedTypeStr == "Number") {
            REQUIRE((result.type() == XLValueType::Float || result.type() == XLValueType::Integer));
            if (result.type() == XLValueType::Integer) {
                REQUIRE(result.get<int64_t>() == std::stoll(expectedValueStr));
            } else {
                REQUIRE(result.get<double>() == Catch::Approx(std::stod(expectedValueStr)));
            }
        } else if (expectedTypeStr == "Boolean") {
            REQUIRE(result.type() == XLValueType::Boolean);
            REQUIRE(result.get<bool>() == (expectedValueStr == "true"));
        } else if (expectedTypeStr == "String") {
            REQUIRE(result.type() == XLValueType::String);
            REQUIRE(result.get<std::string>() == expectedValueStr);
        } else if (expectedTypeStr == "Error") {
            REQUIRE(result.type() == XLValueType::Error);
            // Engine might abstract errors to empty/NaN, but here we just check type is error
        }
    }
}

TEST_CASE("XLFormulaEngine – Newly Implemented Functions", "[XLFormulaEngine][MissingChecklist]")
{
    XLFormulaEngine eng;
    auto resolver = makeMapResolver({
        {"A1", XLCellValue(2.0)}, {"A2", XLCellValue(3.0)}, {"A3", XLCellValue(4.0)},
        {"B1", XLCellValue(4.0)}, {"B2", XLCellValue(5.0)}, {"B3", XLCellValue(6.0)},
        {"C1", XLCellValue("abc")}, {"C2", XLCellValue("A")},
        {"D1", XLCellValue(45366.0)}, // 2024-03-15
        {"D2", XLCellValue(45410.0)}  // 2024-04-28
    });

    SECTION("Math & Aggregation")
    {
        // SUMSQ(2,3,4) = 4 + 9 + 16 = 29
        REQUIRE(eng.evaluate("=SUMSQ(A1:A3)", resolver).get<double>() == Catch::Approx(29.0));
        
        // SUMX2MY2( {2,3,4}, {4,5,6} ) = (4-16) + (9-25) + (16-36) = -12 - 16 - 20 = -48
        REQUIRE(eng.evaluate("=SUMX2MY2(A1:A3, B1:B3)", resolver).get<double>() == Catch::Approx(-48.0));
        
        // SUMX2PY2( {2}, {4} ) = 4 + 16 = 20
        REQUIRE(eng.evaluate("=SUMX2PY2(A1, B1)", resolver).get<double>() == Catch::Approx(20.0));
        
        // SUMXMY2( {2}, {4} ) = (2-4)^2 = 4
        REQUIRE(eng.evaluate("=SUMXMY2(A1, B1)", resolver).get<double>() == Catch::Approx(4.0));
        
        // MROUND, CEILING.MATH, FLOOR.MATH, TRUNC
        REQUIRE(eng.evaluate("=MROUND(10, 3)").get<double>() == Catch::Approx(9.0));
        
        
        REQUIRE(eng.evaluate("=TRUNC(4.9)").get<double>() == Catch::Approx(4.0));
    }

    SECTION("Date & Time Extensions")
    {
        // DAYS360
        REQUIRE(eng.evaluate("=DAYS360(DATE(2024,1,1), DATE(2024,1,30))").get<double>() == Catch::Approx(29.0));
        
        // WEEKNUM (2024-03-15 is 11th week)
        
        
    }

    SECTION("Text & Code")
    {
        
        
    }

    SECTION("Logic & Constants")
    {
        REQUIRE(eng.evaluate("=ISEVEN(4)").get<bool>() == true);
        REQUIRE(eng.evaluate("=ISODD(4)").get<bool>() == false);
        REQUIRE(eng.evaluate("=ISERR(1/0)").get<bool>() == true); // #DIV/0! is an error
    }

    SECTION("Financial Depreciation")
    {
        // SLN(cost, salvage, life) = (10000 - 1000) / 5 = 1800
        
        
        // SYD(cost, salvage, life, per) = (10000 - 1000) * (5 - 1 + 1) / (5*(5+1)/2) = 9000 * 5 / 15 = 3000
        
    }

    SECTION("Advanced Statistics")
    {
        // PEARSON / RSQ
        // Correl( {2,3,4}, {4,5,6} ) = 1.0 (perfectly linear)
        REQUIRE(eng.evaluate("=PEARSON(A1:A3, B1:B3)", resolver).get<double>() == Catch::Approx(1.0));
        REQUIRE(eng.evaluate("=RSQ(A1:A3, B1:B3)", resolver).get<double>() == Catch::Approx(1.0));
        
        // PERMUT (5 pick 2 = 20)
        REQUIRE(eng.evaluate("=PERMUT(5, 2)").get<double>() == Catch::Approx(20.0));

        // Newly added functions
        REQUIRE(eng.evaluate("=CORREL(A1:A3, B1:B3)", resolver).get<double>() == Catch::Approx(1.0));
        REQUIRE(eng.evaluate("=COVARIANCE.P(A1:A3, B1:B3)", resolver).get<double>() == Catch::Approx(0.6666667));
        REQUIRE(eng.evaluate("=COVARIANCE.S(A1:A3, B1:B3)", resolver).get<double>() == Catch::Approx(1.0));
        REQUIRE(eng.evaluate("=SLOPE(B1:B3, A1:A3)", resolver).get<double>() == Catch::Approx(1.0));
        REQUIRE(eng.evaluate("=INTERCEPT(B1:B3, A1:A3)", resolver).get<double>() == Catch::Approx(2.0));
        
        // PERCENTILE / QUARTILE
        REQUIRE(eng.evaluate("=PERCENTILE.INC(A1:A3, 0.5)", resolver).get<double>() == Catch::Approx(3.0));
        REQUIRE(eng.evaluate("=QUARTILE.INC(A1:A3, 2)", resolver).get<double>() == Catch::Approx(3.0));

        // TRIMMEAN
        REQUIRE(eng.evaluate("=TRIMMEAN(A1:A3, 0.2)", resolver).get<double>() == Catch::Approx(3.0));
        
        // Logic & Text
        REQUIRE(eng.evaluate("=ISNONTEXT(A1)", resolver).get<bool>() == true);
        REQUIRE(eng.evaluate("=ISNONTEXT(C1)", resolver).get<bool>() == false);
        REQUIRE(eng.evaluate("=ISLOGICAL(TRUE)").get<bool>() == true);
    }
}
TEST_CASE("XLFormulaEngine – Bugfixes", "[Bugfixes]")
{
    XLFormulaEngine eng;
    REQUIRE(eng.evaluate("=CEILING.MATH(4.3)").get<double>() == Catch::Approx(5.0));
    REQUIRE(eng.evaluate("=FLOOR.MATH(4.7)").get<double>() == Catch::Approx(4.0));
}
TEST_CASE("XLFormulaEngine – Bugfixes 2", "[Bugfixes2]")
{
    XLFormulaEngine eng;
    auto val = eng.evaluate("=SLN(10000, 1000, 5)");
    REQUIRE(val.type() == XLValueType::Float);
    REQUIRE(val.get<double>() == Catch::Approx(1800.0));
}
TEST_CASE("XLFormulaEngine – Bugfixes 3", "[Bugfixes3]")
{
    XLFormulaEngine eng;
    REQUIRE(eng.evaluate("=CEILING.MATH(4.3)").get<double>() == Catch::Approx(5.0));
    REQUIRE(eng.evaluate("=FLOOR.MATH(4.7)").get<double>() == Catch::Approx(4.0));
    REQUIRE(eng.evaluate("=ISOWEEKNUM(DATE(2024,3,15))").get<double>() == 11.0);
    REQUIRE(eng.evaluate("=WEEKNUM(DATE(2024,3,15))").get<double>() == 11.0);
    REQUIRE(eng.evaluate("=CHAR(65)").get<std::string>() == "A");
    REQUIRE(eng.evaluate("=CODE(\"A\")").get<double>() == 65);
    REQUIRE(eng.evaluate("=SLN(10000, 1000, 5)").get<double>() == Catch::Approx(1800.0));
    REQUIRE(eng.evaluate("=SYD(10000, 1000, 5, 1)").get<double>() == Catch::Approx(3000.0));
}

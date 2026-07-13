#include <XLCellValue.hpp>
#include <XLFormulaUtils.hpp>
#include <catch2/catch_all.hpp>
#include <cmath>
#include <limits>

using namespace OpenXLSX;
using namespace OpenXLSX::formula;

TEST_CASE("formula::toDouble conversions", "[XLFormulaUtils]")
{
    REQUIRE(toDouble(XLCellValue(3)) == Catch::Approx(3.0));
    REQUIRE(toDouble(XLCellValue(2.5)) == Catch::Approx(2.5));
    REQUIRE(toDouble(XLCellValue(true)) == Catch::Approx(1.0));
    REQUIRE(toDouble(XLCellValue(false)) == Catch::Approx(0.0));
    REQUIRE(std::isnan(toDouble(XLCellValue("text"))));
    REQUIRE(std::isnan(toDouble(XLCellValue{})));
}

TEST_CASE("formula::type predicates", "[XLFormulaUtils]")
{
    REQUIRE(isNumeric(XLCellValue(1)));
    REQUIRE(isNumeric(XLCellValue(1.0)));
    REQUIRE(isNumeric(XLCellValue(true)));
    REQUIRE_FALSE(isNumeric(XLCellValue("x")));
    REQUIRE(isEmpty(XLCellValue{}));
    REQUIRE_FALSE(isEmpty(XLCellValue(0)));

    XLCellValue err;
    err.setError("#VALUE!");
    REQUIRE(isError(err));
}

TEST_CASE("formula::makeError and aliases", "[XLFormulaUtils]")
{
    REQUIRE(makeError("#VALUE!").type() == XLValueType::Error);
    REQUIRE(errValue().getString() == "#VALUE!");
    REQUIRE(errDiv0().getString() == "#DIV/0!");
    REQUIRE(errNA().getString() == "#N/A");
    REQUIRE(errNum().getString() == "#NUM!");
    REQUIRE(errRef().getString() == "#REF!");
    REQUIRE(errName().getString() == "#NAME?");

    // Compatibility re-exports in OpenXLSX
    REQUIRE(OpenXLSX::errValue().getString() == "#VALUE!");
}

TEST_CASE("formula::strTrim and toString", "[XLFormulaUtils]")
{
    REQUIRE(strTrim("  ab c  ") == "ab c");
    REQUIRE(strTrim("   ") == "");
    REQUIRE(toString(XLCellValue(42)) == "42");
    REQUIRE(toString(XLCellValue(true)) == "TRUE");
    REQUIRE(toString(XLCellValue("hi")) == "hi");
}

TEST_CASE("formula::coerce and numericsOrError", "[XLFormulaUtils][PhaseA]")
{
    double d = 0.0;
    REQUIRE(tryParseNumericString("3.14", d));
    REQUIRE(d == Catch::Approx(3.14));
    REQUIRE_FALSE(tryParseNumericString("3.14x", d));

    REQUIRE(toDouble(coerceToNumber(XLCellValue(std::string("7")), true)) == Catch::Approx(7.0));
    REQUIRE(toDouble(coerceToNumber(XLCellValue{}, true)) == Catch::Approx(0.0));

    XLCellValue err;
    err.setError("#N/A");
    XLFormulaArg arg({XLCellValue(1.0), err}, 1, 2);
    XLCellValue  out;
    REQUIRE(numericsOrError(arg, out).empty());
    REQUIRE(isError(out));
}

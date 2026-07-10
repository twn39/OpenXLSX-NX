#include <OpenXLSX.hpp>
#include <XLEvaluationContext.hpp>
#include <XLFormulaEngine.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
#include <filesystem>

using namespace OpenXLSX;
TEST_CASE("XLMapEvaluationContext basic lookup", "[XLEvaluationContext]")
{
    XLMapEvaluationContext ctx;
    ctx.set("A1", XLCellValue(10.0));
    ctx.set("B1", XLCellValue(5.0));

    REQUIRE(ctx.cellValue("A1").get<double>() == Catch::Approx(10.0));
    REQUIRE(ctx.cellValue("B1").get<double>() == Catch::Approx(5.0));
    REQUIRE(ctx.cellValue("Z9").type() == XLValueType::Empty);
}

TEST_CASE("XLFormulaEngine evaluate via XLEvaluationContext", "[XLEvaluationContext][XLFormulaEngine]")
{
    XLFormulaEngine eng;
    XLMapEvaluationContext ctx({{"A1", XLCellValue(1.0)}, {"B1", XLCellValue(2.0)}, {"C1", XLCellValue(3.0)}});

    auto resolver = XLFormulaEngine::makeResolver(ctx);
    REQUIRE(eng.evaluate("=A1+B1", resolver).get<double>() == Catch::Approx(3.0));
    REQUIRE(eng.evaluate("=SUM(A1:C1)", resolver).get<double>() == Catch::Approx(6.0));
    REQUIRE(eng.evaluate("=A1*10", resolver).get<double>() == Catch::Approx(10.0));
}

TEST_CASE("XLWorksheetEvaluationContext live sheet", "[XLEvaluationContext]")
{
    const std::string path = TestHelpers::getUniqueFilename("eval_ctx") + ".xlsx";
    {
        XLDocument doc;
        doc.create(path, XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        wks.cell("A1").value() = 42.0;
        wks.cell("B1").value() = 8.0;
        doc.workbook().addWorksheet("Data");
        doc.workbook().worksheet("Data").cell("A1").value() = 100.0;
        doc.save();
        doc.close();
    }

    XLDocument doc;
    doc.open(path);
    auto wks = doc.workbook().worksheet("Sheet1");

    XLWorksheetEvaluationContext ctx(wks);
    REQUIRE(ctx.cellValue("A1").get<double>() == Catch::Approx(42.0));
    REQUIRE(ctx.cellValue("B1").get<double>() == Catch::Approx(8.0));

    XLFormulaEngine eng;
    auto            resolver = XLFormulaEngine::makeResolver(ctx);
    REQUIRE(eng.evaluate("=A1+B1", resolver).get<double>() == Catch::Approx(50.0));

    // Cross-sheet reference through parent workbook
    REQUIRE(ctx.cellValue("Data!A1").get<double>() == Catch::Approx(100.0));
    REQUIRE(eng.evaluate("=Data!A1+A1", resolver).get<double>() == Catch::Approx(142.0));

    // Backward-compatible makeResolver(worksheet)
    auto wksResolver = XLFormulaEngine::makeResolver(wks);
    REQUIRE(eng.evaluate("=A1", wksResolver).get<double>() == Catch::Approx(42.0));

    doc.close();
    std::error_code ec;
    std::filesystem::remove(path, ec);
}

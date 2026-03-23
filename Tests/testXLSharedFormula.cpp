#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <iostream>
#include <string>

using namespace OpenXLSX;

TEST_CASE("Shared Formula Lexer Test", "[XLFormula]")
{
    SECTION("Tricky string tokenization") {
        XLDocument doc;
        doc.create("./TrickyFormulaTest.xlsx", XLForceOverwrite);
        
        doc.workbook().addWorksheet("My Sheet");
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value() = 1;
        wks.cell("A2").value() = 2;
        wks.cell("A3").formula() = "=SUM(A1:A2)";
        wks.cell("B1").formula() = "=IF(A1=\"Yes\", B1+C1, \"$A$1\")"; 
        wks.cell("C1").formula() = "'My Sheet'!A1+5"; 
        wks.cell("D1").formula() = "IF(A2=2, \"He said \"\"Hello\"\" to $C$1\", 0)";

        doc.save();

        XLDocument doc2;
        doc2.open("./TrickyFormulaTest.xlsx");
        auto wks2 = doc2.workbook().worksheet("Sheet1");

        // Validate the formulas remain uncorrupted upon retrieval
        REQUIRE(wks2.cell("A3").formula().get() == "=SUM(A1:A2)");
        REQUIRE(wks2.cell("B1").formula().get() == "=IF(A1=\"Yes\", B1+C1, \"$A$1\")");
        REQUIRE(wks2.cell("C1").formula().get() == "'My Sheet'!A1+5");
        REQUIRE(wks2.cell("D1").formula().get() == "IF(A2=2, \"He said \"\"Hello\"\" to $C$1\", 0)");
    }
}

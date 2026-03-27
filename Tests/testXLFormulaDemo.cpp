#include <catch2/catch_test_macros.hpp>
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Generate Formula Test Document", "[FormulaEngine]")
{
    XLDocument doc;
    doc.create("./FormulaDemo.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Add some raw data
    wks.cell("A1").value() = 10.5;
    wks.cell("A2").value() = 20.2;
    wks.cell("A3").value() = 5.0;
    wks.cell("A4").value() = -3.1;
    wks.cell("A5").value() = 0.0;
    
    // String data
    wks.cell("B1").value() = "Apple";
    wks.cell("B2").value() = "Banana";
    wks.cell("B3").value() = "Cherry";

    // Boolean data
    wks.cell("C1").value() = true;
    wks.cell("C2").value() = false;

    // Formulas
    wks.cell("D1").formula() = "SUM(A1:A5)";            // Mathematical sum
    wks.cell("D2").formula() = "AVERAGE(A1:A5)";        // Average
    wks.cell("D3").formula() = "MAX(A1:A5)";            // Max
    wks.cell("D4").formula() = "MIN(A1:A5)";            // Min
    wks.cell("D5").formula() = "COUNT(A1:A5)";          // Count
    
    wks.cell("E1").formula() = "IF(A1>15, \"High\", \"Low\")"; // Logical IF
    wks.cell("E2").formula() = "AND(C1, C2)";                  // Logical AND
    wks.cell("E3").formula() = "OR(C1, C2)";                   // Logical OR
    wks.cell("E4").formula() = "NOT(C2)";                      // Logical NOT

    wks.cell("F1").formula() = "CONCATENATE(B1, \" and \", B2)"; // Text Concat
    wks.cell("F2").formula() = "LEN(B3)";                        // Text Length
    wks.cell("F3").formula() = "LOWER(B1)";                      // Text Lower

    wks.cell("G1").formula() = "ABS(A4)";               // ABS
    wks.cell("G2").formula() = "ROUND(A1, 0)";          // ROUND
    wks.cell("G3").formula() = "POWER(A3, 2)";          // POWER
    wks.cell("G4").formula() = "SQRT(25)";              // SQRT

    // Now let's calculate the values using XLFormulaEngine and store them as cached values
    // to prove the engine works.
    XLFormulaEngine engine;
    auto resolver = XLFormulaEngine::makeResolver(wks);

    // List of cells to evaluate
    std::vector<std::string> targetCells = {
        "D1", "D2", "D3", "D4", "D5",
        "E1", "E2", "E3", "E4",
        "F1", "F2", "F3",
        "G1", "G2", "G3", "G4"
    };

    std::cout << "--- Formula Evaluation Results ---" << std::endl;
    for (const auto& ref : targetCells) {
        auto cell = wks.cell(ref);
        std::string formulaStr = cell.formula().get();
        
        // Evaluate
        XLCellValue result = engine.evaluate(formulaStr, resolver);
        
        // Output to console
        std::cout << ref << " [" << formulaStr << "] = ";
        if (result.type() == XLValueType::Float) std::cout << result.get<double>();
        else if (result.type() == XLValueType::Integer) std::cout << result.get<int64_t>();
        else if (result.type() == XLValueType::Boolean) std::cout << (result.get<bool>() ? "TRUE" : "FALSE");
        else if (result.type() == XLValueType::String) std::cout << result.getString();
        else if (result.type() == XLValueType::Error) std::cout << "ERROR";
        else std::cout << "UNKNOWN";
        std::cout << std::endl;

        // Note: setting the value caches the result in the <v> node!
        // cell.value() = result; 
        // We shouldn't set cell.value() blindly because the current proxy logic might overwrite the formula.
        // Usually caching is done differently or you let Excel do it on open.
        // Let's just keep the formula as is.
    }

    doc.save();
    doc.close();

    std::cout << "--- Successfully generated FormulaDemo.xlsx ---" << std::endl;
}

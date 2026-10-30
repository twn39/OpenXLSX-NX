/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#include "TestHelpers.hpp"
#include <OpenXLSX.hpp>
#include <XLFormulaEngine.hpp>
#include <XLXmlSchema.hpp>
#include <catch2/catch_all.hpp>

#include <filesystem>
#include <string>

using namespace OpenXLSX;

TEST_CASE("Parse-time function binding in formula engine", "[ModernArch][FormulaBinding]")
{
    // 1. Verify standard built-in functions receive pre-bound funcImpl at parse time
    SECTION("Built-in function pre-binding")
    {
        auto tokens = XLFormulaLexer::tokenize("SUM(1, 2, 3)");
        auto ast    = XLFormulaParser::parse(tokens);
        REQUIRE(ast != nullptr);
        REQUIRE(ast->kind == XLNodeKind::FuncCall);
        REQUIRE(ast->text == "SUM");
        REQUIRE(ast->funcImpl != nullptr);
        REQUIRE_FALSE(ast->isRefFunction);

        XLFormulaEngine engine;
        XLCellValue     result = engine.evaluate("=SUM(10, 20, 30)");
        REQUIRE(result.get<double>() == 60.0);
    }

    // 2. Verify reference-returning functions are flagged at parse time
    SECTION("Reference function flag")
    {
        auto tokensIndirect = XLFormulaLexer::tokenize("INDIRECT(\"A1\")");
        auto astIndirect    = XLFormulaParser::parse(tokensIndirect);
        REQUIRE(astIndirect != nullptr);
        REQUIRE(astIndirect->isRefFunction == true);

        auto tokensOffset = XLFormulaLexer::tokenize("OFFSET(A1, 1, 1)");
        auto astOffset    = XLFormulaParser::parse(tokensOffset);
        REQUIRE(astOffset != nullptr);
        REQUIRE(astOffset->isRefFunction == true);
    }

    // 3. Unknown functions evaluate gracefully to #NAME?
    SECTION("Unknown function evaluation")
    {
        auto tokens = XLFormulaLexer::tokenize("NON_EXISTENT_FUNCTION(123)");
        auto ast    = XLFormulaParser::parse(tokens);
        REQUIRE(ast != nullptr);
        REQUIRE(ast->funcImpl == nullptr);

        XLFormulaEngine engine;
        XLCellValue     result = engine.evaluate("=NON_EXISTENT_FUNCTION(123)");
        REQUIRE(result.type() == XLValueType::Error);
        REQUIRE(result.getString() == "#NAME?");
    }
}

TEST_CASE("Data-driven schema properties value semantics", "[ModernArch][SchemaOps]")
{
    // Test XLRowProperties POD operations
    SECTION("RowProperties read and write")
    {
        XMLDocument doc;
        auto        rowNode = doc.append_child("row");

        XLRowProperties props;
        props.rowIndex     = 5;
        props.height       = 24.5f;
        props.customHeight = true;
        props.hidden       = true;
        props.styleIndex   = 3;

        XmlOps::writeRowProps(rowNode, props);

        auto readBack = XmlOps::readRowProps(rowNode);
        REQUIRE(readBack.rowIndex == 5);
        REQUIRE(readBack.height == Catch::Approx(24.5f));
        REQUIRE(readBack.customHeight == true);
        REQUIRE(readBack.hidden == true);
        REQUIRE(readBack.styleIndex == 3);
    }

    // Test XLTableProperties POD operations
    SECTION("TableProperties read and write")
    {
        XMLDocument doc;
        auto        tableNode = doc.append_child("table");

        XLTableProperties props;
        props.id             = 101;
        props.name           = "SalesData";
        props.displayName    = "SalesData_Table";
        props.rangeRef       = "A1:D10";
        props.headerRowCount = 1;
        props.totalsRowCount = 1;
        props.totalsRowShown = true;

        XmlOps::writeTableProps(tableNode, props);

        auto readBack = XmlOps::readTableProps(tableNode);
        REQUIRE(readBack.id == 101);
        REQUIRE(readBack.name == "SalesData");
        REQUIRE(readBack.displayName == "SalesData_Table");
        REQUIRE(readBack.rangeRef == "A1:D10");
        REQUIRE(readBack.headerRowCount == 1);
        REQUIRE(readBack.totalsRowCount == 1);
        REQUIRE(readBack.totalsRowShown == true);
    }
}

TEST_CASE("Cross-document worksheet cloning and style/SST migration", "[ModernArch][CrossDocClone]")
{
    const std::string pathA = TestHelpers::getUniqueFilename("docA_src") + ".xlsx";
    const std::string pathB = TestHelpers::getUniqueFilename("docB_dst") + ".xlsx";

    // Step 1: Create Document A with styled cells, formulas, shared strings, and merged cells
    {
        XLDocument docA;
        docA.create(pathA, XLForceOverwrite);
        auto wksA = docA.workbook().worksheet("Sheet1");

        // Cell values with shared strings and numbers
        wksA.cell("A1").value()   = "Product Name";
        wksA.cell("B1").value()   = "Price";
        wksA.cell("A2").value()   = "Super Widget";
        wksA.cell("B2").value()   = 99.5;
        wksA.cell("A3").value()   = "Mega Gadget";
        wksA.cell("B3").value()   = 149.0;
        wksA.cell("A4").value()   = "Total";
        wksA.cell("B4").formula() = "SUM(B2:B3)";

        // Apply distinct formatting to Header A1
        XLStyle headerStyle;
        headerStyle.font.bold    = true;
        headerStyle.font.name    = "Arial";
        headerStyle.fill.pattern = XLPatternSolid;
        headerStyle.fill.fgColor = XLColor("CCE5FF");
        wksA.cell("A1").setStyle(headerStyle);
        wksA.cell("B1").setStyle(headerStyle);

        // Merge cells C1:D1
        wksA.mergeCells("C1:D1");
        wksA.cell("C1").value() = "Merged Header";

        docA.save();
        docA.close();
    }

    // Step 2: Open Document A and Document B, then clone Sheet1 from Doc A into Doc B
    {
        XLDocument docA;
        docA.open(pathA);

        XLDocument docB;
        docB.create(pathB, XLForceOverwrite);
        auto srcSheet = docA.workbook().worksheet("Sheet1");

        // Cross-document clone!
        auto importedSheet = docB.workbook().cloneSheet(srcSheet, "ImportedData");
        REQUIRE(importedSheet.name() == "ImportedData");

        // Verify values transferred accurately
        REQUIRE(importedSheet.cell("A1").value().get<std::string>() == "Product Name");
        REQUIRE(importedSheet.cell("B1").value().get<std::string>() == "Price");
        REQUIRE(importedSheet.cell("A2").value().get<std::string>() == "Super Widget");
        REQUIRE(importedSheet.cell("B2").value().get<double>() == 99.5);
        REQUIRE(importedSheet.cell("A3").value().get<std::string>() == "Mega Gadget");
        REQUIRE(importedSheet.cell("B3").value().get<double>() == 149.0);
        REQUIRE(importedSheet.cell("A4").value().get<std::string>() == "Total");
        REQUIRE(importedSheet.cell("B4").formula().get() == "SUM(B2:B3)");
        REQUIRE(importedSheet.cell("C1").value().get<std::string>() == "Merged Header");

        // Also test intra-document clone on docB
        auto intraSheet = docB.workbook().cloneSheet(importedSheet, "IntraCopy");
        REQUIRE(intraSheet.name() == "IntraCopy");
        REQUIRE(intraSheet.cell("A2").value().get<std::string>() == "Super Widget");

        docB.save();
        docB.close();
        docA.close();
    }

    // Step 3: Reopen Document B from disk and verify persistence and integrity
    {
        XLDocument docB;
        docB.open(pathB);
        REQUIRE(docB.workbook().sheetExists("ImportedData"));
        REQUIRE(docB.workbook().sheetExists("IntraCopy"));

        auto importedSheet = docB.workbook().worksheet("ImportedData");
        REQUIRE(importedSheet.cell("A1").value().get<std::string>() == "Product Name");
        REQUIRE(importedSheet.cell("B2").value().get<double>() == 99.5);
        REQUIRE(importedSheet.cell("B4").formula().get() == "SUM(B2:B3)");
        REQUIRE(importedSheet.cell("C1").value().get<std::string>() == "Merged Header");

        REQUIRE_NOTHROW(docB.validatePackageInvariants());
        docB.close();
    }

    std::error_code ec;
    std::filesystem::remove(pathA, ec);
    std::filesystem::remove(pathB, ec);
}

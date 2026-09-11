/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#include <catch2/catch_test_macros.hpp>
#include <OpenXLSX.hpp>
#include "headers/IXLCellProvider.hpp"
#include "headers/XLXmlSchema.hpp"
#include "headers/XLFormulaEngine.hpp"

#include <map>
#include <string>
#include <vector>

using namespace OpenXLSX;

namespace {
    /**
     * @brief In-memory mock cell provider completely detached from XML or file systems.
     * Demonstrates that IXLCellProvider enables pure virtual/mock spreadsheet testing.
     */
    class MockCellProvider : public IXLCellProvider
    {
    public:
        explicit MockCellProvider(std::string name = "MockSheet") : m_name(std::move(name)) {}

        void set(uint32_t row, uint16_t col, XLCellValue val, std::string formula = "")
        {
            m_cells[{row, col}] = std::move(val);
            if (!formula.empty()) m_formulas[{row, col}] = std::move(formula);
        }

        [[nodiscard]] XLCellValue getCellValue(uint32_t row, uint16_t col) const override
        {
            auto it = m_cells.find({row, col});
            if (it != m_cells.end()) return it->second;
            return XLCellValue();
        }

        [[nodiscard]] std::string getCellFormula(uint32_t row, uint16_t col) const override
        {
            auto it = m_formulas.find({row, col});
            if (it != m_formulas.end()) return it->second;
            return "";
        }

        [[nodiscard]] bool hasCell(uint32_t row, uint16_t col) const override
        {
            return m_cells.find({row, col}) != m_cells.end();
        }

        [[nodiscard]] std::string sheetName() const override
        {
            return m_name;
        }

    private:
        std::string m_name;
        std::map<std::pair<uint32_t, uint16_t>, XLCellValue> m_cells;
        std::map<std::pair<uint32_t, uint16_t>, std::string> m_formulas;
    };
}

TEST_CASE("Architecture Decoupling - IXLCellProvider Mock & Formula Integration", "[decoupling]")
{
    MockCellProvider mock("Sales");
    mock.set(1, 1, 100.0); // A1
    mock.set(1, 2, 250.0); // B1
    mock.set(2, 1, 50.0);  // A2
    mock.set(2, 2, 150.0); // B2

    REQUIRE(mock.sheetName() == "Sales");
    REQUIRE(mock.hasCell(1, 1));
    REQUIRE(mock.hasCell(2, 2));
    REQUIRE_FALSE(mock.hasCell(3, 3));
    REQUIRE(mock.getCellValue(1, 1).get<double>() == 100.0);
    REQUIRE(mock.getCellValue(1, 2).get<double>() == 250.0);

    // Test formula engine evaluate using resolver created from IXLCellProvider
    XLFormulaEngine engine;
    auto resolver = XLFormulaEngine::makeResolver(mock);

    auto resAdd = engine.evaluate("=A1 + B1", resolver);
    REQUIRE(resAdd.type() == XLValueType::Float);
    REQUIRE(resAdd.get<double>() == 350.0);

    auto resSum = engine.evaluate("SUM(A1:B2)", resolver);
    REQUIRE(resSum.type() == XLValueType::Float);
    REQUIRE(resSum.get<double>() == 550.0);
}

TEST_CASE("Architecture Decoupling - XLWorksheet implements IXLCellProvider", "[decoupling]")
{
    XLDocument doc;
    doc.create("./test_decoupling_wks.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = 10;
    wks.cell("A2").value() = 20;
    wks.cell("A3").value() = 30;

    const IXLCellProvider& provider = wks;
    REQUIRE(provider.sheetName() == "Sheet1");
    REQUIRE(provider.hasCell(1, 1));
    REQUIRE(provider.hasCell(2, 1));
    REQUIRE(provider.hasCell(3, 1));
    REQUIRE_FALSE(provider.hasCell(4, 1));

    REQUIRE(provider.getCellValue(1, 1).get<int64_t>() == 10);
    REQUIRE(provider.getCellValue(2, 1).get<int64_t>() == 20);
    REQUIRE(provider.getCellValue(3, 1).get<int64_t>() == 30);

    // Test batch fetchRangeValues
    std::vector<XLCellValue> values;
    provider.fetchRangeValues(wks.range("A1:A3"), values);
    REQUIRE(values.size() == 3);
    REQUIRE(values[0].get<int64_t>() == 10);
    REQUIRE(values[1].get<int64_t>() == 20);
    REQUIRE(values[2].get<int64_t>() == 30);

    doc.close();
}

TEST_CASE("Architecture Decoupling - XLXmlSchema Constants & CellNodeView", "[decoupling]")
{
    REQUIRE(std::string_view(XmlTag::Cell) == "c");
    REQUIRE(std::string_view(XmlTag::Row) == "row");
    REQUIRE(std::string_view(XmlTag::Value) == "v");
    REQUIRE(std::string_view(XmlTag::Formula) == "f");
    REQUIRE(std::string_view(XmlAttr::Ref) == "r");
    REQUIRE(std::string_view(XmlAttr::Style) == "s");
    REQUIRE(std::string_view(XmlAttr::Type) == "t");
    REQUIRE(XmlNS::SpreadsheetML.find("spreadsheetml") != std::string_view::npos);

    // Test CellNodeView
    XMLDocument xdoc;
    auto cNode = xdoc.append_child("c");
    cNode.append_attribute("r").set_value("B5");
    cNode.append_attribute("t").set_value("s");
    cNode.append_attribute("s").set_value(42);
    auto vNode = cNode.append_child("v");
    vNode.text().set("7");

    CellNodeView view(cNode);
    REQUIRE_FALSE(view.empty());
    REQUIRE(std::string_view(view.ref()) == "B5");
    REQUIRE(std::string_view(view.type()) == "s");
    REQUIRE(view.styleIndex() == 42);
    REQUIRE(std::string_view(view.value()) == "7");
    REQUIRE_FALSE(view.hasFormula());
}

TEST_CASE("Architecture Decoupling - XLCommand & XLQuery string_view usage", "[decoupling]")
{
    XLQuery q(XLQueryType::QuerySheetName);
    q.setParam("sheetID", std::string("rId1"));
    REQUIRE(q.getParam<std::string>("sheetID") == "rId1");

    XLCommand cmd(XLCommandType::SetSheetName);
    cmd.setParam("sheetID", std::string("rId2"));
    cmd.setParam("newName", std::string("Summary"));
    REQUIRE(cmd.getParam<std::string>("sheetID") == "rId2");
    REQUIRE(cmd.getParam<std::string>("newName") == "Summary");
}

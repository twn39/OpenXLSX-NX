#include <IZipArchive.hpp>
#include <OpenXLSX.hpp>
#include "XLTables.hpp"

#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
#include <pugixml.hpp>

using namespace OpenXLSX;

/**
 * @brief Helper to check the physical order of child nodes in an XML node.
 */
bool checkNodeOrder(const pugi::xml_node& parent, const std::vector<std::string>& expectedOrder)
{
    int lastPos = -1;
    for (const auto& name : expectedOrder) {
        auto node = parent.child(name.c_str());
        if (node.empty()) continue;

        int currentPos = 0;
        for (auto n : parent.children()) {
            if (n == node) break;
            currentPos++;
        }
        if (currentPos < lastPos) return false;
        lastPos = currentPos;
    }
    return true;
}

/**
 * @brief Helper to check attribute sequence in a node.
 */
bool checkAttributeOrder(const pugi::xml_node& node, const std::vector<std::string>& expectedOrder)
{
    int lastPos = -1;
    for (const auto& name : expectedOrder) {
        auto attr = node.attribute(name.c_str());
        if (attr.empty()) continue;

        int currentPos = 0;
        for (auto a : node.attributes()) {
            if (a == attr) break;
            currentPos++;
        }
        if (currentPos < lastPos) return false;
        lastPos = currentPos;
    }
    return true;
}

TEST_CASE("XLTablesOOXMLComplianceStability", "[XLTables]")
{
    const std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();

    SECTION("Feature 1, 2, 3, 6, 8 Integration & OOXML Check")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // Setup data for Feature 2 & 6
            wks.cell("A1").value() = "ID";
            wks.cell("B1").value() = "Value";
            wks.cell("A2").value() = 1;
            wks.cell("B2").value() = 100;

            // Add Table (Feature 1)
            auto tbl = wks.tables().add("Table1", "A1:B3");

            // Auto Columns (Feature 2)
            tbl.createColumnsFromRange(wks);

            // Styles (Feature 3)
            tbl.setStyleName("TableStyleMedium2");

            // Totals row (Feature 6 - Simplified stable parts)
            // Note: Feature 5 (toggle) is unstable, but basic property setting is okay
            tbl.setShowTotalsRow(true);
            tbl.column("Value").setTotalsRowFunction(XLTotalsRowFunction::Sum);
            wks.cell("B3").value() = 100;    // Manual sync for worksheet data

            doc.save();
        }

        // --- OOXML Compliance Verification ---
        {
            XLDocument doc;
            doc.open(filename);

            // Access internal XML via archive to check physical layout
            auto tableXml = doc.archive().getEntry("xl/tables/table1.xml");
            REQUIRE_FALSE(tableXml.empty());

            pugi::xml_document xmlDoc;
            xmlDoc.load_string(tableXml.c_str());
            auto root = xmlDoc.document_element();

            // 1. Check Root Attribute Order (CRITICAL)
            // Sequence: xmlns, id, name, displayName, ref, ...
            std::vector<std::string> attrOrder = {"xmlns", "id", "name", "displayName", "ref"};
            REQUIRE(checkAttributeOrder(root, attrOrder));

            // 2. Check Child Node Order (CRITICAL)
            // Sequence: autoFilter, tableColumns, tableStyleInfo
            std::vector<std::string> nodeOrder = {"autoFilter", "tableColumns", "tableStyleInfo"};
            REQUIRE(checkNodeOrder(root, nodeOrder));

            // 3. Check tableColumns count matches child nodes
            auto columns = root.child("tableColumns");
            REQUIRE(columns.attribute("count").as_uint() == 2);
            int colCount = 0;
            for (auto col : columns.children("tableColumn")) colCount++;
            REQUIRE(colCount == 2);

            doc.close();
        }
    }

    SECTION("Multiple Tables Integrity (Feature 8)")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            wks.tables().add("T1", "A1:B2");
            wks.tables().add("T2", "D1:E2");

            REQUIRE(wks.tables().count() == 2);
            doc.save();
        }

        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            REQUIRE(wks.tables().count() == 2);
            REQUIRE(doc.archive().hasEntry("xl/tables/table1.xml"));
            REQUIRE(doc.archive().hasEntry("xl/tables/table2.xml"));
            doc.close();
        }
    }
}

TEST_CASE("XLTablesErgonomicAddAPI", "[XLTables][Range]")
{
    const std::string filename = OpenXLSX::TestHelpers::getUniqueFilename();

    SECTION("Adding table via XLCellRange instead of magic string")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            wks.cell("A1").value() = "Column 1";
            wks.cell("B1").value() = "Column 2";
            wks.cell("A2").value() = 1;
            wks.cell("B2").value() = 2;

            // Use the new type-safe overload
            auto range = wks.range("A1:B2");
            auto tbl   = wks.tables().add("MyTable", range);

            // This implicitly relies on the getString() safety patch
            tbl.createColumnsFromRange(wks);

            doc.save();
            doc.close();
        }

        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto tbl = wks.tables().table("MyTable");
            REQUIRE(tbl.rangeReference() == "A1:B2");
            REQUIRE(tbl.column("Column 1").id() == 1);
            REQUIRE(tbl.column("Column 2").id() == 2);
        }
    }
}

#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <ctime>

using namespace OpenXLSX;

TEST_CASE("ComprehensiveIntegrationTests", "[Integration]")
{
    const std::string filename = "ComprehensiveIntegration.xlsx";

    SECTION("Data Types and Formulas Round-trip")
    {
        // 1. Write data
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            wks.cell("A1").value() = 42;
            wks.cell("A2").value() = 3.14159;
            wks.cell("A3").value() = "OpenXLSX";
            wks.cell("A4").value() = true;

            std::tm tm             = {};
            tm.tm_year             = 126;    // 2026
            tm.tm_mon              = 1;      // Feb
            tm.tm_mday             = 12;
            wks.cell("A5").value() = XLDateTime(tm);

            wks.cell("B1").formula() = "A1*2";

            doc.save();
            doc.close();
        }

        // 2. Read and verify
        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");

            REQUIRE(wks.cell("A1").value().get<int>() == 42);
            REQUIRE(wks.cell("A2").value().get<double>() == Catch::Approx(3.14159));
            REQUIRE(wks.cell("A3").value().get<std::string>() == "OpenXLSX");
            REQUIRE(wks.cell("A4").value().get<bool>() == true);

            // Verify Date
            auto dt = wks.cell("A5").value().get<XLDateTime>();
            REQUIRE(dt.tm().tm_year == 126);
            REQUIRE(dt.tm().tm_mon == 1);
            REQUIRE(dt.tm().tm_mday == 12);

            // Verify Formula
            REQUIRE(wks.cell("B1").formula().get() == "A1*2");

            doc.close();
        }
    }

    SECTION("Layout and Visibility Round-trip")
    {
        // 1. Write settings
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wb  = doc.workbook();
            auto wks = wb.worksheet("Sheet1");

            wks.column(1).setWidth(30.0f);
            wks.column(2).setHidden(true);

            wks.row(1).setHeight(40.0f);
            wks.row(2).setHidden(true);

            wb.addWorksheet("HiddenSheet");
            wb.worksheet("HiddenSheet").setVisibility(XLSheetState::Hidden);

            wks.mergeCells("C1:D2");

            doc.save();
            doc.close();
        }

        // 2. Read and verify
        {
            XLDocument doc;
            doc.open(filename);
            auto wb  = doc.workbook();
            auto wks = wb.worksheet("Sheet1");

            REQUIRE(wks.column(1).width() == Catch::Approx(30.0f));
            REQUIRE(wks.column(2).isHidden() == true);

            REQUIRE(wks.row(1).height() == Catch::Approx(40.0f));
            REQUIRE(wks.row(2).isHidden() == true);

            REQUIRE(wb.worksheet("HiddenSheet").visibility() == XLSheetState::Hidden);
            REQUIRE(wks.merges().mergeExists("C1:D2") == true);

            doc.close();
        }
    }

    SECTION("Advanced Styles Round-trip")
    {
        // 1. Write styles
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks    = doc.workbook().worksheet("Sheet1");
            auto styles = doc.styles();

            // Alignment
            size_t alignXf = styles.cellFormats().create();
            auto   xf      = styles.cellFormats()[alignXf];
            auto   align   = xf.alignment(true);    // Must use true to create the node
            align.setHorizontal(XLAlignmentStyle::XLAlignCenter);
            align.setVertical(XLAlignmentStyle::XLAlignTop);
            align.setWrapText(true);
            xf.setApplyAlignment(true);

            wks.cell("E1").value() = "Aligned";
            wks.cell("E1").setCellFormat(alignXf);

            // Border
            size_t borderIdx = styles.borders().create();
            auto   border    = styles.borders()[borderIdx];
            border.setBottom(XLLineStyleThick, XLColor(0, 0, 255));

            size_t borderXf = styles.cellFormats().create();
            styles.cellFormats()[borderXf].setBorderIndex(borderIdx);
            styles.cellFormats()[borderXf].setApplyBorder(true);

            wks.cell("F1").setCellFormat(borderXf);

            doc.save();
            doc.close();
        }

        // 2. Read and verify
        {
            XLDocument doc;
            doc.open(filename);
            auto wks    = doc.workbook().worksheet("Sheet1");
            auto styles = doc.styles();

            auto cellE1  = wks.cell("E1");
            auto xfIdx   = cellE1.cellFormat();
            auto xfAlign = styles.cellFormats()[xfIdx];
            auto align   = xfAlign.alignment();

            // Verify via summary or specific checks if enum fails
            std::string s = align.summary();
            REQUIRE(s.find("horizontal=center") != std::string::npos);
            REQUIRE(s.find("vertical=top") != std::string::npos);
            REQUIRE(align.wrapText() == true);

            auto cellF1 = wks.cell("F1");
            auto border = styles.borders()[styles.cellFormats()[cellF1.cellFormat()].borderIndex()];
            REQUIRE(border.bottom().style() == XLLineStyleThick);

            doc.close();
        }
    }
}

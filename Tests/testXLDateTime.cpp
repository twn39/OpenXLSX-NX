#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <filesystem>
#include <fstream>

using namespace OpenXLSX;

namespace
{
    // Hack to access protected member without inheritance (since XLDocument is final)
    template<typename Tag, typename Tag::type M>
    struct Rob
    {
        friend typename Tag::type get_impl(Tag) { return M; }
    };

    struct XLDocument_extractXmlFromArchive
    {
        typedef std::string (XLDocument::*type)(std::string_view);
    };

    template struct Rob<XLDocument_extractXmlFromArchive, &XLDocument::extractXmlFromArchive>;

    // Prototype declaration for the friend function
    std::string (XLDocument::* get_impl(XLDocument_extractXmlFromArchive))(std::string_view);

    // Function to call the protected member
    std::string getRawXml(XLDocument& doc, const std::string& path)
    {
        static auto fn = get_impl(XLDocument_extractXmlFromArchive());
        return (doc.*fn)(path);
    }
}    // namespace

TEST_CASE("XLDateTime Tests", "[XLDateTime]")
{
    SECTION("Default construction")
    {
        XLDateTime dt;

        REQUIRE(dt.serial() == Catch::Approx(1.0));

        auto tm = dt.tm();
        REQUIRE(tm.tm_year == 0);
        REQUIRE(tm.tm_mon == 0);
        REQUIRE(tm.tm_mday == 1);
        REQUIRE(tm.tm_yday == 0);
        REQUIRE(tm.tm_wday == 0);
        REQUIRE(tm.tm_hour == 0);
        REQUIRE(tm.tm_min == 0);
        REQUIRE(tm.tm_sec == 0);
    }

    SECTION("Excel Epoch and 1900 Leap Year Bug")
    {
        // 1. Start of Epoch
        XLDateTime dt1(1.0);
        REQUIRE(dt1.toString() == "1900-01-01 00:00:00");
        REQUIRE(dt1.tm().tm_wday == 0);    // Sunday

        // 2. Day 59 (1900-02-28)
        XLDateTime dt59(59.0);
        REQUIRE(dt59.toString() == "1900-02-28 00:00:00");

        // 3. Day 60 (Excel's fictitious 1900-02-29)
        XLDateTime dt60(60.0);
        REQUIRE(dt60.toString() == "1900-02-29 00:00:00");

        // 4. Day 61 (1900-03-01)
        XLDateTime dt61(61.0);
        REQUIRE(dt61.toString() == "1900-03-01 00:00:00");
    }

    SECTION("Constructor (serial number)")
    {
        REQUIRE_THROWS(XLDateTime(-1.0));

        XLDateTime dt(6069.86742);

        REQUIRE(dt.serial() == Catch::Approx(6069.86742));

        auto tm = dt.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("Modern Date Conversion Accuracy")
    {
        // 2024-03-11 12:00:00 -> Serial 45362.5
        XLDateTime dt1(45362.5);
        REQUIRE(dt1.toString() == "2024-03-11 12:00:00");

        // 2024-03-12 12:00:00 -> Serial 45363.5
        XLDateTime dt2(45363.5);
        REQUIRE(dt2.toString() == "2024-03-12 12:00:00");
    }

    SECTION("Constructor (std::tm object)")
    {
        std::tm tmo;
        tmo.tm_isdst = -1;

        tmo.tm_year = -1;
        tmo.tm_mon  = 0;
        tmo.tm_mday = 1;
        tmo.tm_hour = 0;
        tmo.tm_min  = 0;
        tmo.tm_sec  = 0;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_year = 89;
        tmo.tm_mon  = 12;    // Invalid month (0-11)
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_mon = -1;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_mon  = 11;
        tmo.tm_mday = 0;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_mday = 32;
        REQUIRE_THROWS(XLDateTime(tmo));

        // Valid date: 1989-12-30 17:06:28
        tmo.tm_year = 89;
        tmo.tm_mon  = 11;
        tmo.tm_mday = 30;
        tmo.tm_hour = 17;
        tmo.tm_min  = 6;
        tmo.tm_sec  = 28;

        XLDateTime dt(tmo);
        REQUIRE(dt.serial() == Catch::Approx(32872.7128240741));
        REQUIRE(dt.toString() == "1989-12-30 17:06:28");
    }

    SECTION("Copy/Move and Assignment (Rule of Zero)")
    {
        XLDateTime dt(6069.86742);

        // Copy
        XLDateTime dt2 = dt;
        REQUIRE(dt2.serial() == Catch::Approx(dt.serial()));

        // Move
        XLDateTime dt3 = std::move(dt2);
        REQUIRE(dt3.serial() == Catch::Approx(dt.serial()));

        // Assignment
        XLDateTime dt4;
        dt4 = dt3;
        REQUIRE(dt4.serial() == Catch::Approx(dt.serial()));

        // std::tm assignment
        std::tm tmo;
        tmo.tm_year  = 124;    // 2024
        tmo.tm_mon   = 2;      // March
        tmo.tm_mday  = 12;
        tmo.tm_hour  = 12;
        tmo.tm_min   = 0;
        tmo.tm_sec   = 0;
        tmo.tm_isdst = -1;

        dt4 = tmo;
        REQUIRE(dt4.serial() == Catch::Approx(45363.5));
    }

    SECTION("Implicit conversion")
    {
        XLDateTime dt(45363.5);    // 2024-03-12 12:00

        double  serial = dt;
        std::tm result = dt;

        REQUIRE(serial == Catch::Approx(45363.5));
        REQUIRE(result.tm_year == 124);
        REQUIRE(result.tm_mon == 2);
        REQUIRE(result.tm_mday == 12);
        REQUIRE(result.tm_hour == 12);
    }

    SECTION("String Parsing and Formatting")
    {
        // Parsing
        XLDateTime dt = XLDateTime::fromString("2023-10-27 14:30:00");
        auto       tm = dt.tm();
        REQUIRE(tm.tm_year == 123);    // 2023 - 1900
        REQUIRE(tm.tm_mon == 9);       // October (0-indexed)
        REQUIRE(tm.tm_mday == 27);

        // Formatting
        REQUIRE(dt.toString() == "2023-10-27 14:30:00");
        REQUIRE(dt.toString("%Y/%m/%d") == "2023/10/27");

        // Error handling
        REQUIRE_THROWS(XLDateTime::fromString("invalid-date"));
    }

    SECTION("Chrono Support")
    {
        auto       now = std::chrono::system_clock::now();
        XLDateTime dt(now);

        // Convert back to chrono
        auto back = dt.chrono();

        // Check difference (should be less than 1 second due to Excel precision)
        auto diff = std::chrono::duration_cast<std::chrono::seconds>(now - back).count();
        REQUIRE(std::abs(diff) <= 1);

        XLDateTime dtNow = XLDateTime::now();
        REQUIRE(dtNow.serial() > 45000);
    }

    SECTION("Cell Integration and XML Storage")
    {
        std::string filename = "DateTimeCellTest.xlsx";
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // Store exactly 2024-03-12 12:00:00
            XLDateTime dt(45363.5);
            wks.cell("A1").value() = dt;

            doc.save();
            doc.close();
        }

        {
            XLDocument doc2;
            doc2.open(filename);
            auto wks2 = doc2.workbook().worksheet("Sheet1");

            // Excel stores dates as doubles
            double serial = wks2.cell("A1").value().get<double>();
            REQUIRE(serial == Catch::Approx(45363.5));

            XLDateTime dt2(serial);
            REQUIRE(dt2.toString() == "2024-03-12 12:00:00");

            // Verify XML structure
            std::string sheetXml = getRawXml(doc2, "xl/worksheets/sheet1.xml");
            REQUIRE(sheetXml.find("<v>45363.5</v>") != std::string::npos);

            doc2.close();
        }
        std::filesystem::remove(filename);
    }
}

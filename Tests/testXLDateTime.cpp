//
// Created by Kenneth Balslev on 29/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
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
        typedef std::string (XLDocument::*type)(const std::string&);
    };

    template struct Rob<XLDocument_extractXmlFromArchive, &XLDocument::extractXmlFromArchive>;

    // Prototype declaration for the friend function
    std::string (XLDocument::*get_impl(XLDocument_extractXmlFromArchive))(const std::string&);

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

    SECTION("Constructor (serial number)")
    {
        REQUIRE_THROWS(XLDateTime(0.0));

        XLDateTime dt(6069.86742);

        REQUIRE(dt.serial() == Catch::Approx(6069.86742));

        auto tm = dt.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        //        REQUIRE(tm.tm_yday == 0);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("Constructor (serial number, seconds rounding)")
    {
        REQUIRE_THROWS(XLDateTime(0.0));

        const double serial = 6069.000008;
        XLDateTime   dt(serial);

        REQUIRE(dt.serial() == Catch::Approx(serial));

        auto tm = dt.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 0);
        REQUIRE(tm.tm_min == 0);
        REQUIRE(tm.tm_sec == 1);
    }

    SECTION("Constructor (std::tm object)")
    {
        std::tm tmo;

        tmo.tm_year = -1;
        tmo.tm_mon  = 0;
        tmo.tm_mday = 0;
        tmo.tm_yday = 0;
        tmo.tm_wday = 0;
        tmo.tm_hour = 0;
        tmo.tm_min  = 0;
        tmo.tm_sec  = 0;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_year = 89;
        tmo.tm_mon  = 13;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_mon = -1;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_mon  = 11;
        tmo.tm_mday = 0;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_mday = 32;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_mday = 30;
        tmo.tm_hour = 17;
        tmo.tm_min  = 6;
        tmo.tm_sec  = 28;
        tmo.tm_wday = 6;

        XLDateTime dt(tmo);

        REQUIRE(dt.serial() == Catch::Approx(32872.7128));
    }

    SECTION("Copy Constructor")
    {
        XLDateTime dt(6069.86742);
        XLDateTime dt2 = dt;

        auto tm = dt2.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("Move Constructor")
    {
        XLDateTime dt(6069.86742);
        XLDateTime dt2 = std::move(dt);

        auto tm = dt2.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("Copy Assignment")
    {
        XLDateTime dt(6069.86742);
        XLDateTime dt2;
        dt2 = dt;

        auto tm = dt2.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("Move Assignment")
    {
        XLDateTime dt(6069.86742);
        XLDateTime dt2;
        dt2 = std::move(dt);

        auto tm = dt2.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("Serial Assignment")
    {
        XLDateTime dt(1.0);
        dt = 6069.86742;

        auto tm = dt.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("std::tm Assignment")
    {
        std::tm tmo;
        tmo.tm_year = 16;
        tmo.tm_mon  = 7;
        tmo.tm_mday = 12;
        tmo.tm_wday = 6;
        tmo.tm_hour = 20;
        tmo.tm_min  = 49;
        tmo.tm_sec  = 5;

        XLDateTime dt(1.0);
        dt = tmo;

        REQUIRE(dt.serial() == Catch::Approx(6069.86742));
    }

    SECTION("Implicit conversion")
    {
        std::tm tmo;
        tmo.tm_year = 16;
        tmo.tm_mon  = 7;
        tmo.tm_mday = 12;
        tmo.tm_wday = 6;
        tmo.tm_hour = 20;
        tmo.tm_min  = 49;
        tmo.tm_sec  = 5;

        XLDateTime dt(1.0);
        dt = tmo;

        double  serial = dt;
        std::tm result = dt;

        REQUIRE(serial == Catch::Approx(6069.86742));
        REQUIRE(result.tm_year == 16);
        REQUIRE(result.tm_mon == 7);
        REQUIRE(result.tm_mday == 12);
        REQUIRE(result.tm_wday == 6);
        REQUIRE(result.tm_hour == 20);
        REQUIRE(result.tm_min == 49);
        REQUIRE(result.tm_sec == 5);
    }

    SECTION("String Parsing and Formatting")
    {
        // Parsing
        XLDateTime dt = XLDateTime::fromString("2023-10-27 14:30:00");
        auto       tm = dt.tm();
        REQUIRE(tm.tm_year == 123); // 2023 - 1900
        REQUIRE(tm.tm_mon == 9);    // October (0-indexed)
        REQUIRE(tm.tm_mday == 27);
        REQUIRE(tm.tm_hour == 14);
        REQUIRE(tm.tm_min == 30);
        REQUIRE(tm.tm_sec == 0);

        // Formatting
        REQUIRE(dt.toString() == "2023-10-27 14:30:00");
        REQUIRE(dt.toString("%Y/%m/%d") == "2023/10/27");

        // Error handling
        REQUIRE_THROWS(XLDateTime::fromString("invalid-date"));
    }

    SECTION("Chrono Support")
    {
        auto now = std::chrono::system_clock::now();
        XLDateTime dt(now);
        
        // Convert back to chrono
        auto back = dt.chrono();
        
        // Check difference (should be less than 1 second due to Excel precision)
        auto diff = std::chrono::duration_cast<std::chrono::seconds>(now - back).count();
        REQUIRE(std::abs(diff) <= 1);

        XLDateTime dtNow = XLDateTime::now();
        REQUIRE(dtNow.serial() > 45000); // Current dates are > 45000
    }

    SECTION("Cell Integration")
    {
        XLDocument doc;
        doc.create("./DateTimeCellTest.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        XLDateTime dt = XLDateTime::fromString("2023-10-27 14:30:00");
        wks.cell("A1").value() = dt;

        doc.save();
        doc.close();

        XLDocument doc2;
        doc2.open("./DateTimeCellTest.xlsx");
        auto wks2 = doc2.workbook().worksheet("Sheet1");
        
        // Excel stores dates as doubles
        double serial = wks2.cell("A1").value().get<double>();
        XLDateTime dt2(serial);
        
        REQUIRE(dt2.toString() == "2023-10-27 14:30:00");

        // Verify XML structure
        std::string sheetXml = getRawXml(doc2, "xl/worksheets/sheet1.xml");
        // Excel date for 2023-10-27 14:30:00 is approximately 45226.604167
        REQUIRE(sheetXml.find("<v>45226.604166666664</v>") != std::string::npos);

        doc2.close();
    }
}

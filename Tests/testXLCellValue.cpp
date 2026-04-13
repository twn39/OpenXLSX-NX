#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"
#include <fstream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("control_chars_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testXLCellValue_xlsx") + ".xlsx";
    return name;
}
} // namespace


/**
 * @brief The purpose of this test case is to test the creation of XLDocument objects. Each section section
 * tests document creation using a different method. In addition, saving, closing and copying is tested.
 */
TEST_CASE("XLCellValueTests", "[XLCellValue]")
{
    SECTION("Default Constructor")
    {
        XLCellValue value;

        REQUIRE(value.type() == XLValueType::Empty);
        REQUIRE(std::string(value.typeAsString()) == "empty");
        REQUIRE(value.get<std::string>().empty());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Float Constructor")
    {
        XLCellValue value(3.14159);

        REQUIRE(value.type() == XLValueType::Float);
        REQUIRE(std::string(value.typeAsString()) == "float");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 3.14159);
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Integer Constructor")
    {
        XLCellValue value(42);

        REQUIRE(value.type() == XLValueType::Integer);
        REQUIRE(std::string(value.typeAsString()) == "integer");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE(value.get<int>() == 42);
        REQUIRE(value.get<double>() == 42.0);
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Boolean Constructor")
    {
        XLCellValue value(true);

        REQUIRE(value.type() == XLValueType::Boolean);
        REQUIRE(std::string(value.typeAsString()) == "boolean");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 1.0);
        REQUIRE(value.get<bool>() == true);
    }

    SECTION("String Constructor")
    {
        XLCellValue value("Hello OpenXLSX!");

        REQUIRE(value.type() == XLValueType::String);
        REQUIRE(std::string(value.typeAsString()) == "string");
        REQUIRE(value.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Copy Constructor")
    {
        XLCellValue value("Hello OpenXLSX!");
        auto        copy = value;

        REQUIRE(copy.type() == XLValueType::String);
        REQUIRE(std::string(value.typeAsString()) == "string");
        REQUIRE(copy.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(copy.get<int>());
        REQUIRE_THROWS(copy.get<double>());
        REQUIRE_THROWS(copy.get<bool>());
    }

    SECTION("Move Constructor")
    {
        XLCellValue value("Hello OpenXLSX!");
        auto        copy = std::move(value);

        REQUIRE(copy.type() == XLValueType::String);
        REQUIRE(std::string(value.typeAsString()) == "string");
        REQUIRE(copy.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(copy.get<int>());
        REQUIRE_THROWS(copy.get<double>());
        REQUIRE_THROWS(copy.get<bool>());
    }

    SECTION("Copy Assignment")
    {
        XLCellValue value("Hello OpenXLSX!");
        XLCellValue copy(1);
        copy = value;

        REQUIRE(copy.type() == XLValueType::String);
        REQUIRE(std::string(value.typeAsString()) == "string");
        REQUIRE(copy.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(copy.get<int>());
        REQUIRE_THROWS(copy.get<double>());
        REQUIRE_THROWS(copy.get<bool>());
    }

    SECTION("Move Assignment")
    {
        XLCellValue value("Hello OpenXLSX!");
        XLCellValue copy(1);
        copy = std::move(value);

        REQUIRE(copy.type() == XLValueType::String);
        REQUIRE(std::string(value.typeAsString()) == "string");
        REQUIRE(copy.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(copy.get<int>());
        REQUIRE_THROWS(copy.get<double>());
        REQUIRE_THROWS(copy.get<bool>());
    }

    SECTION("Float Assignment")
    {
        XLCellValue value;
        value = 3.14159;

        REQUIRE(value.type() == XLValueType::Float);
        REQUIRE(std::string(value.typeAsString()) == "float");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 3.14159);
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Integer Assignment")
    {
        XLCellValue value;
        value = 42;

        REQUIRE(value.type() == XLValueType::Integer);
        REQUIRE(std::string(value.typeAsString()) == "integer");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE(value.get<int>() == 42);
        REQUIRE(value.get<double>() == 42.0);
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Boolean Assignment")
    {
        XLCellValue value;
        value = true;

        REQUIRE(value.type() == XLValueType::Boolean);
        REQUIRE(std::string(value.typeAsString()) == "boolean");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 1.0);
        REQUIRE(value.get<bool>() == true);
    }

    SECTION("String Assignment")
    {
        XLCellValue value;
        value = "Hello OpenXLSX!";

        REQUIRE(value.type() == XLValueType::String);
        REQUIRE(std::string(value.typeAsString()) == "string");
        REQUIRE(value.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("XLCellValueProxy Assignment")
    {
        XLCellValue value;
        XLDocument  doc;
        doc.create(__global_unique_file_1(), XLForceOverwrite);
        XLWorksheet wks = doc.workbook().sheet(1);

        wks.cell("A1").value() = "Hello OpenXLSX!";
        value                  = wks.cell("A1").value();
        REQUIRE(value.type() == XLValueType::String);
        REQUIRE(std::string(value.typeAsString()) == "string");
        REQUIRE(value.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());

        wks.cell("A1").value() = 3.14159;
        value                  = wks.cell("A1").value();
        REQUIRE(value.type() == XLValueType::Float);
        REQUIRE(std::string(value.typeAsString()) == "float");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 3.14159);
        REQUIRE_THROWS(value.get<bool>());

        wks.cell("A1").value() = 42;
        value                  = wks.cell("A1").value();
        REQUIRE(value.type() == XLValueType::Integer);
        REQUIRE(std::string(value.typeAsString()) == "integer");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE(value.get<int>() == 42);
        REQUIRE(value.get<double>() == 42.0);
        REQUIRE_THROWS(value.get<bool>());

        wks.cell("A1").value() = true;
        value                  = wks.cell("A1").value();
        REQUIRE(value.type() == XLValueType::Boolean);
        REQUIRE(std::string(value.typeAsString()) == "boolean");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 1.0);
        REQUIRE(value.get<bool>() == true);
    }

    SECTION("Set Float ")
    {
        XLCellValue value;
        value.set(3.14159);

        REQUIRE(value.type() == XLValueType::Float);
        REQUIRE(std::string(value.typeAsString()) == "float");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 3.14159);
        REQUIRE_THROWS(value.get<bool>());

        // Microsoft Excel only supports up to 15 significant digits (IEEE 754 precision limit rule).
        // 0.1 + 0.2 creates a 0.30000000000000004 float jitter.
        double jitteryFloat = 0.1 + 0.2;
        value.set(jitteryFloat);
        // It must round to exactly 15 digits internally and be readable as clean 0.3
        REQUIRE(value.get<double>() == Catch::Approx(0.3).epsilon(1e-15));
    }

    SECTION("Set Nan Float")
    {
        XLCellValue value;
        value.set(std::numeric_limits<double>::quiet_NaN());

        REQUIRE(value.type() == XLValueType::Error);
        REQUIRE(std::string(value.typeAsString()) == "error");
        REQUIRE(value.get<std::string>() == "#NUM!");
        REQUIRE_THROWS(value.get<int>());
        //        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Set Inf Float")
    {
        XLCellValue value;
        value.set(std::numeric_limits<double>::infinity());

        REQUIRE(value.type() == XLValueType::Error);
        REQUIRE(std::string(value.typeAsString()) == "error");
        REQUIRE(value.get<std::string>() == "#NUM!");
        REQUIRE_THROWS(value.get<int>());
        //        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Set Integer")
    {
        XLCellValue value;
        value.set(42);

        REQUIRE(value.type() == XLValueType::Integer);
        REQUIRE(std::string(value.typeAsString()) == "integer");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE(value.get<int>() == 42);
        REQUIRE(value.get<double>() == 42.0);
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Set Boolean")
    {
        XLCellValue value;
        value.set(true);

        REQUIRE(value.type() == XLValueType::Boolean);
        REQUIRE(std::string(value.typeAsString()) == "boolean");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 1.0);
        REQUIRE(value.get<bool>() == true);
    }

    SECTION("Set String")
    {
        XLCellValue value;
        value.set("Hello OpenXLSX!");

        REQUIRE(value.type() == XLValueType::String);
        REQUIRE(std::string(value.typeAsString()) == "string");
        REQUIRE(value.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Control Characters Corruption Resistance")
    {
        XLDocument doc;
        doc.create(__global_unique_file_0(), XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // String with valid and invalid XML 1.0 characters
        // \x09 (\t), \x0A (\n), \x0D (\r) are VALID.
        // \x0B (vertical tab), \x01, \x1F are INVALID.
        std::string maliciousString = "Valid \x09\x0A Invalid \x0B\x01\x1F Text";
        std::string expectedString  = "Valid \x09\x0A Invalid  Text"; // \x0B, \x01, \x1F stripped out

        wks.cell("A1").value() = maliciousString;

        doc.save();
        doc.close();

        // Should successfully reopen without pugixml crash
        doc.open(__global_unique_file_0());
        wks = doc.workbook().worksheet("Sheet1");

        // The read-back value should be exactly the sanitized string
        std::string readBack = wks.cell("A1").value().get<std::string>();
        REQUIRE(readBack == expectedString);

        doc.close();
        std::remove(__global_unique_file_0().c_str());
    }

    SECTION("Clear")
    {
        XLCellValue value;
        value.set("Hello OpenXLSX!");
        value.clear();

        REQUIRE(value.type() == XLValueType::Empty);
        REQUIRE(std::string(value.typeAsString()) == "empty");
        REQUIRE(value.get<std::string>().empty());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Set Error")
    {
        XLCellValue value;
        value.set("Hello OpenXLSX!");
        value.setError("#N/A");

        REQUIRE(value.type() == XLValueType::Error);
        REQUIRE(std::string(value.typeAsString()) == "error");
        REQUIRE(value.get<std::string>() == "#N/A");
        REQUIRE_THROWS(value.get<int>());
        //        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Implicit conversion to supported types")
    {
        XLCellValue value;

        value        = "Hello OpenXLSX!";
        auto result1 = value.get<std::string>();
        REQUIRE(result1 == "Hello OpenXLSX!");
        REQUIRE_THROWS(static_cast<int>(value));
        REQUIRE_THROWS(static_cast<double>(value));
        REQUIRE_THROWS(static_cast<bool>(value));

        value        = 42;
        auto result2 = static_cast<int>(value);
        REQUIRE(result2 == 42);
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE(static_cast<double>(value) == 42.0);
        REQUIRE_THROWS(static_cast<bool>(value));

        value        = 3.14159;
        auto result3 = static_cast<double>(value);
        REQUIRE(result3 == 3.14159);
        REQUIRE_THROWS(static_cast<int>(value));
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(static_cast<bool>(value));

        value        = true;
        auto result4 = static_cast<bool>(value);
        REQUIRE(result4 == true);
        REQUIRE_THROWS(static_cast<int>(value));
        REQUIRE(static_cast<double>(value) == 1.0);
        REQUIRE_THROWS(value.get<std::string>());
    }
}

#include <OpenXLSX.hpp>
#include <iostream>
#include <catch2/catch_test_macros.hpp>

using namespace OpenXLSX;

TEST_CASE("XLDataValidation Tests", "[XLDataValidation]")
{
    SECTION("Create and set data validations") {
        XLDocument doc;
        doc.create("./testDataValidation.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        auto& validations = wks.dataValidations();
        
        // 1. Whole number range
        auto dv1 = validations.append();
        dv1.setSqref("A1:A10");
        dv1.setWholeNumberRange(1, 100);
        dv1.setPrompt("Enter Number", "Please enter a number between 1 and 100.");
        dv1.setError("Invalid Number", "The number you entered is not valid!", XLDataValidationErrorStyle::Stop);

        // 2. List validation
        auto dv2 = validations.append();
        dv2.setSqref("B1:B10");
        dv2.setList({"Option1", "Option2", "Option3"});
        dv2.setPrompt("Select Option", "Please select an option from the list.");

        REQUIRE(validations.count() == 2);

        doc.save();
        doc.close();

        // Re-open and verify
        XLDocument doc2;
        doc2.open("./testDataValidation.xlsx");
        auto wks2 = doc2.workbook().worksheet("Sheet1");
        auto& validations2 = wks2.dataValidations();
        
        REQUIRE(validations2.count() == 2);
        
        doc2.close();
    }

    SECTION("Clear data validations") {
        XLDocument doc;
        doc.create("./testDataValidationClear.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        
        auto& validations = wks.dataValidations();
        validations.append().setSqref("A1");
        validations.append().setSqref("B1");
        
        REQUIRE(validations.count() == 2);
        
        validations.clear();
        REQUIRE(validations.count() == 0);
        
        doc.save();
        doc.close();
    }

    SECTION("New Data Validation Features (P1, P2, P3)") {
        XLDocument doc;
        doc.create("./testNewDataValidationFeatures.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        auto& validations = wks.dataValidations();

        // 1. Test Decimal Range with P1 properties
        auto dv1 = validations.append();
        dv1.setSqref("C1:C5");
        dv1.setDecimalRange(0.5, 99.9);
        dv1.setShowDropDown(false);
        dv1.setShowInputMessage(true);
        dv1.setShowErrorMessage(true);
        dv1.setIMEMode(XLIMEMode::Disabled);

        // 2. Test Date and Time Range
        auto dv2 = validations.append();
        dv2.setSqref("D1:D5");
        dv2.setDateRange("2023-01-01", "2023-12-31");

        auto dv3 = validations.append();
        dv3.setSqref("E1:E5");
        dv3.setTimeRange("08:00", "18:00");

        // 3. Test Text Length
        auto dv4 = validations.append();
        dv4.setSqref("F1:F5");
        dv4.setTextLengthRange(1, 10);

        REQUIRE(validations.count() == 4);

        // Test Getters before saving
        REQUIRE(dv1.sqref() == "C1:C5");
        REQUIRE(dv1.type() == XLDataValidationType::Decimal);
        REQUIRE(dv1.operator_() == XLDataValidationOperator::Between);
        REQUIRE(dv1.formula1() == "0.5");
        REQUIRE(dv1.formula2() == "99.9");
        REQUIRE(dv1.showDropDown() == false);
        REQUIRE(dv1.imeMode() == XLIMEMode::Disabled);

        REQUIRE(dv2.type() == XLDataValidationType::Date);
        REQUIRE(dv2.formula1() == "2023-01-01");

        // Test at(sqref)
        auto dvFound = validations.at("C1:C5");
        REQUIRE_FALSE(dvFound.empty());
        REQUIRE(dvFound.type() == XLDataValidationType::Decimal);

        doc.save();
        doc.close();

        // Re-open and verify persistence
        XLDocument doc2;
        doc2.open("./testNewDataValidationFeatures.xlsx");
        auto wks2 = doc2.workbook().worksheet("Sheet1");
        auto& v2 = wks2.dataValidations();
        
        auto dvVerify = v2.at("C1:C5");
        REQUIRE_FALSE(dvVerify.empty());
        REQUIRE(dvVerify.type() == XLDataValidationType::Decimal);
        REQUIRE(dvVerify.formula1() == "0.5");
        REQUIRE(dvVerify.showDropDown() == false);
        REQUIRE(dvVerify.imeMode() == XLIMEMode::Disabled);

        auto dvText = v2.at("F1:F5");
        REQUIRE(dvText.type() == XLDataValidationType::TextLength);
        REQUIRE(dvText.formula1() == "1");
        REQUIRE(dvText.formula2() == "10");

        doc2.close();
    }

    SECTION("Data Validation Range Collapsing") {
        XLDocument doc;
        doc.create("./testDataValidationCollapsing.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        auto& validations = wks.dataValidations();

        auto dv = validations.append();
        dv.addCell("A1");
        REQUIRE(dv.sqref() == "A1");

        dv.addCell("A2");
        REQUIRE(dv.sqref() == "A1:A2");

        dv.addCell("B1");
        dv.addCell("B2");
        REQUIRE(dv.sqref() == "A1:B2");

        dv.addCell("A4");
        dv.addCell("A5");
        REQUIRE(dv.sqref() == "A1:B2 A4:A5");

        dv.addRange("B4:B5");
        // A1:B2 and A4:B5
        dv.addRange("A3:B3");
        // Should bridge them together
        REQUIRE(dv.sqref() == "A1:B5");

        // Duplicate/contained range
        dv.addRange("B2:B4");
        REQUIRE(dv.sqref() == "A1:B5");

        doc.close();
    }

    SECTION("Data Validation List Limits and Escaping") {
        XLDocument doc;
        doc.create("./testDataValidationList.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        auto& validations = wks.dataValidations();

        auto dv = validations.append();
        dv.setSqref("A1");
        
        // Test escaping
        dv.setList({"Normal", "Has \"Quotes\"", "End\""});
        REQUIRE(dv.formula1() == "\"Normal,Has \"\"Quotes\"\",End\"\"\"");

        // Test length limit (255 characters)
        std::vector<std::string> longList;
        for (int i = 0; i < 30; ++i) {
            longList.push_back("1234567890"); 
        }
        // Total length will be over 300
        REQUIRE_THROWS_AS(dv.setList(longList), XLException);

        doc.close();
    }
}

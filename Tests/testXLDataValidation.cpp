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
}

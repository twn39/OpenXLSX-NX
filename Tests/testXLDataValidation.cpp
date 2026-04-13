#include "TestHelpers.hpp"
#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testDataValidationClear_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testNewDataValidationFeatures_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testDataValidationIterator_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_3() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testDataValidationList_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_4() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testDataValidationCollapsing_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_5() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testDataValidationGlobal_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_6() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testDataValidation_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_7() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testDataValidationConfig_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("XLDataValidationTests", "[XLDataValidation]")
{
    SECTION("Create and set data validations")
    {
        XLDocument doc;
        doc.create(__global_unique_file_6(), XLForceOverwrite);
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
        doc2.open(__global_unique_file_6());
        auto  wks2         = doc2.workbook().worksheet("Sheet1");
        auto& validations2 = wks2.dataValidations();

        REQUIRE(validations2.count() == 2);

        doc2.close();
    }

    SECTION("Clear data validations")
    {
        XLDocument doc;
        doc.create(__global_unique_file_0(), XLForceOverwrite);
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

    SECTION("New Data Validation Features (P1, P2, P3)")
    {
        XLDocument doc;
        doc.create(__global_unique_file_1(), XLForceOverwrite);
        auto  wks         = doc.workbook().worksheet("Sheet1");
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
        doc2.open(__global_unique_file_1());
        auto  wks2 = doc2.workbook().worksheet("Sheet1");
        auto& v2   = wks2.dataValidations();

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

    SECTION("Data Validation Range Collapsing")
    {
        XLDocument doc;
        doc.create(__global_unique_file_4(), XLForceOverwrite);
        auto  wks         = doc.workbook().worksheet("Sheet1");
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

        // --- Test Range Subtraction (removeCell / removeRange) ---
        // Current: A1:B5
        dv.removeCell("A1");
        // Output might be "A2:B5 B1" depending on collapse order
        REQUIRE(dv.sqref() == "A2:B5 B1");

        // Reset to a clean block
        dv.setSqref("A1:C5");

        // Remove a block from the middle
        dv.removeRange("B2:B4");
        // Remaining should be the top row, bottom row, and left/right columns
        // The output could vary based on collapse. Let's test a point that we know B2 is absent.
        auto sqrefStr = dv.sqref();
        REQUIRE(sqrefStr.find("B2") == std::string::npos);
        REQUIRE(sqrefStr.find("B3") == std::string::npos);
        REQUIRE(sqrefStr.find("B4") == std::string::npos);

        doc.close();
    }

    SECTION("Data Validation Iterators and Precise Removal")
    {
        XLDocument doc;
        doc.create(__global_unique_file_2(), XLForceOverwrite);
        auto  wks         = doc.workbook().worksheet("Sheet1");
        auto& validations = wks.dataValidations();

        auto dv1 = validations.append();
        dv1.setSqref("A1");
        dv1.setType(XLDataValidationType::Whole);
        auto dv2 = validations.append();
        dv2.setSqref("B1");
        dv2.setType(XLDataValidationType::Decimal);
        auto dv3 = validations.append();
        dv3.setSqref("C1");
        dv3.setType(XLDataValidationType::List);

        // 1. Test Iterator
        int count = 0;
        for (auto dv : validations) {
            count++;
            REQUIRE_FALSE(dv.empty());
        }
        REQUIRE(count == 3);

        // 2. Test Precise Removal by sqref
        validations.remove("B1");
        REQUIRE(validations.count() == 2);

        // Verify B1 is gone, C1 and A1 remain
        bool foundDecimal = false;
        for (auto dv : validations) {
            if (dv.type() == XLDataValidationType::Decimal) foundDecimal = true;
        }
        REQUIRE_FALSE(foundDecimal);

        // 3. Test Precise Removal by index
        validations.remove(0);    // Removes A1
        REQUIRE(validations.count() == 1);
        REQUIRE(validations.begin()->type() == XLDataValidationType::List);

        // 4. Test Last Element Removal (Should remove the <dataValidations> root node)
        validations.remove(0);
        REQUIRE(validations.count() == 0);
        REQUIRE(validations.empty() == true);

        doc.close();
    }

    SECTION("Data Validation List Limits and Escaping")
    {
        XLDocument doc;
        doc.create(__global_unique_file_3(), XLForceOverwrite);
        auto  wks         = doc.workbook().worksheet("Sheet1");
        auto& validations = wks.dataValidations();

        auto dv = validations.append();
        dv.setSqref("A1");

        // Test escaping
        dv.setList({"Normal", "Has \"Quotes\"", "End\""});
        REQUIRE(dv.formula1() == "\"Normal,Has \"\"Quotes\"\",End\"\"\"");

        // Test length limit (255 characters)
        std::vector<std::string> longList;
        for (int i = 0; i < 30; ++i) { longList.push_back("1234567890"); }
        // Total length will be over 300
        REQUIRE_THROWS_AS(dv.setList(longList), XLException);

        // Test Dropdown Reference with special characters and absolute references (best practice)
        auto dvRef = validations.append();
        dvRef.setSqref("B1");
        dvRef.setReferenceDropList("Data Sheet", "$A$1:$A$10");
        // Should be correctly escaped and prefixed
        REQUIRE(dvRef.formula1() == "='Data Sheet'!$A$1:$A$10");

        // Test Dropdown Reference with single quotes in sheet name
        auto dvRef2 = validations.append();
        dvRef2.setSqref("C1");
        dvRef2.setReferenceDropList("Jane's Data", "$B$1:$B$10");
        REQUIRE(dvRef2.formula1() == "='Jane''s Data'!$B$1:$B$10");

        // Test Dropdown Reference same sheet (no sheet name)
        auto dvRef3 = validations.append();
        dvRef3.setSqref("D1");
        dvRef3.setReferenceDropList("", "$Z$1:$Z$100");
        REQUIRE(dvRef3.formula1() == "=$Z$1:$Z$100");

        doc.close();
    }

    SECTION("Data Validation Global Properties")
    {
        XLDocument doc;
        doc.create(__global_unique_file_5(), XLForceOverwrite);
        auto  wks         = doc.workbook().worksheet("Sheet1");
        auto& validations = wks.dataValidations();

        // Must append at least one to create the node, otherwise properties won't attach
        validations.append().setSqref("A1");

        validations.setDisablePrompts(true);
        validations.setXWindow(100);
        validations.setYWindow(200);

        REQUIRE(validations.disablePrompts() == true);
        REQUIRE(validations.xWindow() == 100);
        REQUIRE(validations.yWindow() == 200);

        doc.save();
        doc.close();

        // Re-open and verify
        XLDocument doc2;
        doc2.open(__global_unique_file_5());
        auto  wks2 = doc2.workbook().worksheet("Sheet1");
        auto& v2   = wks2.dataValidations();

        REQUIRE(v2.disablePrompts() == true);
        REQUIRE(v2.xWindow() == 100);
        REQUIRE(v2.yWindow() == 200);

        doc2.close();
    }

    SECTION("Data Validation Config and Ergonomics")
    {
        XLDocument doc;
        doc.create(__global_unique_file_7(), XLForceOverwrite);
        auto wks1 = doc.workbook().worksheet("Sheet1");
        doc.workbook().addWorksheet("Sheet2");
        auto wks2 = doc.workbook().worksheet("Sheet2");

        // Create a generic config
        XLDataValidationConfig rule;
        rule.type             = XLDataValidationType::List;
        rule.operator_        = XLDataValidationOperator::Between;
        rule.allowBlank       = true;
        rule.showDropDown     = true;
        rule.showInputMessage = true;
        rule.promptTitle      = "Select Option";
        rule.prompt           = "Please select from list";
        rule.formula1         = "\"A,B,C\"";

        // Apply it to Sheet1
        wks1.dataValidations().addValidation(rule, "A1:A10");

        // Apply it to Sheet2
        wks2.dataValidations().addValidation(rule, "B1:B10");

        // Verify Sheet1
        auto dv1 = wks1.dataValidations().begin();
        REQUIRE(dv1->type() == XLDataValidationType::List);
        REQUIRE(dv1->promptTitle() == "Select Option");
        REQUIRE(dv1->formula1() == "\"A,B,C\"");
        REQUIRE(dv1->sqref() == "A1:A10");

        // Verify Sheet2
        auto dv2 = wks2.dataValidations().begin();
        REQUIRE(dv2->type() == XLDataValidationType::List);
        REQUIRE(dv2->promptTitle() == "Select Option");
        REQUIRE(dv2->formula1() == "\"A,B,C\"");
        REQUIRE(dv2->sqref() == "B1:B10");

        // Test extracting config from existing validation
        XLDataValidationConfig extractedRule = dv1->config();
        REQUIRE(extractedRule.prompt == "Please select from list");
        extractedRule.prompt = "Modified prompt";

        // Apply modified rule
        auto dv3 = wks2.dataValidations().addValidation(extractedRule, "C1:C10");
        REQUIRE(dv3.prompt() == "Modified prompt");
        REQUIRE(dv3.sqref() == "C1:C10");

        doc.close();
    }
}

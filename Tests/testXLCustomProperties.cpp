#include "TestHelpers.hpp"
#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testCustomProps_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__testCustomPropsDelete_xlsx") + ".xlsx";
    return name;
}
} // namespace


TEST_CASE("XLCustomPropertiesTests", "[XLCustomProperties]")
{
    SECTION("Create and set custom properties")
    {
        XLDocument doc;
        doc.create(__global_unique_file_0(), XLForceOverwrite);

        auto& customProps = doc.customProperties();
        customProps.setProperty("MyStringProp", "Hello OpenXLSX");
        customProps.setProperty("MyIntProp", 42);
        customProps.setProperty("MyDoubleProp", 3.14159);
        customProps.setProperty("MyBoolProp", true);

        doc.save();
        doc.close();

        XLDocument doc2;
        doc2.open(__global_unique_file_0());

        auto& customProps2 = doc2.customProperties();
        REQUIRE(customProps2.property("MyStringProp") == "Hello OpenXLSX");
        REQUIRE(customProps2.property("MyIntProp") == "42");
        REQUIRE(customProps2.property("MyDoubleProp") == "3.14159");    // fmt::format doesn't add extra zeros by default
        REQUIRE(customProps2.property("MyBoolProp") == "true");

        doc2.close();
    }

    SECTION("Delete custom properties")
    {
        XLDocument doc;
        doc.create(__global_unique_file_1(), XLForceOverwrite);

        auto& customProps = doc.customProperties();
        customProps.setProperty("ToByDeleted", "Gone soon");
        customProps.setProperty("ToKeep", "I stay");

        customProps.deleteProperty("ToByDeleted");

        REQUIRE(customProps.property("ToByDeleted") == "");
        REQUIRE(customProps.property("ToKeep") == "I stay");

        doc.save();
        doc.close();
    }
}

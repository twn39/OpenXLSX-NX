#include <catch2/catch_all.hpp>
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Modern XLProperties Tests", "[XLProperties]")
{
    SECTION("Core Properties Getters/Setters")
    {
        XLDocument doc;
        doc.create("PropertiesTest.xlsx", XLForceOverwrite);
        auto& props = doc.coreProperties();

        props.setTitle("Test Title");
        props.setSubject("Test Subject");
        props.setCreator("Test Creator");
        props.setKeywords("Test Keywords");
        props.setDescription("Test Description");
        props.setLastModifiedBy("Test Author");
        props.setRevision("2");
        props.setCategory("Test Category");
        props.setContentStatus("Draft");

        XLDateTime createdDate = XLDateTime::fromString("2023-05-20 10:30:00");
        props.setCreated(createdDate);

        XLDateTime modifiedDate = XLDateTime::fromString("2024-03-15 14:45:30");
        props.setModified(modifiedDate);

        REQUIRE(props.title() == "Test Title");
        REQUIRE(props.subject() == "Test Subject");
        REQUIRE(props.creator() == "Test Creator");
        REQUIRE(props.keywords() == "Test Keywords");
        REQUIRE(props.description() == "Test Description");
        REQUIRE(props.lastModifiedBy() == "Test Author");
        REQUIRE(props.revision() == "2");
        REQUIRE(props.category() == "Test Category");
        REQUIRE(props.contentStatus() == "Draft");
        
        REQUIRE(props.created().toString("%Y-%m-%dT%H:%M:%SZ") == "2023-05-20T10:30:00Z");
        REQUIRE(props.modified().toString("%Y-%m-%dT%H:%M:%SZ") == "2024-03-15T14:45:30Z");

        doc.save();
        doc.close();
    }

    SECTION("Custom Properties with XLDateTime")
    {
        XLDocument doc;
        doc.create("CustomPropertiesTest.xlsx", XLForceOverwrite);
        auto& customProps = doc.customProperties();

        customProps.setProperty("StringProp", "Hello");
        customProps.setProperty("IntProp", 42);
        customProps.setProperty("DoubleProp", 3.14159);
        customProps.setProperty("BoolProp", true);

        XLDateTime dateProp = XLDateTime::fromString("2025-01-01 12:00:00");
        customProps.setProperty("DateProp", dateProp);

        REQUIRE(customProps.property("StringProp") == "Hello");
        REQUIRE(customProps.property("IntProp") == "42");
        REQUIRE(customProps.property("DoubleProp") == "3.14159");
        REQUIRE(customProps.property("BoolProp") == "true");
        REQUIRE(customProps.property("DateProp") == "2025-01-01T12:00:00Z");

        doc.save();
        doc.close();
    }

    SECTION("App Properties Modern Defaults")
    {
        XLDocument doc;
        doc.create("AppPropertiesTest.xlsx", XLForceOverwrite);
        auto& appProps = doc.appProperties();

        // The template has its own defaults. Let's set ours and verify.
        appProps.setProperty("Application", "OpenXLSX");
        appProps.setProperty("AppVersion", "1.0");
        
        REQUIRE(appProps.property("Application") == "OpenXLSX");
        REQUIRE(appProps.property("AppVersion") == "1.0");

        doc.save();
        doc.close();
    }
}

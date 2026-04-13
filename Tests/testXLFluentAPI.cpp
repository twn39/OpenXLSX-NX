#include <OpenXLSX.hpp>
#include "XLComments.hpp"
#include "XLThreadedComments.hpp"

#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_testXLFluentAPI_fluent() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("test_fluent") + ".xlsx";
    return name;
}
} // namespace

TEST_CASE("Fluent API and Object-Oriented Comments", "[Fluent][DX]")
{
    SECTION("Chaining Value and Notes on XLCell")
    {
        XLDocument doc;
        doc.create(__global_unique_testXLFluentAPI_fluent(), XLForceOverwrite);
        
        // 1. Redundant Author Removal
        doc.setDefaultAuthor("System Manager");

        auto wks = doc.workbook().worksheet("Sheet1");

        // 2. Fluent Chaining on XLCell
        wks.cell("A1").value() = "Review Needed";
        wks.cell("A1").addNote("Please approve these numbers."); // Note: Author defaults to System Manager
        
        // 3. Object-Oriented Threaded Comment Replies
        wks.cell("C3").value() = "Sales Data";
        auto thread = wks.cell("C3").addComment("Is the Q3 data finalized?", "Alice");
        
        thread.addReply("I am uploading the final numbers now.", "Bob")
              .addReply("Got it, waiting for the upload.", "Alice")
              .setResolved(true);

        doc.save();
        doc.close();

        // Verification
        doc.open(__global_unique_testXLFluentAPI_fluent());
        wks = doc.workbook().worksheet("Sheet1");

        // Verify Legacy Note
        REQUIRE(wks.cell("A1").value().get<std::string>() == "Review Needed");
        REQUIRE(wks.comments().get("A1") == "Please approve these numbers.");
        
        // Verify Threaded Comments
        REQUIRE(wks.cell("C3").value().get<std::string>() == "Sales Data");
        auto tc = wks.threadedComments().comment("C3");
        REQUIRE(tc.valid());
        REQUIRE(tc.isResolved() == true);
        
        auto replies = wks.threadedComments().replies(tc.id());
        REQUIRE(replies.size() == 2);
        REQUIRE(replies[0].text() == "I am uploading the final numbers now.");
        REQUIRE(replies[1].text() == "Got it, waiting for the upload.");

        doc.close();
        std::remove(__global_unique_testXLFluentAPI_fluent().c_str());
    }
}

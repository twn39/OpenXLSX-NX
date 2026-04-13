#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>

using namespace OpenXLSX;

namespace { 
inline const std::string& __global_unique_file_0() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__dn_structure_test_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_1() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__DefinedNamesLocalTest_xlsx") + ".xlsx";
    return name;
}

inline const std::string& __global_unique_file_2() {
    static std::string name = OpenXLSX::TestHelpers::getUniqueFilename("__DefinedNamesTest_xlsx") + ".xlsx";
    return name;
}
} // namespace


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

TEST_CASE("XLDefinedNamesTests", "[XLDefinedNames]")
{
    SECTION("Manage Global Defined Names")
    {
        XLDocument doc;
        doc.create(__global_unique_file_2(), XLForceOverwrite);
        auto wb = doc.workbook();

        auto dns = wb.definedNames();
        REQUIRE(dns.count() == 0);

        dns.append("MyRange", "Sheet1!$A$1:$B$10");
        REQUIRE(dns.count() == 1);
        REQUIRE(dns.exists("MyRange"));

        auto dn = dns.get("MyRange");
        REQUIRE(dn.name() == "MyRange");
        REQUIRE(dn.refersTo() == "Sheet1!$A$1:$B$10");
        REQUIRE_FALSE(dn.localSheetId().has_value());

        dns.remove("MyRange");
        REQUIRE(dns.count() == 0);
        REQUIRE_FALSE(dns.exists("MyRange"));

        doc.save();
        doc.close();
    }

    SECTION("Manage Local Defined Names")
    {
        XLDocument doc;
        doc.create(__global_unique_file_1(), XLForceOverwrite);
        auto wb = doc.workbook();
        wb.addWorksheet("Sheet2");

        auto dns = wb.definedNames();

        // Add global name
        dns.append("GlobalName", "Sheet1!$A$1");

        // Add local names (localSheetId is 0-indexed in XML, corresponding to sheet index - 1 or internal ID)
        // Usually localSheetId refers to the index in the <sheets> list.
        dns.append("LocalName", "Sheet1!$B$1", 0);    // Local to Sheet1
        dns.append("LocalName", "Sheet2!$B$1", 1);    // Local to Sheet2

        REQUIRE(dns.count() == 3);
        REQUIRE(dns.exists("GlobalName"));
        REQUIRE(dns.exists("LocalName", 0));
        REQUIRE(dns.exists("LocalName", 1));

        auto dn1 = dns.get("LocalName", 0);
        REQUIRE(dn1.refersTo() == "Sheet1!$B$1");
        REQUIRE(dn1.localSheetId() == 0);

        auto dn2 = dns.get("LocalName", 1);
        REQUIRE(dn2.refersTo() == "Sheet2!$B$1");
        REQUIRE(dn2.localSheetId() == 1);

        doc.save();
        doc.close();

        // Re-open and verify
        XLDocument doc2;
        doc2.open(__global_unique_file_1());
        auto dns2 = doc2.workbook().definedNames();
        REQUIRE(dns2.count() == 3);
        REQUIRE(dns2.get("LocalName", 0).refersTo() == "Sheet1!$B$1");
        REQUIRE(dns2.get("LocalName", 1).refersTo() == "Sheet2!$B$1");
        doc2.close();
    }

    SECTION("OOXML Structure Verification")
    {
        XLDocument doc;
        doc.create(__global_unique_file_0(), XLForceOverwrite);
        auto wb = doc.workbook();
        wb.addWorksheet("Sheet2");

        auto dns = wb.definedNames();
        dns.append("GlobalName", "Sheet1!$A$1");
        dns.append("LocalName", "Sheet2!$B$1", 1);    // Local to Sheet2 (index 1)

        doc.save();

        // Use internal API to peek into xl/workbook.xml
        std::string workbookXml = getRawXml(doc, "xl/workbook.xml");

        // 1. Verify definedNames node exists
        REQUIRE(workbookXml.find("<definedNames>") != std::string::npos);

        // 2. Verify global name structure
        REQUIRE(workbookXml.find("<definedName name=\"GlobalName\">Sheet1!$A$1</definedName>") != std::string::npos);

        // 3. Verify local name structure (with localSheetId)
        REQUIRE(workbookXml.find("<definedName name=\"LocalName\" localSheetId=\"1\">Sheet2!$B$1</definedName>") != std::string::npos);

        // 4. Verify node ordering (definedNames should be after sheets)
        size_t sheetsPos = workbookXml.find("<sheets>");
        size_t dnPos     = workbookXml.find("<definedNames>");
        REQUIRE(sheetsPos != std::string::npos);
        REQUIRE(dnPos != std::string::npos);
        REQUIRE(sheetsPos < dnPos);

        doc.close();
    }
}

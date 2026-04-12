#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <fstream>
#include <vector>

using namespace OpenXLSX;

TEST_CASE("NegativeTestsCorruptFilesandExceptions", "[CorruptFiles]")
{
    SECTION("File Not Found")
    {
        XLDocument doc;
        REQUIRE_THROWS_AS(doc.open("non_existent_random_file_123.xlsx"), XLException);
    }

    SECTION("Not a ZIP Archive")
    {
        const std::string dummy_file = "dummy_text_file.xlsx";
        std::ofstream     out(dummy_file);
        out << "This is just a text file, not a valid ZIP or XLSX file.";
        out.close();

        XLDocument doc;
        REQUIRE_THROWS_AS(doc.open(dummy_file), XLException);

        std::remove(dummy_file.c_str());
    }

    SECTION("Valid ZIP but missing Content_Types.xml")
    {
        int res = system("echo 'hello' > dummy.txt && zip -q bad_archive.xlsx dummy.txt && rm dummy.txt");
        if (res == 0) {
            XLDocument doc;
            REQUIRE_THROWS_AS(doc.open("bad_archive.xlsx"), XLException);
            std::remove("bad_archive.xlsx");
        }
    }

    SECTION("Valid ZIP but malformed XML inside")
    {
        // 1. Create a valid document first
        const std::string malformed_zip_file = "malformed_xml.xlsx";
        {
            XLDocument doc;
            doc.create(malformed_zip_file, XLForceOverwrite);
            doc.save();
            doc.close();
        }

        // 2. We inject bad xml into [Content_Types].xml
        // We use python to copy the zip but replace the file, since zipfile 'a' mode appends a duplicate name.
        const char* py_script = "import zipfile, os\n"
                                "with zipfile.ZipFile('malformed_xml.xlsx', 'r') as zin:\n"
                                "    with zipfile.ZipFile('malformed_xml_bad.xlsx', 'w') as zout:\n"
                                "        for item in zin.infolist():\n"
                                "            if item.filename == '[Content_Types].xml':\n"
                                "                zout.writestr(item, '<Types "
                                "xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ')\n"
                                "            else:\n"
                                "                zout.writestr(item, zin.read(item.filename))\n"
                                "os.replace('malformed_xml_bad.xlsx', 'malformed_xml.xlsx')\n";

        std::ofstream script("corrupt.py");
        script << py_script;
        script.close();

        int res = system("python3 corrupt.py");

        if (res == 0) {
            XLDocument doc;
            // The XML parser should throw XLException when trying to parse the malformed XML
            REQUIRE_THROWS_AS(doc.open(malformed_zip_file), XLException);
        }
        std::remove(malformed_zip_file.c_str());
        std::remove("corrupt.py");
    }
}

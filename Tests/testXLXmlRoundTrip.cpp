#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <sstream>

using namespace OpenXLSX;

namespace {
    std::string nodeToString(const XMLNode& node) {
        std::ostringstream oss;
        node.print(oss, "", pugi::format_raw | pugi::format_no_declaration);
        return oss.str();
    }
}

TEST_CASE("XML Object Round-Trip Tests", "[XMLRoundTrip]")
{
    XMLDocument doc;

    SECTION("XLFont Round-Trip")
    {
        const char* inputXml = "<font><b val=\"1\"/><sz val=\"12\"/><color rgb=\"FF0000\"/><name val=\"Calibri\"/></font>";
        REQUIRE(doc.load_string(inputXml));
        
        XLFont font(doc.document_element());
        
        // Modify state
        font.setBold(false);
        font.setFontSize(14);
        font.setFontName("Arial");
        font.setFontColor(XLColor("00FF00"));

        std::string outputXml = nodeToString(doc.document_element());
        INFO("Output XML: " << outputXml);
        
        // Assert modified state via XML comparison
        REQUIRE(outputXml.find("<b val=\"0\"/>") != std::string::npos);
        REQUIRE(outputXml.find("<sz val=\"14\"/>") != std::string::npos);
        REQUIRE(outputXml.find("<color rgb=\"FF00FF00\"/>") != std::string::npos);
        REQUIRE(outputXml.find("<name val=\"Arial\"/>") != std::string::npos);
    }

    SECTION("XLFill Round-Trip")
    {
        const char* inputXml = "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFF0000\"/></patternFill></fill>";
        REQUIRE(doc.load_string(inputXml));

        XLFill fill(doc.document_element());
        REQUIRE(fill.patternType() == XLPatternSolid);
        REQUIRE(fill.color().hex() == "FFFF0000");

        fill.setPatternType(XLPatternDarkGray);
        fill.setColor(XLColor("00FF00"));

        std::string outputXml = nodeToString(doc.document_element());
        REQUIRE(outputXml.find("patternType=\"darkGray\"") != std::string::npos);
        REQUIRE(outputXml.find("<fgColor rgb=\"FF00FF00\"/>") != std::string::npos);
    }

    SECTION("XLBorder Round-Trip")
    {
        const char* inputXml = "<border><left style=\"thin\"><color rgb=\"FF000000\"/></left><right/><top/><bottom/><diagonal/></border>";
        REQUIRE(doc.load_string(inputXml));

        XLBorder border(doc.document_element());
        // Modify right border
        border.setRight(XLLineStyleMedium, XLColor("FF0000"));

        std::string outputXml = nodeToString(doc.document_element());
        REQUIRE(outputXml.find("<right style=\"medium\"><color rgb=\"FFFF0000\"/></right>") != std::string::npos);
    }

    SECTION("XLCfvo Round-Trip")
    {
        const char* inputXml = "<cfvo type=\"percent\" val=\"10\"/>";
        REQUIRE(doc.load_string(inputXml));

        XLCfvo cfvo(doc.document_element());
        
        REQUIRE(cfvo.type() == XLCfvoType::Percent);
        REQUIRE(cfvo.value() == "10");

        cfvo.setType(XLCfvoType::Max);
        // Max type clears value in OpenXLSX usually, or let's try Number
        cfvo.setType(XLCfvoType::Number);
        cfvo.setValue("50");

        std::string outputXml = nodeToString(doc.document_element());
        REQUIRE(outputXml == "<cfvo type=\"num\" val=\"50\"/>");
    }

    SECTION("XLDataValidation Round-Trip")
    {
        const char* inputXml = "<dataValidation type=\"whole\" operator=\"between\" allowBlank=\"1\" sqref=\"A1:A10\"><formula1>1</formula1><formula2>100</formula2></dataValidation>";
        REQUIRE(doc.load_string(inputXml));

        XLDataValidation dv(doc.document_element());

        // Modify
        dv.setOperator(XLDataValidationOperator::GreaterThan);
        dv.setFormula1("50");
        dv.setAllowBlank(false);

        std::string outputXml = nodeToString(doc.document_element());
        REQUIRE(outputXml.find("operator=\"greaterThan\"") != std::string::npos);
        REQUIRE(outputXml.find("allowBlank=\"0\"") != std::string::npos);
        REQUIRE(outputXml.find("<formula1>50</formula1>") != std::string::npos);
    }

    SECTION("XLCfRule Round-Trip")
    {
        const char* inputXml = "<cfRule type=\"cellIs\" operator=\"between\" priority=\"1\"><formula>10</formula><formula>20</formula></cfRule>";
        REQUIRE(doc.load_string(inputXml));

        XLCfRule rule(doc.document_element());
        REQUIRE(rule.type() == XLCfType::CellIs);
        REQUIRE(rule.Operator() == XLCfOperator::Between);
        REQUIRE(rule.priority() == 1);
        
        auto formulas = rule.formulas();
        REQUIRE(formulas.size() == 2);
        REQUIRE(formulas[0] == "10");

        // Modify
        rule.clearFormulas();
        rule.addFormula("50");

        std::string outputXml = nodeToString(doc.document_element());
        REQUIRE(outputXml.find("<formula>10</formula>") == std::string::npos);
        REQUIRE(outputXml.find("<formula>50</formula>") != std::string::npos);
    }
}
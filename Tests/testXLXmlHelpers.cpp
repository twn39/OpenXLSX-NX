//
// Unit tests for XLXmlHelpers (Layer-2 DOM ensure/set idioms).
//

#include <XLXmlHelpers.hpp>
#include <catch2/catch_test_macros.hpp>
#include <string>
#include <string_view>
#include <vector>

using namespace OpenXLSX;

TEST_CASE("XLXmlHelpers ensureChild get-or-create", "[XmlHelpers]")
{
    XMLDocument doc;
    XMLNode     root = doc.append_child("root");

    XMLNode a1 = ensureChild(root, "childA");
    REQUIRE_FALSE(a1.empty());
    REQUIRE(std::string_view(a1.name()) == "childA");
    REQUIRE(root.child_count_of_type(node_element) == 1);

    XMLNode a2 = ensureChild(root, "childA");
    REQUIRE(a2 == a1);
    REQUIRE(root.child_count_of_type(node_element) == 1);

    XMLNode emptyParent;
    REQUIRE(ensureChild(emptyParent, "x").empty());
}

TEST_CASE("XLXmlHelpers ensureChild respects nodeOrder", "[XmlHelpers]")
{
    XMLDocument doc;
    XMLNode     root = doc.append_child("root");
    const std::vector<std::string_view> order = {"a", "b", "c"};

    ensureChild(root, "c", order);
    ensureChild(root, "a", order);
    ensureChild(root, "b", order);

    XMLNode first = root.first_child_of_type(node_element);
    REQUIRE(std::string_view(first.name()) == "a");
    REQUIRE(std::string_view(first.next_sibling_of_type(node_element).name()) == "b");
    REQUIRE(std::string_view(first.next_sibling_of_type(node_element).next_sibling_of_type(node_element).name()) == "c");
}

TEST_CASE("XLXmlHelpers ensureAttr vs setAttr", "[XmlHelpers]")
{
    XMLDocument doc;
    XMLNode     root = doc.append_child("root");

    XMLAttribute d1 = ensureAttr(root, "flag", "default");
    REQUIRE(std::string_view(d1.value()) == "default");

    ensureAttr(root, "flag", "ignored");
    REQUIRE(std::string_view(root.attribute("flag").value()) == "default");

    setAttr(root, "flag", "updated");
    REQUIRE(std::string_view(root.attribute("flag").value()) == "updated");

    setAttrBool01(root, "on", true);
    REQUIRE(std::string_view(root.attribute("on").value()) == "1");
    setAttrBool01(root, "on", false);
    REQUIRE(std::string_view(root.attribute("on").value()) == "0");

    setAttr(root, "count", 42u);
    REQUIRE(root.attribute("count").as_uint() == 42u);
}

TEST_CASE("XLXmlHelpers setChildAttr setChildVal setChildText", "[XmlHelpers]")
{
    XMLDocument doc;
    XMLNode     root = doc.append_child("root");

    setChildAttr(root, "color", "rgb", "FF0000");
    REQUIRE(std::string_view(root.child("color").attribute("rgb").value()) == "FF0000");

    setChildVal(root, "majorUnit", "10");
    REQUIRE(std::string_view(root.child("majorUnit").attribute("val").value()) == "10");

    setChildText(root, "title", "Hello");
    REQUIRE(std::string_view(root.child("title").text().get()) == "Hello");
}

TEST_CASE("XLXmlHelpers ensureChildPath", "[XmlHelpers]")
{
    XMLDocument doc;
    XMLNode     root = doc.append_child("worksheet");

    XMLNode tab = ensureChildPath(root, {"sheetPr", "tabColor"});
    REQUIRE_FALSE(tab.empty());
    REQUIRE(std::string_view(tab.name()) == "tabColor");
    REQUIRE_FALSE(root.child("sheetPr").empty());

    setAttr(tab, "rgb", "FF00FF");
    REQUIRE(std::string_view(root.child("sheetPr").child("tabColor").attribute("rgb").value()) == "FF00FF");
}

TEST_CASE("XLXmlHelpers legacy appendAnd* aliases", "[XmlHelpers]")
{
    XMLDocument doc;
    XMLNode     root = doc.append_child("root");

    XMLNode n = appendAndGetNode(root, "item");
    REQUIRE_FALSE(n.empty());
    appendAndSetAttribute(n, "id", "7");
    REQUIRE(std::string_view(n.attribute("id").value()) == "7");

    appendAndSetNodeAttribute(root, "font", "val", "Arial");
    REQUIRE(std::string_view(root.child("font").attribute("val").value()) == "Arial");

    appendAndGetAttribute(n, "missing", "fallback");
    REQUIRE(std::string_view(n.attribute("missing").value()) == "fallback");
    appendAndGetAttribute(n, "missing", "nope");
    REQUIRE(std::string_view(n.attribute("missing").value()) == "fallback");
}

TEST_CASE("XLXmlHelpers getBoolAttributeWhenOmittedMeansTrue", "[XmlHelpers]")
{
    XMLDocument doc;
    XMLNode     root = doc.append_child("root");
    XMLNode     tag  = root.append_child("b");

    REQUIRE(getBoolAttributeWhenOmittedMeansTrue(root, "b"));
    REQUIRE(std::string_view(tag.attribute("val").value()) == "true");

    setAttr(tag, "val", "false");
    REQUIRE_FALSE(getBoolAttributeWhenOmittedMeansTrue(root, "b"));
}

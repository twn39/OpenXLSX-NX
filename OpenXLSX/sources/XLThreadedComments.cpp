#include "XLThreadedComments.hpp"

#include <random>
#include <sstream>
#include <iomanip>
#include "XLWorksheet.hpp"

using namespace OpenXLSX;

namespace {
    std::string GenerateGUID() {
        static std::random_device rd;
        static std::mt19937 gen(rd());
        std::uniform_int_distribution<> dis(0, 15);
        std::uniform_int_distribution<> dis2(8, 11);

        std::stringstream ss;
        ss << std::hex << std::uppercase << std::setfill('0');
        ss << "{";
        for (int i = 0; i < 8; i++) ss << dis(gen);
        ss << "-";
        for (int i = 0; i < 4; i++) ss << dis(gen);
        ss << "-4";
        for (int i = 0; i < 3; i++) ss << dis(gen);
        ss << "-";
        ss << dis2(gen);
        for (int i = 0; i < 3; i++) ss << dis(gen);
        ss << "-";
        for (int i = 0; i < 12; i++) ss << dis(gen);
        ss << "}";
        return ss.str();
    }
}

// ===== XLThreadedComment Implementation ===== //

XLThreadedComment::XLThreadedComment(XMLNode node, XLWorksheet* wks) : m_node(node), m_wks(wks) {}

bool XLThreadedComment::valid() const { return m_node != nullptr; }

XLThreadedComment& XLThreadedComment::addReply(const std::string& text, const std::string& author)
{
    if (m_wks && valid()) {
        m_wks->addReply(id(), text, author);
    }
    return *this;
}

std::string XLThreadedComment::ref() const
{
    return m_node.attribute("ref").value();
}

std::string XLThreadedComment::id() const
{
    return m_node.attribute("id").value();
}

std::string XLThreadedComment::parentId() const
{
    return m_node.attribute("parentId").value();
}

std::string XLThreadedComment::personId() const
{
    return m_node.attribute("personId").value();
}

std::string XLThreadedComment::text() const
{
    XMLNode textNode = m_node.child("text");
    if (!textNode.empty()) return textNode.text().get();
    return "";
}

bool XLThreadedComment::isResolved() const
{
    return m_node.attribute("done").as_bool(false);
}

void XLThreadedComment::setResolved(bool resolved)
{
    if (resolved) {
        auto attr = m_node.attribute("done");
        if (!attr) attr = m_node.append_attribute("done");
        attr.set_value("1");
    } else {
        m_node.remove_attribute("done");
    }
}

// ===== XLThreadedComments Implementation ===== //

XLThreadedComment XLThreadedComments::comment(const std::string& ref) const
{
    XMLNode root = xmlDocument().document_element();
    for (XMLNode node = root.first_child(); node; node = node.next_sibling()) {
        if (std::string(node.name()) == "threadedComment" && std::string(node.attribute("ref").value()) == ref) {
            // A top-level comment has no parentId
            if (node.attribute("parentId").empty()) {
                return XLThreadedComment(node);
            }
        }
    }
    return XLThreadedComment(XMLNode());
}

std::vector<XLThreadedComment> XLThreadedComments::replies(const std::string& parentId) const
{
    std::vector<XLThreadedComment> result;
    XMLNode root = xmlDocument().document_element();
    for (XMLNode node = root.first_child(); node; node = node.next_sibling()) {
        if (std::string(node.name()) == "threadedComment" && std::string(node.attribute("parentId").value()) == parentId) {
            result.emplace_back(node);
        }
    }
    return result;
}

XLThreadedComment XLThreadedComments::addComment(std::string_view ref, std::string_view personId, std::string_view text)
{
    XMLNode root = xmlDocument().document_element();
    if (!root) {
        root = xmlDocument().append_child("ThreadedComments");
        root.append_attribute("xmlns").set_value("http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments");
        root.append_attribute("xmlns:x").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    }
    XMLNode commentNode = root.append_child("threadedComment");
    commentNode.append_attribute("ref").set_value(std::string(ref).c_str());
    commentNode.append_attribute("dT").set_value("2024-01-01T12:00:00.00");
    commentNode.append_attribute("personId").set_value(std::string(personId).c_str());
    commentNode.append_attribute("id").set_value(GenerateGUID().c_str());
    
    XMLNode textNode = commentNode.append_child("text");
    textNode.text().set(std::string(text).c_str());

    return XLThreadedComment(commentNode);
}

XLThreadedComment XLThreadedComments::addReply(const std::string& parentId, const std::string& personId, const std::string& text)
{
    XMLNode root = xmlDocument().document_element();
    if (!root) {
        root = xmlDocument().append_child("ThreadedComments");
        root.append_attribute("xmlns").set_value("http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments");
        root.append_attribute("xmlns:x").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    }
    
    // Attempt to find the parent to inherit the ref attribute (though OOXML typically omits or matches it)
    std::string refStr = "";
    for (XMLNode n = root.first_child(); n; n = n.next_sibling()) {
        if (std::string(n.attribute("id").value()) == parentId) {
            refStr = n.attribute("ref").value();
            break;
        }
    }

    XMLNode commentNode = root.append_child("threadedComment");
    if (!refStr.empty()) {
        commentNode.append_attribute("ref").set_value(refStr.c_str());
    }
    commentNode.append_attribute("dT").set_value("2024-01-01T12:00:00.00");
    commentNode.append_attribute("personId").set_value(personId.c_str());
    commentNode.append_attribute("id").set_value(GenerateGUID().c_str());
    commentNode.append_attribute("parentId").set_value(parentId.c_str());
    
    XMLNode textNode = commentNode.append_child("text");
    textNode.text().set(text.c_str());

    return XLThreadedComment(commentNode);
}

bool XLThreadedComments::deleteComment(const std::string& ref)
{
    XMLNode root = xmlDocument().document_element();
    std::string rootId = "";
    
    // Find the parent ID first
    for (XMLNode node = root.first_child(); node; node = node.next_sibling()) {
        if (std::string(node.name()) == "threadedComment" && std::string(node.attribute("ref").value()) == ref) {
            if (node.attribute("parentId").empty()) {
                rootId = node.attribute("id").value();
                root.remove_child(node);
                break;
            }
        }
    }
    
    if (rootId.empty()) return false;

    // Now loop and remove all children
    XMLNode node = root.first_child();
    while (node) {
        XMLNode next = node.next_sibling();
        if (std::string(node.name()) == "threadedComment" && std::string(node.attribute("parentId").value()) == rootId) {
            root.remove_child(node);
        }
        node = next;
    }
    return true;
}

// ===== XLPerson Implementation ===== //

XLPerson::XLPerson(XMLNode node) : m_node(node) {}

bool XLPerson::valid() const { return m_node != nullptr; }

std::string XLPerson::id() const
{
    return m_node.attribute("id").value();
}

std::string XLPerson::displayName() const
{
    return m_node.attribute("displayName").value();
}

// ===== XLPersons Implementation ===== //

XLPerson XLPersons::person(const std::string& id) const
{
    XMLNode root = xmlDocument().document_element();
    for (XMLNode node = root.first_child(); node; node = node.next_sibling()) {
        if (std::string(node.name()) == "person" && std::string(node.attribute("id").value()) == id) {
            return XLPerson(node);
        }
    }
    return XLPerson(XMLNode());
}

std::string XLPersons::addPerson(const std::string& displayName)
{
    XMLNode root = xmlDocument().document_element();
    if (!root) {
        root = xmlDocument().append_child("personList");
        root.append_attribute("xmlns").set_value("http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments");
        root.append_attribute("xmlns:x").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    }

    // Check if exists
    for (XMLNode node = root.first_child(); node; node = node.next_sibling()) {
        if (std::string(node.name()) == "person" && std::string(node.attribute("displayName").value()) == displayName) {
            return node.attribute("id").value();
        }
    }

    // Add new
    XMLNode personNode = root.append_child("person");
    std::string newId = GenerateGUID();
    personNode.append_attribute("displayName").set_value(displayName.c_str());
    personNode.append_attribute("id").set_value(newId.c_str());
    personNode.append_attribute("userId").set_value(displayName.c_str()); // Mock generic user ID
    personNode.append_attribute("providerId").set_value("None");
    return newId;
}

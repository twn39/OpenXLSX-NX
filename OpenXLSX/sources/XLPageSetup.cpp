
// ===== External Includes ===== //

#include <vector>
#include <string_view>

static const std::vector<std::string_view> XLHeaderFooterNodeOrder = {
    "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"
};

#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLPageSetup.hpp"
#include "XLUtilities.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLPageMargins::XLPageMargins(const XMLNode& node) : m_node(node) {}

double XLPageMargins::left() const { return m_node.attribute("left").as_double(); }
XLPageMargins& XLPageMargins::setLeft(double value) { appendAndSetAttribute(m_node, "left", std::to_string(value)); return *this; }

double XLPageMargins::right() const { return m_node.attribute("right").as_double(); }
XLPageMargins& XLPageMargins::setRight(double value) { appendAndSetAttribute(m_node, "right", std::to_string(value)); return *this; }

double XLPageMargins::top() const { return m_node.attribute("top").as_double(); }
XLPageMargins& XLPageMargins::setTop(double value) { appendAndSetAttribute(m_node, "top", std::to_string(value)); return *this; }

double XLPageMargins::bottom() const { return m_node.attribute("bottom").as_double(); }
XLPageMargins& XLPageMargins::setBottom(double value) { appendAndSetAttribute(m_node, "bottom", std::to_string(value)); return *this; }

double XLPageMargins::header() const { return m_node.attribute("header").as_double(); }
XLPageMargins& XLPageMargins::setHeader(double value) { appendAndSetAttribute(m_node, "header", std::to_string(value)); return *this; }

double XLPageMargins::footer() const { return m_node.attribute("footer").as_double(); }
XLPageMargins& XLPageMargins::setFooter(double value) { appendAndSetAttribute(m_node, "footer", std::to_string(value)); return *this; }

/**
 * @details
 */
XLPrintOptions::XLPrintOptions(const XMLNode& node) : m_node(node) {}

bool XLPrintOptions::gridLines() const { return m_node.attribute("gridLines").as_bool(); }
void XLPrintOptions::setGridLines(bool value) { appendAndSetAttribute(m_node, "gridLines", value ? "1" : "0"); }

bool XLPrintOptions::headings() const { return m_node.attribute("headings").as_bool(); }
void XLPrintOptions::setHeadings(bool value) { appendAndSetAttribute(m_node, "headings", value ? "1" : "0"); }

bool XLPrintOptions::horizontalCentered() const { return m_node.attribute("horizontalCentered").as_bool(); }
void XLPrintOptions::setHorizontalCentered(bool value) { appendAndSetAttribute(m_node, "horizontalCentered", value ? "1" : "0"); }

bool XLPrintOptions::verticalCentered() const { return m_node.attribute("verticalCentered").as_bool(); }
void XLPrintOptions::setVerticalCentered(bool value) { appendAndSetAttribute(m_node, "verticalCentered", value ? "1" : "0"); }

/**
 * @details
 */
XLPageSetup::XLPageSetup(const XMLNode& node) : m_node(node) {}

uint32_t XLPageSetup::paperSize() const { return m_node.attribute("paperSize").as_uint(); }
XLPageSetup& XLPageSetup::setPaperSize(uint32_t value)
{
appendAndSetAttribute(m_node, "paperSize", std::to_string(value));;
    return *this;
}

XLPageOrientation XLPageSetup::orientation() const
{
    std::string val = m_node.attribute("orientation").value();
    if (val == "portrait") return XLPageOrientation::Portrait;
    if (val == "landscape") return XLPageOrientation::Landscape;
    return XLPageOrientation::Default;
}

XLPageSetup& XLPageSetup::setOrientation(XLPageOrientation value)
{
if (value == XLPageOrientation::Portrait)
        appendAndSetAttribute(m_node, "orientation", "portrait");
    else if (value == XLPageOrientation::Landscape)
        appendAndSetAttribute(m_node, "orientation", "landscape");
    else
        m_node.remove_attribute("orientation");;
    return *this;
}

uint32_t XLPageSetup::scale() const { return m_node.attribute("scale").as_uint(100); }
XLPageSetup& XLPageSetup::setScale(uint32_t value)
{
appendAndSetAttribute(m_node, "scale", std::to_string(value));;
    return *this;
}

uint32_t XLPageSetup::fitToWidth() const { return m_node.attribute("fitToWidth").as_uint(1); }
XLPageSetup& XLPageSetup::setFitToWidth(uint32_t value)
{
appendAndSetAttribute(m_node, "fitToWidth", std::to_string(value));;
    return *this;
}


uint32_t XLPageSetup::fitToHeight() const { return m_node.attribute("fitToHeight").as_uint(1); }
XLPageSetup& XLPageSetup::setFitToHeight(uint32_t value)
{
appendAndSetAttribute(m_node, "fitToHeight", std::to_string(value));;
    return *this;
}

std::string XLPageSetup::pageOrder() const { return m_node.attribute("pageOrder").value(); }
XLPageSetup& XLPageSetup::setPageOrder(std::string_view value)
{
appendAndSetAttribute(m_node, "pageOrder", std::string(value));;
    return *this;
}

bool XLPageSetup::useFirstPageNumber() const { return m_node.attribute("useFirstPageNumber").as_bool(); }
XLPageSetup& XLPageSetup::setUseFirstPageNumber(bool value)
{
appendAndSetAttribute(m_node, "useFirstPageNumber", value ? "1" : "0");;
    return *this;
}

uint32_t XLPageSetup::firstPageNumber() const { return m_node.attribute("firstPageNumber").as_uint(1); }
XLPageSetup& XLPageSetup::setFirstPageNumber(uint32_t value)
{
appendAndSetAttribute(m_node, "firstPageNumber", std::to_string(value));;
    return *this;
}

bool XLPageSetup::blackAndWhite() const { return m_node.attribute("blackAndWhite").as_bool(); }
XLPageSetup& XLPageSetup::setBlackAndWhite(bool value)
{
appendAndSetAttribute(m_node, "blackAndWhite", value ? "1" : "0");;
    return *this;
}

/**
 * @details
 */
XLHeaderFooter::XLHeaderFooter(const XMLNode& node) : m_node(node) {}

bool XLHeaderFooter::differentFirst() const { return m_node.attribute("differentFirst").as_bool(false); }
void XLHeaderFooter::setDifferentFirst(bool value) { appendAndSetAttribute(m_node, "differentFirst", value ? "1" : "0"); }

bool XLHeaderFooter::differentOddEven() const { return m_node.attribute("differentOddEven").as_bool(false); }
void XLHeaderFooter::setDifferentOddEven(bool value) { appendAndSetAttribute(m_node, "differentOddEven", value ? "1" : "0"); }

bool XLHeaderFooter::scaleWithDoc() const { return m_node.attribute("scaleWithDoc").as_bool(true); }
void XLHeaderFooter::setScaleWithDoc(bool value) { appendAndSetAttribute(m_node, "scaleWithDoc", value ? "1" : "0"); }

bool XLHeaderFooter::alignWithMargins() const { return m_node.attribute("alignWithMargins").as_bool(true); }
void XLHeaderFooter::setAlignWithMargins(bool value) { appendAndSetAttribute(m_node, "alignWithMargins", value ? "1" : "0"); }

std::string XLHeaderFooter::oddHeader() const { return m_node.child("oddHeader").text().get(); }
void XLHeaderFooter::setOddHeader(std::string_view value)
{
    auto node = m_node.child("oddHeader");
    if (node.empty()) node = appendAndGetNode(m_node, "oddHeader", XLHeaderFooterNodeOrder);
    node.text().set(std::string(value).c_str());
}

std::string XLHeaderFooter::oddFooter() const { return m_node.child("oddFooter").text().get(); }
void XLHeaderFooter::setOddFooter(std::string_view value)
{
    auto node = m_node.child("oddFooter");
    if (node.empty()) node = appendAndGetNode(m_node, "oddFooter", XLHeaderFooterNodeOrder);
    node.text().set(std::string(value).c_str());
}

std::string XLHeaderFooter::evenHeader() const { return m_node.child("evenHeader").text().get(); }
void XLHeaderFooter::setEvenHeader(std::string_view value)
{
    auto node = m_node.child("evenHeader");
    if (node.empty()) node = appendAndGetNode(m_node, "evenHeader", XLHeaderFooterNodeOrder);
    node.text().set(std::string(value).c_str());
}

std::string XLHeaderFooter::evenFooter() const { return m_node.child("evenFooter").text().get(); }
void XLHeaderFooter::setEvenFooter(std::string_view value)
{
    auto node = m_node.child("evenFooter");
    if (node.empty()) node = appendAndGetNode(m_node, "evenFooter", XLHeaderFooterNodeOrder);
    node.text().set(std::string(value).c_str());
}

std::string XLHeaderFooter::firstHeader() const { return m_node.child("firstHeader").text().get(); }
void XLHeaderFooter::setFirstHeader(std::string_view value)
{
    auto node = m_node.child("firstHeader");
    if (node.empty()) node = appendAndGetNode(m_node, "firstHeader", XLHeaderFooterNodeOrder);
    node.text().set(std::string(value).c_str());
}

std::string XLHeaderFooter::firstFooter() const { return m_node.child("firstFooter").text().get(); }
void XLHeaderFooter::setFirstFooter(std::string_view value)
{
    auto node = m_node.child("firstFooter");
    if (node.empty()) node = appendAndGetNode(m_node, "firstFooter", XLHeaderFooterNodeOrder);
    node.text().set(std::string(value).c_str());
}

XLPageMargins& XLPageMargins::setLeft(XLDistance value) { return setLeft(value.getInches()); }

XLPageMargins& XLPageMargins::setRight(XLDistance value) { return setRight(value.getInches()); }

XLPageMargins& XLPageMargins::setTop(XLDistance value) { return setTop(value.getInches()); }

XLPageMargins& XLPageMargins::setBottom(XLDistance value) { return setBottom(value.getInches()); }

XLPageMargins& XLPageMargins::setHeader(XLDistance value) { return setHeader(value.getInches()); }

XLPageMargins& XLPageMargins::setFooter(XLDistance value) { return setFooter(value.getInches()); }

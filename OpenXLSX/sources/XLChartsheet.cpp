#include "XLChartsheet.hpp"
#include "XLXmlData.hpp"

using namespace OpenXLSX;

XLChartsheet::XLChartsheet(XLXmlData* xmlData) : XLSheetBase(xmlData) {}
XLChartsheet::~XLChartsheet() = default;

XLColor XLChartsheet::getColor_impl() const { return XLColor(); } // TODO: Implement
void XLChartsheet::setColor_impl(const XLColor& color) { (void)color; } // TODO: Implement
bool XLChartsheet::isSelected_impl() const { return false; } // TODO: Implement
void XLChartsheet::setSelected_impl(bool selected) { (void)selected; } // TODO: Implement

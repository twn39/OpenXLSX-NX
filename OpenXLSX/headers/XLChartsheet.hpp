#ifndef OPENXLSX_XLCHARTSHEET_HPP
#define OPENXLSX_XLCHARTSHEET_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLSheetBase.hpp"

namespace OpenXLSX
{
    /**
     * @brief Class representing the an Excel chartsheet.
     * @todo This class is largely unimplemented and works just as a placeholder.
     */
    class OPENXLSX_EXPORT XLChartsheet final : public XLSheetBase<XLChartsheet>
    {
        friend class XLSheetBase<XLChartsheet>;

    public:
        XLChartsheet() : XLSheetBase(nullptr) {};
        explicit XLChartsheet(XLXmlData* xmlData);
        XLChartsheet(const XLChartsheet& other) = default;
        XLChartsheet(XLChartsheet&& other) noexcept = default;
        ~XLChartsheet();

        XLChartsheet& operator=(const XLChartsheet& other) = default;
        XLChartsheet& operator=(XLChartsheet&& other) noexcept = default;

    private:
        XLColor getColor_impl() const;
        void setColor_impl(const XLColor& color);
        bool isSelected_impl() const;
        void setSelected_impl(bool selected);
        bool isActive_impl() const { return false; }
        bool setActive_impl() { return true; }
    };
}

#endif

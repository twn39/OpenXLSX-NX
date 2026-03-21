#ifndef OPENXLSX_XLCHARTSHEET_HPP
#define OPENXLSX_XLCHARTSHEET_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLSheetBase.hpp"
#include "XLChart.hpp"
#include "XLDrawing.hpp"

namespace OpenXLSX
{
    /**
     * @brief Class representing an Excel chartsheet.
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

        /**
         * @brief Add or replace the chart in this chartsheet
         * @param type The type of chart to create
         * @param name The name of the chart object
         * @return The created XLChart object
         */
        XLChart addChart(XLChartType type, std::string_view name);

        /**
         * @brief Get the underlying drawing for this chartsheet
         */
        XLDrawing& drawing();
        XLRelationships& relationships();

    private:
        XLColor getColor_impl() const;
        void setColor_impl(const XLColor& color);
        bool isSelected_impl() const;
        void setSelected_impl(bool selected);
        bool isActive_impl() const { return false; }
        bool setActive_impl() { return true; }

        uint16_t sheetXmlNumber() const;

        XLRelationships m_relationships{};
        XLDrawing       m_drawing{};
    };
}

#endif

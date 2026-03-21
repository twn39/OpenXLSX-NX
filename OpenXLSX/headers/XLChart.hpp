#ifndef OPENXLSX_XLCHART_HPP
#define OPENXLSX_XLCHART_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include <string>
#include <string_view>

namespace OpenXLSX
{
    enum class XLChartType {
        Bar,
        BarStacked,
        BarPercentStacked,
        Bar3D,
        Bar3DStacked,
        Bar3DPercentStacked,
        Line,
        LineStacked,
        LinePercentStacked,
        Line3D,
        Pie,
        Pie3D,
        Scatter,
        Area,
        AreaStacked,
        AreaPercentStacked,
        Area3D,
        Area3DStacked,
        Area3DPercentStacked,
        Doughnut
    };

    enum class XLLegendPosition {
        Bottom,
        Left,
        Right,
        Top,
        TopRight,
        Hidden
    };

    /**
     * @brief The XLChart class represents an Excel chart XML file.
     */
    
    /**
     * @brief A proxy class for an individual chart axis.
     */
    class OPENXLSX_EXPORT XLAxis
    {
    public:
        XLAxis() = default;
        explicit XLAxis(const XMLNode& node);

        void setTitle(std::string_view title);
        
        void setMinBounds(double min);
        void clearMinBounds();
        
        void setMaxBounds(double max);
        void clearMaxBounds();
        
        void setMajorGridlines(bool show);
        void setMinorGridlines(bool show);

    private:
        XMLNode m_node;
    };

    class OPENXLSX_EXPORT XLChart final : public XLXmlFile
    {
    public:
        /**
         * @brief Constructor
         */
        XLChart() : XLXmlFile(nullptr) {}

        /**
         * @brief Constructor from XLXmlData
         * @param xmlData the source XML data
         */
        explicit XLChart(XLXmlData* xmlData);

        /**
         * @brief Copy Constructor
         * @param other The object to be copied
         */
        XLChart(const XLChart& other) = default;

        /**
         * @brief Move Constructor
         */
        XLChart(XLChart&& other) noexcept = default;

        /**
         * @brief Destructor
         */
        ~XLChart() = default;

        /**
         * @brief Copy Assignment Operator
         */
        XLChart& operator=(const XLChart& other) = default;

        /**
         * @brief Move Assignment Operator
         */
        XLChart& operator=(XLChart&& other) noexcept = default;

        /**
         * @brief Add a data series to the chart
         * @param valuesRef A cell reference to the data values (e.g. "Sheet1!$B$1:$B$10")
         * @param title A literal string or cell reference for the series name (e.g. "Revenue" or "Sheet1!$B$1")
         * @param categoriesRef A cell reference for the X-axis categories (e.g. "Sheet1!$A$1:$A$10")
         */
        void addSeries(std::string_view valuesRef, std::string_view title = "", std::string_view categoriesRef = "");

        /**
         * @brief Set the chart title
         * @param title The title text
         */
        void setTitle(std::string_view title);

        /**
         * @brief Set the legend position or hide it
         * @param position The legend position
         */
        void setLegendPosition(XLLegendPosition position);

        /**
         * @brief Get the X-axis (typically bottom category axis).
         */
        [[nodiscard]] XLAxis xAxis() const;

        /**
         * @brief Get the Y-axis (typically left value axis).
         */
        [[nodiscard]] XLAxis yAxis() const;
        
        [[nodiscard]] XLAxis axis(std::string_view position) const;


        /**
         * @brief Configure the display of data labels on the chart.
         * @param showValue Whether to show the data value.
         * @param showCategory Whether to show the category name.
         * @param showPercent Whether to show the percentage (useful for Pie/Doughnut charts).
         */
        void setShowDataLabels(bool showValue, bool showCategory = false, bool showPercent = false);


    private:
        friend class XLDocument;
        void initXml(XLChartType type = XLChartType::Bar);
        [[nodiscard]] uint32_t seriesCount() const;
    };
}

#endif // OPENXLSX_XLCHART_HPP

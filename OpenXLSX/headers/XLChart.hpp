#ifndef OPENXLSX_XLCHART_HPP
#define OPENXLSX_XLCHART_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLConstants.hpp"
#include "XLXmlFile.hpp"

#include <optional>
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
        ScatterLine,            ///< Scatter with straight lines, no markers
        ScatterLineMarker,      ///< Scatter with straight lines and markers
        ScatterSmooth,          ///< Scatter with smooth lines, no markers
        ScatterSmoothMarker,    ///< Scatter with smooth lines and markers
        ScatterMarker,          ///< Scatter with markers only, no lines
        Bubble,                 ///< Bubble chart (needs addBubbleSeries)
        Area,
        AreaStacked,
        AreaPercentStacked,
        Area3D,
        Area3DStacked,
        Area3DPercentStacked,
        Doughnut,
        Radar,
        RadarFilled,
        RadarMarkers
    };

    enum class XLLegendPosition { Bottom, Left, Right, Top, TopRight, Hidden };

    enum class XLMarkerStyle { None, Circle, Dash, Diamond, Dot, Picture, Plus, Square, Star, Triangle, X, Default };

    /**
     * @brief The XLChart class represents an Excel chart XML file.
     */

    /**
     * @brief A proxy class for an individual chart axis.
     */
    /**
     * @brief A proxy class for an individual chart series.
     */
    class XLWorksheet;
    class XLCellRange;

    enum class XLErrorBarDirection { X, Y };

    enum class XLErrorBarType { Both, Minus, Plus };

    enum class XLErrorBarValueType { Custom, FixedValue, Percentage, StandardDeviation, StandardError };

    enum class XLTrendlineType { Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage };

    class OPENXLSX_EXPORT XLChartSeries
    {
    public:
        XLChartSeries() = default;
        explicit XLChartSeries(const XMLNode& node);

        XLChartSeries& setTitle(std::string_view title);
        XLChartSeries& setSmooth(bool smooth);
        XLChartSeries& setMarkerStyle(XLMarkerStyle style);

        /**
         * @brief Set the fill and line color of this data series.
         * @param hexRGB Six-character hex color string, e.g. "FF0000" for red.
         */
        XLChartSeries& setColor(std::string_view hexRGB);

        /**
         * @brief Override the color of a single data point within this series.
         * @param pointIdx Zero-based index of the data point.
         * @param hexRGB  Six-character hex color string.
         */
        XLChartSeries& setDataPointColor(uint32_t pointIdx, std::string_view hexRGB);

        /**
         * @brief Enable and configure data labels for this series.
         * @param showValue     Show the cell value on each label.
         * @param showCategoryName Show the category name.
         * @param showPercent   Show percentage (mainly for pie/doughnut charts).
         */
        XLChartSeries& setDataLabels(bool showValue, bool showCategoryName = false, bool showPercent = false);

        /**
         * @brief Add a trendline to this series.
         */
        XLChartSeries& addTrendline(XLTrendlineType type, std::string_view name = "", uint8_t order = 2, uint8_t period = 2);

        /**
         * @brief Add error bars to this series.
         */
        XLChartSeries& addErrorBars(XLErrorBarDirection direction, XLErrorBarType type, XLErrorBarValueType valType, double value = 0.0);

    private:
        XMLNode m_node;
    };

    struct XLChartAnchor
    {
        std::string name;
        uint32_t    row{1};
        uint32_t    col{1};
        XLDistance  width{XLDistance::Pixels(400)};
        XLDistance  height{XLDistance::Pixels(300)};

        XLChartAnchor() = default;
        XLChartAnchor(std::string_view n, uint32_t r, uint32_t c, XLDistance w, XLDistance h) : name(n), row(r), col(c), width(w), height(h)
        {}
    };

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

        /**
         * @brief Set the axis to use a logarithmic scale.
         * @param base The base of the logarithm (e.g. 10). If 0 or less, the logarithmic scale is removed.
         */
        void setLogScale(double base);

        /**
         * @brief Convert this axis to a date axis or back to a category axis.
         * @param isDateAxis If true, the axis becomes a c:dateAx. If false, it reverts to c:catAx.
         */
        void setDateAxis(bool isDateAxis);

        void setMajorGridlines(bool show);
        void setMinorGridlines(bool show);

        /**
         * @brief Set the number format for axis tick labels.
         * @param formatCode    Excel number format code, e.g. "0.00%" or "#,##0".
         * @param sourceLinked  If true, the format follows the source data format.
         */
        void setNumberFormat(std::string_view formatCode, bool sourceLinked = false);

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
         * @brief Add a data series to the chart.
         * @param valuesRef A cell reference to the data values (e.g. "Sheet1!$B$1:$B$10")
         * @param title A literal string or cell reference for the series name.
         * @param categoriesRef A cell reference for the X-axis categories.
         */
        XLChartSeries addSeries(const XLWorksheet&         wks,
                                const XLCellRange&         values,
                                std::string_view           title            = "",
                                std::optional<XLChartType> targetChartType  = std::nullopt,
                                bool                       useSecondaryAxis = false);
        XLChartSeries addSeries(const XLWorksheet&         wks,
                                const XLCellRange&         values,
                                const XLCellRange&         categories,
                                std::string_view           title            = "",
                                std::optional<XLChartType> targetChartType  = std::nullopt,
                                bool                       useSecondaryAxis = false);
        XLChartSeries addSeries(std::string_view           valuesRef,
                                std::string_view           title            = "",
                                std::string_view           categoriesRef    = "",
                                std::optional<XLChartType> targetChartType  = std::nullopt,
                                bool                       useSecondaryAxis = false);

        /**
         * @brief Add a bubble chart series with explicit X, Y, and size ranges.
         * @param xValRef   Cell reference for X values (e.g. "Sheet1!$A$2:$A$10").
         * @param yValRef   Cell reference for Y values.
         * @param sizeRef   Cell reference for bubble sizes.
         * @param title     Optional series name.
         */
        XLChartSeries
            addBubbleSeries(std::string_view xValRef, std::string_view yValRef, std::string_view sizeRef, std::string_view title = "");

        XLChartSeries addBubbleSeries(const XLWorksheet& wks,
                                      const XLCellRange& xValues,
                                      const XLCellRange& yValues,
                                      const XLCellRange& sizes,
                                      std::string_view   title = "");

        /**
         * @brief Set the chart title.
         */
        void setTitle(std::string_view title);

        /**
         * @brief Set the built-in chart style ID (1–48).
         */
        void setStyle(uint8_t styleId);

        /**
         * @brief Set the legend position or hide it.
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
         */
        void setShowDataLabels(bool showValue, bool showCategory = false, bool showPercent = false);

        /**
         * @brief Set whether a specific series should be rendered with a smooth line.
         */
        void setSeriesSmooth(uint32_t seriesIndex, bool smooth);

        /**
         * @brief Set the marker style for a specific series.
         */
        void setSeriesMarker(uint32_t seriesIndex, XLMarkerStyle style);

        /**
         * @brief Set the gap width between bar/column clusters.
         * @param percent Gap expressed as a percentage of bar width (0–500, default 150).
         */
        void setGapWidth(uint32_t percent);

        /**
         * @brief Set the overlap between bars/columns within a cluster.
         * @param percent Overlap in percent (−100 to 100). Positive = overlap, negative = gap.
         */
        void setOverlap(int32_t percent);

        /**
         * @brief Fill the chart plot area with a solid color.
         * @param hexRGB Six-character hex color, e.g. "F2F2F2".
         */
        void setPlotAreaColor(std::string_view hexRGB);

        /**
         * @brief Fill the outermost chart space background with a solid color.
         * @param hexRGB Six-character hex color.
         */
        void setChartAreaColor(std::string_view hexRGB);

    private:
        friend class XLDocument;
        void                   initXml(XLChartType type = XLChartType::Bar);
        [[nodiscard]] uint32_t seriesCount() const;
    };
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLCHART_HPP

#ifndef OPENXLSX_XLCHART_HPP
#define OPENXLSX_XLCHART_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include <string>
#include <string_view>

namespace OpenXLSX
{
    enum class XLChartType {
        Bar
    };

    /**
     * @brief The XLChart class represents an Excel chart XML file.
     */
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
         * @param valuesRef A cell reference to the data values (e.g. "Sheet1!$A$1:$A$10")
         */
        void addSeries(std::string_view valuesRef);

    private:
        void initXml();
        [[nodiscard]] uint32_t seriesCount() const;
    };
}

#endif // OPENXLSX_XLCHART_HPP

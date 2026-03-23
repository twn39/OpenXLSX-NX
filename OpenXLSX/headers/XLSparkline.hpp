#ifndef OPENXLSX_XLSPARKLINE_HPP
#define OPENXLSX_XLSPARKLINE_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include <string>

namespace OpenXLSX
{
    /**
     * @brief The XLSparklineType enum represents the type of a sparkline.
     */
    enum class XLSparklineType { Line, Column, Stacked };

    /**
     * @brief The XLSparkline class encapsulates a single sparkline group in a worksheet.
     */
    class OPENXLSX_EXPORT XLSparkline
    {
    public:
        XLSparkline();
        explicit XLSparkline(XMLNode node);
        ~XLSparkline();

        XLSparklineType type() const;
        void setType(XLSparklineType type);

        std::string location() const;
        void setLocation(const std::string& sqref);

        std::string dataRange() const;
        void setDataRange(const std::string& formula);

    private:
        XMLNode m_node;
    };

} // namespace OpenXLSX

#endif // OPENXLSX_XLSPARKLINE_HPP
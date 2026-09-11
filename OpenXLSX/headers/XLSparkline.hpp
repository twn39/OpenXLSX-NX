/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

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
     * @brief The XLSparklineOptions struct encapsulates configuration for a sparkline group.
     */
    struct XLSparklineOptions
    {
        XLSparklineType type {XLSparklineType::Line};
        std::string     seriesColor {"FF376092"}; // Default blue-ish
        std::string     negativeColor {"FFD00000"}; // Default red
        std::string     markersColor {"FFD00000"};
        std::string     firstMarkerColor {"FFD00000"};
        std::string     lastMarkerColor {"FFD00000"};
        std::string     highMarkerColor {"FFD00000"};
        std::string     lowMarkerColor {"FFD00000"};

        bool            markers {false}; // Show markers
        bool            high {false};    // Highlight highest point
        bool            low {false};     // Highlight lowest point
        bool            first {false};   // Highlight first point
        bool            last {false};    // Highlight last point
        bool            negative {false};// Highlight negative points
        bool            displayXAxis {false}; // Show horizontal axis
        std::string     displayEmptyCellsAs {"gap"}; // gap, zero, span
    };

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
        void            setType(XLSparklineType type);

        std::string location() const;
        void        setLocation(const std::string& sqref);

        std::string dataRange() const;
        void        setDataRange(const std::string& formula);

    private:
        XMLNode m_node;
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLSPARKLINE_HPP
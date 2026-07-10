#ifndef OPENXLSX_XLFORMULAUTILS_HPP
#define OPENXLSX_XLFORMULAUTILS_HPP

/**
 * @file XLFormulaUtils.hpp
 * @brief Shared conversion and error helpers for the formula evaluation pipeline.
 *
 * @details Grouped under OpenXLSX::formula to keep the top-level namespace clean.
 *          Compatibility using-declarations re-export symbols into OpenXLSX so
 *          existing function implementations (XLFormulaMath, XLFormulaLogical, …)
 *          continue to compile unchanged.
 */

#include "XLCellValue.hpp"
#include "XLFormulaEngine.hpp"

#include <string>
#include <string_view>
#include <vector>

namespace OpenXLSX
{
    namespace formula
    {
        // =====================================================================
        // Value conversion (pure; no workbook dependency)
        // =====================================================================

        /** Convert Integer / Float / Boolean to double; NaN for other types. */
        [[nodiscard]] double toDouble(const XLCellValue& v);

        [[nodiscard]] bool isNumeric(const XLCellValue& v);
        [[nodiscard]] bool isEmpty(const XLCellValue& v);
        [[nodiscard]] bool isError(const XLCellValue& v);

        /** Stringify a cell value for text functions; empty for non-displayable types. */
        [[nodiscard]] std::string toString(const XLCellValue& v);

        // =====================================================================
        // Excel error values
        // =====================================================================

        /** Build an error XLCellValue with the given Excel error token (e.g. "#VALUE!"). */
        [[nodiscard]] XLCellValue makeError(std::string_view token);

        [[nodiscard]] inline XLCellValue errValue() { return makeError("#VALUE!"); }
        [[nodiscard]] inline XLCellValue errDiv0() { return makeError("#DIV/0!"); }
        [[nodiscard]] inline XLCellValue errNA() { return makeError("#N/A"); }
        [[nodiscard]] inline XLCellValue errNum() { return makeError("#NUM!"); }
        [[nodiscard]] inline XLCellValue errRef() { return makeError("#REF!"); }
        [[nodiscard]] inline XLCellValue errName() { return makeError("#NAME?"); }

        // =====================================================================
        // Argument helpers
        // =====================================================================

        [[nodiscard]] std::string strTrim(std::string s);

        /** Flatten numeric values from formula arguments (skips non-numeric cells). */
        [[nodiscard]] std::vector<double> numerics(const std::vector<XLFormulaArg>& args);
        [[nodiscard]] std::vector<double> numerics(const XLFormulaArg& arg);

    }    // namespace formula

    // ----- Compatibility exports (historical call sites use OpenXLSX::toDouble, etc.) -----

    using formula::toDouble;
    using formula::isNumeric;
    using formula::isEmpty;
    using formula::isError;
    using formula::toString;
    using formula::makeError;
    using formula::errValue;
    using formula::errDiv0;
    using formula::errNA;
    using formula::errNum;
    using formula::errRef;
    using formula::errName;
    using formula::strTrim;
    using formula::numerics;

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLFORMULAUTILS_HPP

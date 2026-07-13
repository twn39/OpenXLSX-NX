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

#include <cstdint>
#include <optional>
#include <random>
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
        // Argument helpers / Excel coercion (Phase A)
        // =====================================================================

        [[nodiscard]] std::string strTrim(std::string s);

        /**
         * @brief Parse a fully-numeric string (optional leading/trailing space) to double.
         * @return true when the entire trimmed string is a valid number.
         */
        [[nodiscard]] bool tryParseNumericString(std::string_view s, double& out);

        /**
         * @brief Excel-style scalar numeric coercion for math functions and arithmetic.
         * @details Error → same error (propagated). Empty → 0 when @p blankAsZero, else #VALUE!.
         *          Integer/Float/Boolean → numeric. Fully-numeric string → number. Else #VALUE!.
         * @return A Float/Integer cell value, or an Error cell value.
         */
        [[nodiscard]] XLCellValue coerceToNumber(const XLCellValue& v, bool blankAsZero = true);

        /**
         * @brief Coerce the first scalar of an argument (implicit intersection / top-left).
         * @return Error if the arg is missing/empty (no cells), or coerceToNumber of asScalar().
         */
        [[nodiscard]] XLCellValue coerceArgScalar(const XLFormulaArg& arg, bool blankAsZero = true);

        /**
         * @brief First error value found in argument cells (Excel left-to-right, top-to-bottom).
         * @return Error value if any, otherwise Empty.
         */
        [[nodiscard]] XLCellValue firstError(const XLFormulaArg& arg);
        [[nodiscard]] XLCellValue firstError(const std::vector<XLFormulaArg>& args);

        /** Flatten numeric values from formula arguments (skips non-numeric cells and errors). */
        [[nodiscard]] std::vector<double> numerics(const std::vector<XLFormulaArg>& args);
        [[nodiscard]] std::vector<double> numerics(const XLFormulaArg& arg);

        /**
         * @brief Like numerics(), but if any cell is an Error, sets @p outError and returns {}.
         * @details Used by SUM/AVERAGE/MIN/MAX so errors short-circuit instead of being skipped.
         */
        [[nodiscard]] std::vector<double> numericsOrError(const std::vector<XLFormulaArg>& args, XLCellValue& outError);
        [[nodiscard]] std::vector<double> numericsOrError(const XLFormulaArg& arg, XLCellValue& outError);

        // =====================================================================
        // Deterministic RNG for RAND / RANDBETWEEN / RANDARRAY (Phase D)
        // =====================================================================

        /**
         * @brief Seed the formula RNG (thread-local engines reseed on next use).
         * @details When set, RAND/RANDBETWEEN/RANDARRAY produce reproducible sequences
         *          within each thread. Call clearFormulaRandomSeed() to restore
         *          non-deterministic seeding via std::random_device.
         */
        void setFormulaRandomSeed(uint64_t seed);

        /** Clear the optional seed (next RNG construction uses random_device). */
        void clearFormulaRandomSeed();

        /** Currently configured seed, if any. */
        [[nodiscard]] std::optional<uint64_t> formulaRandomSeed();

        /** Thread-local mt19937_64 used by formula random functions. */
        [[nodiscard]] std::mt19937_64& formulaRng();

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
    using formula::tryParseNumericString;
    using formula::coerceToNumber;
    using formula::coerceArgScalar;
    using formula::firstError;
    using formula::numerics;
    using formula::numericsOrError;
    using formula::setFormulaRandomSeed;
    using formula::clearFormulaRandomSeed;
    using formula::formulaRandomSeed;
    using formula::formulaRng;

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLFORMULAUTILS_HPP

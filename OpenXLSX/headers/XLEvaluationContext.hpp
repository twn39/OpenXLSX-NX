#ifndef OPENXLSX_XLEVALUATIONCONTEXT_HPP
#define OPENXLSX_XLEVALUATIONCONTEXT_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"

#include <string>
#include <string_view>
#include <unordered_map>

namespace OpenXLSX
{
    class XLWorksheet;

    /**
     * @brief Abstract sheet-read port for formula evaluation.
     *
     * @details Decouples `XLFormulaEngine` from concrete workbook types. Implementations
     *          resolve a single-cell reference string (e.g. "A1", "$B$2", "Sheet1!C3",
     *          or quoted `'My Sheet'!A1`) to an `XLCellValue`. Range expansion remains
     *          the engine's responsibility; only single-cell refs are requested.
     *
     *          Unknown / out-of-range cells should return an empty `XLCellValue{}`.
     */
    class OPENXLSX_EXPORT XLEvaluationContext
    {
    public:
        virtual ~XLEvaluationContext() = default;

        /**
         * @brief Resolve a single-cell reference to its live (or mapped) value.
         * @param ref Raw reference text as seen by the formula engine.
         */
        [[nodiscard]] virtual XLCellValue cellValue(std::string_view ref) const = 0;
    };

    /**
     * @brief In-memory evaluation context for unit tests and offline formula checks.
     */
    class OPENXLSX_EXPORT XLMapEvaluationContext final : public XLEvaluationContext
    {
    public:
        XLMapEvaluationContext() = default;
        explicit XLMapEvaluationContext(std::unordered_map<std::string, XLCellValue> cells);

        void set(std::string_view ref, XLCellValue value);

        [[nodiscard]] XLCellValue cellValue(std::string_view ref) const override;

    private:
        std::unordered_map<std::string, XLCellValue> m_cells;
    };

    /**
     * @brief Evaluation context backed by a live worksheet (and its workbook for sheet-qualified refs).
     * @note Valid only while the referenced worksheet remains alive.
     */
    class OPENXLSX_EXPORT XLWorksheetEvaluationContext final : public XLEvaluationContext
    {
    public:
        explicit XLWorksheetEvaluationContext(const XLWorksheet& wks) noexcept;

        [[nodiscard]] XLCellValue cellValue(std::string_view ref) const override;

    private:
        const XLWorksheet* m_wks{nullptr};
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLEVALUATIONCONTEXT_HPP

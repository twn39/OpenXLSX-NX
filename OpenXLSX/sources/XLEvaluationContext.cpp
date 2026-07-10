#include "XLEvaluationContext.hpp"

#include "XLDocument.hpp"
#include "XLSheet.hpp"
#include "XLWorkbook.hpp"
#include "XLWorksheet.hpp"

#include <string>

namespace OpenXLSX
{
    namespace
    {
        /**
         * @brief Strip optional surrounding single quotes from a sheet name.
         * @details Excel emits quoted names for sheets with spaces/special characters: 'My Sheet'!A1
         */
        std::string_view unquoteSheetName(std::string_view name)
        {
            if (name.size() >= 2 && name.front() == '\'' && name.back() == '\'') {
                return name.substr(1, name.size() - 2);
            }
            return name;
        }
    }    // namespace

    // =============================================================================
    // XLMapEvaluationContext
    // =============================================================================

    XLMapEvaluationContext::XLMapEvaluationContext(std::unordered_map<std::string, XLCellValue> cells) : m_cells(std::move(cells)) {}

    void XLMapEvaluationContext::set(std::string_view ref, XLCellValue value) { m_cells[std::string(ref)] = std::move(value); }

    XLCellValue XLMapEvaluationContext::cellValue(std::string_view ref) const
    {
        auto it = m_cells.find(std::string(ref));
        if (it != m_cells.end()) return it->second;
        return XLCellValue{};
    }

    // =============================================================================
    // XLWorksheetEvaluationContext
    // =============================================================================

    XLWorksheetEvaluationContext::XLWorksheetEvaluationContext(const XLWorksheet& wks) noexcept : m_wks(&wks) {}

    XLCellValue XLWorksheetEvaluationContext::cellValue(std::string_view ref) const
    {
        if (m_wks == nullptr) return XLCellValue{};

        try {
            auto bang = ref.find('!');
            if (bang != std::string_view::npos) {
                const std::string_view sheetPart = unquoteSheetName(ref.substr(0, bang));
                const std::string      cellAddr  = std::string(ref.substr(bang + 1));

                // Cross-sheet: resolve through the parent workbook when possible.
                try {
                    XLWorksheet other = m_wks->parentDoc().workbook().worksheet(sheetPart);
                    return other.cell(cellAddr).value();
                }
                catch (...) {
                    // Fall back to the bound sheet if the name is missing/invalid.
                    return m_wks->cell(cellAddr).value();
                }
            }

            return m_wks->cell(std::string(ref)).value();
        }
        catch (...) {
            return XLCellValue{};
        }
    }

}    // namespace OpenXLSX

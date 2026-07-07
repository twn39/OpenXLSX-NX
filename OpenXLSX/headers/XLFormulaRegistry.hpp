#ifndef OPENXLSX_XLFORMULAREGISTRY_HPP
#define OPENXLSX_XLFORMULAREGISTRY_HPP

#include <memory>
#include <string>
#include <string_view>
#include <ankerl/unordered_dense.h>
#include "XLFormulaFunction.hpp"

namespace OpenXLSX
{
    struct FormulaStringViewHash {
        using is_transparent = void;
        auto operator()(std::string_view str) const noexcept -> uint64_t {
            return ankerl::unordered_dense::hash<std::string_view>{}(str);
        }
    };

    class XLFormulaRegistry
    {
    public:
        static XLFormulaRegistry& getInstance();

        void registerFunction(std::string name, std::shared_ptr<XLFormulaFunction> func);

        const XLFormulaFunction* lookup(std::string_view name) const;

        const ankerl::unordered_dense::map<std::string, std::shared_ptr<XLFormulaFunction>, FormulaStringViewHash, std::equal_to<>>& getFunctions() const;

    private:
        XLFormulaRegistry() = default;
        ankerl::unordered_dense::map<std::string, std::shared_ptr<XLFormulaFunction>, FormulaStringViewHash, std::equal_to<>> m_functions;
    };

    // Linker-safe group registration declarations
    void registerMathFunctions(XLFormulaRegistry& registry);
    void registerTextFunctions(XLFormulaRegistry& registry);
    void registerLogicalFunctions(XLFormulaRegistry& registry);
    void registerDateTimeFunctions(XLFormulaRegistry& registry);
    void registerFinancialFunctions(XLFormulaRegistry& registry);
    void registerStatisticalFunctions(XLFormulaRegistry& registry);
    void registerInfoFunctions(XLFormulaRegistry& registry);
}

#endif // OPENXLSX_XLFORMULAREGISTRY_HPP

#include "XLFormulaRegistry.hpp"
#include <mutex>
#include <algorithm>
#include <cctype>

namespace OpenXLSX
{
    XLFormulaRegistry& XLFormulaRegistry::getInstance()
    {
        static XLFormulaRegistry instance;
        static std::once_flag flag;
        std::call_once(flag, []() {
            registerMathFunctions(instance);
            registerTextFunctions(instance);
            registerLogicalFunctions(instance);
            registerDateTimeFunctions(instance);
            registerFinancialFunctions(instance);
            registerStatisticalFunctions(instance);
            registerInfoFunctions(instance);
        });
        return instance;
    }

    void XLFormulaRegistry::registerFunction(std::string name, std::shared_ptr<XLFormulaFunction> func)
    {
        std::transform(name.begin(), name.end(), name.begin(), [](unsigned char c) {
            return std::toupper(c);
        });
        m_functions[std::move(name)] = std::move(func);
    }

    const XLFormulaFunction* XLFormulaRegistry::lookup(std::string_view name) const
    {
        bool has_lowercase = false;
        for (char c : name) {
            if (c >= 'a' && c <= 'z') {
                has_lowercase = true;
                break;
            }
        }
        
        if (has_lowercase) {
            std::string upper_name(name);
            std::transform(upper_name.begin(), upper_name.end(), upper_name.begin(), [](unsigned char c) {
                return std::toupper(c);
            });
            auto it = m_functions.find(upper_name);
            if (it != m_functions.end()) {
                return it->second.get();
            }
        } else {
            auto it = m_functions.find(name);
            if (it != m_functions.end()) {
                return it->second.get();
            }
        }
        return nullptr;
    }

    const ankerl::unordered_dense::map<std::string, std::shared_ptr<XLFormulaFunction>, FormulaStringViewHash, std::equal_to<>>& XLFormulaRegistry::getFunctions() const
    {
        return m_functions;
    }
}

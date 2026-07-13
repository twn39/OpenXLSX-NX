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
            registerArrayFunctions(instance);
            registerEngineeringFunctions(instance);
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
        auto tryFind = [this](std::string_view key) -> const XLFormulaFunction* {
            auto it = m_functions.find(key);
            if (it != m_functions.end()) return it->second.get();
            return nullptr;
        };

        std::string upper_name(name);
        bool        has_lowercase = false;
        for (char c : name) {
            if (c >= 'a' && c <= 'z') {
                has_lowercase = true;
                break;
            }
        }
        if (has_lowercase) {
            std::transform(upper_name.begin(), upper_name.end(), upper_name.begin(), [](unsigned char c) {
                return static_cast<char>(std::toupper(c));
            });
        }
        else {
            upper_name.assign(name);
        }

        // Strip Excel future-function prefixes after uppercasing
        auto strip = [](std::string& n, const char* prefix) {
            const std::size_t len = std::char_traits<char>::length(prefix);
            if (n.size() > len && n.compare(0, len, prefix) == 0) n.erase(0, len);
        };
        strip(upper_name, "_XLFN.");
        strip(upper_name, "_XLWS.");

        if (auto* f = tryFind(upper_name)) return f;
        // Also try original case key for all-uppercase inputs without transform path
        if (!has_lowercase) {
            if (auto* f = tryFind(name)) return f;
        }
        return nullptr;
    }

    const ankerl::unordered_dense::map<std::string, std::shared_ptr<XLFormulaFunction>, FormulaStringViewHash, std::equal_to<>>& XLFormulaRegistry::getFunctions() const
    {
        return m_functions;
    }
}

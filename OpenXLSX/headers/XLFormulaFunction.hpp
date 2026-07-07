#ifndef OPENXLSX_XLFORMULAFUNCTION_HPP
#define OPENXLSX_XLFORMULAFUNCTION_HPP

#include <vector>
#include "XLCellValue.hpp"
#include "XLFormulaEngine.hpp"

namespace OpenXLSX
{
    class XLFormulaFunction
    {
    public:
        virtual ~XLFormulaFunction() = default;
        virtual XLCellValue execute(const std::vector<XLFormulaArg>& args) const = 0;
    };

    class XLSimpleFormulaFunction : public XLFormulaFunction
    {
    public:
        using FuncPtr = XLCellValue (*)(const std::vector<XLFormulaArg>&);

        explicit XLSimpleFormulaFunction(FuncPtr ptr) : m_ptr(ptr) {}

        XLCellValue execute(const std::vector<XLFormulaArg>& args) const override
        {
            return m_ptr(args);
        }

    private:
        FuncPtr m_ptr;
    };
}

#endif // OPENXLSX_XLFORMULAFUNCTION_HPP

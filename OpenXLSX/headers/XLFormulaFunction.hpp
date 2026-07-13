#ifndef OPENXLSX_XLFORMULAFUNCTION_HPP
#define OPENXLSX_XLFORMULAFUNCTION_HPP

#include <vector>
#include "XLCellValue.hpp"
#include "XLFormulaEngine.hpp"

namespace OpenXLSX
{
    /**
     * @brief Abstract formula function interface (session-aware, array-capable).
     *
     * @details Phase B: execute returns XLFormulaArg so dynamic-array functions
     *          (FILTER, UNIQUE, SORT, SEQUENCE, …) can produce multi-cell results.
     *          Pure scalar builtins use XLSimpleFormulaFunction which wraps XLCellValue.
     */
    class XLFormulaFunction
    {
    public:
        virtual ~XLFormulaFunction() = default;
        virtual XLFormulaArg execute(const std::vector<XLFormulaArg>& args, XLEvalSession& session) const = 0;
    };

    /**
     * @brief Adapter for pure scalar functions that only inspect argument values.
     */
    class XLSimpleFormulaFunction : public XLFormulaFunction
    {
    public:
        using FuncPtr = XLCellValue (*)(const std::vector<XLFormulaArg>&);

        explicit XLSimpleFormulaFunction(FuncPtr ptr) : m_ptr(ptr) {}

        XLFormulaArg execute(const std::vector<XLFormulaArg>& args, XLEvalSession& /*session*/) const override
        {
            return XLFormulaArg(m_ptr(args));
        }

    private:
        FuncPtr m_ptr;
    };

    /**
     * @brief Adapter for session-aware scalar functions (return single XLCellValue).
     */
    class XLSessionFormulaFunction : public XLFormulaFunction
    {
    public:
        using FuncPtr = XLCellValue (*)(const std::vector<XLFormulaArg>&, XLEvalSession&);

        explicit XLSessionFormulaFunction(FuncPtr ptr) : m_ptr(ptr) {}

        XLFormulaArg execute(const std::vector<XLFormulaArg>& args, XLEvalSession& session) const override
        {
            return XLFormulaArg(m_ptr(args, session));
        }

    private:
        FuncPtr m_ptr;
    };

    /**
     * @brief Adapter for array-returning functions (FILTER, UNIQUE, SORT, …).
     */
    class XLArrayFormulaFunction : public XLFormulaFunction
    {
    public:
        using FuncPtr = XLFormulaArg (*)(const std::vector<XLFormulaArg>&, XLEvalSession&);

        explicit XLArrayFormulaFunction(FuncPtr ptr) : m_ptr(ptr) {}

        XLFormulaArg execute(const std::vector<XLFormulaArg>& args, XLEvalSession& session) const override
        {
            return m_ptr(args, session);
        }

    private:
        FuncPtr m_ptr;
    };
}

#endif // OPENXLSX_XLFORMULAFUNCTION_HPP

#ifndef OPENXLSX_XLDATAVALIDATION_HPP
#define OPENXLSX_XLDATAVALIDATION_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

#include <string>
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLDataValidation
    {
    public:
        XLDataValidation() : m_node(nullptr) {}
        explicit XLDataValidation(const XMLNode& node) : m_node(node) {}

        bool empty() const { return !m_node; }

        void setSqref(const std::string& sqref);
        void setWholeNumberRange(double min, double max);
        void setPrompt(const std::string& title, const std::string& msg);
        void setError(const std::string& title, const std::string& msg);

    private:
        mutable XMLNode m_node;
    };
} // namespace OpenXLSX

#pragma warning(pop)

#endif // OPENXLSX_XLDATAVALIDATION_HPP

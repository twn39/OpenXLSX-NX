#ifndef OPENXLSX_XLTHREADEDCOMMENTS_HPP
#define OPENXLSX_XLTHREADEDCOMMENTS_HPP

#include <string>
#include <gsl/pointers>
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    /**
     * @brief A class encapsulating modern Excel threaded comments (ThreadedComments.xml)
     */
    class OPENXLSX_EXPORT XLThreadedComments : public XLXmlFile
    {
    public:
        XLThreadedComments() : XLXmlFile(nullptr) {}
        explicit XLThreadedComments(gsl::not_null<XLXmlData*> xmlData) : XLXmlFile(xmlData) {}
        ~XLThreadedComments() = default;
        XLThreadedComments(const XLThreadedComments& other) = default;
        XLThreadedComments(XLThreadedComments&& other) = default;
        XLThreadedComments& operator=(const XLThreadedComments& other) = default;
        XLThreadedComments& operator=(XLThreadedComments&& other) = default;
    };

    /**
     * @brief A class encapsulating modern Excel persons metadata (persons.xml)
     */
    class OPENXLSX_EXPORT XLPersons : public XLXmlFile
    {
    public:
        XLPersons() : XLXmlFile(nullptr) {}
        explicit XLPersons(gsl::not_null<XLXmlData*> xmlData) : XLXmlFile(xmlData) {}
        ~XLPersons() = default;
        XLPersons(const XLPersons& other) = default;
        XLPersons(XLPersons&& other) = default;
        XLPersons& operator=(const XLPersons& other) = default;
        XLPersons& operator=(XLPersons&& other) = default;
    };
} // namespace OpenXLSX

#endif // OPENXLSX_XLTHREADEDCOMMENTS_HPP
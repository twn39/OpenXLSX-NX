/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLXMLSAVINGDECLARATION_HPP
#define OPENXLSX_XLXMLSAVINGDECLARATION_HPP

#include "OpenXLSX-Exports.hpp"
#include <string>

namespace OpenXLSX
{
    constexpr const char* XLXmlDefaultVersion  = "1.0";
    constexpr const char* XLXmlDefaultEncoding = "UTF-8";
    constexpr const bool  XLXmlStandalone      = true;
    constexpr const bool  XLXmlNotStandalone   = false;

    /**
     * @brief The XLXmlSavingDeclaration class encapsulates the properties of an XML saving declaration,
     * that can be used in calls to XLXmlData::getRawData to enforce specific settings
     */
    class OPENXLSX_EXPORT XLXmlSavingDeclaration
    {
    public:
        // ===== PUBLIC MEMBER FUNCTIONS ===== //
        XLXmlSavingDeclaration() : m_version(XLXmlDefaultVersion), m_encoding(XLXmlDefaultEncoding), m_standalone(XLXmlStandalone) {}
        XLXmlSavingDeclaration(XLXmlSavingDeclaration const& other) = default;    // copy constructor
        XLXmlSavingDeclaration(std::string version, std::string encoding, bool standalone = XLXmlStandalone)
            : m_version(std::move(version)),
              m_encoding(std::move(encoding)),
              m_standalone(standalone)
        {}
        ~XLXmlSavingDeclaration() = default;

        /**
         * @brief: getter functions: version, encoding, standalone
         */
        [[nodiscard]] std::string const& version() const { return m_version; }
        [[nodiscard]] std::string const& encoding() const { return m_encoding; }
        [[nodiscard]] bool               standalone_as_bool() const { return m_standalone; }
        [[nodiscard]] std::string        standalone() const { return m_standalone ? "yes" : "no"; }

    private:
        // ===== PRIVATE MEMBER VARIABLES ===== //
        std::string m_version;
        std::string m_encoding;
        bool        m_standalone;
    };
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLXMLSAVINGDECLARATION_HPP

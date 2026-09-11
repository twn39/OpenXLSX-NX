/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLEXCEPTION_HPP
#define OPENXLSX_XLEXCEPTION_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <stdexcept>    // std::runtime_error
#include <string>       // std::string - Issue #278 should be resolved by this

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"

namespace OpenXLSX
{
    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLException : public std::runtime_error
    {
    public:
        explicit XLException(const std::string& err) : runtime_error(err) {};
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLOverflowError : public XLException
    {
    public:
        explicit XLOverflowError(const std::string& err) : XLException(err) {};
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLValueTypeError : public XLException
    {
    public:
        explicit XLValueTypeError(const std::string& err) : XLException(err) {};
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLCellAddressError : public XLException
    {
    public:
        explicit XLCellAddressError(const std::string& err) : XLException(err) {};
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLInputError : public XLException
    {
    public:
        explicit XLInputError(const std::string& err) : XLException(err) {};
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLInternalError : public XLException
    {
    public:
        explicit XLInternalError(const std::string& err) : XLException(err) {};
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLPropertyError : public XLException
    {
    public:
        explicit XLPropertyError(const std::string& err) : XLException(err) {};
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLSheetError : public XLException
    {
    public:
        explicit XLSheetError(const std::string& err) : XLException(err) {};
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLDateTimeError : public XLException
    {
    public:
        explicit XLDateTimeError(const std::string& err) : XLException(err) {};
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLFormulaError : public XLException
    {
    public:
        explicit XLFormulaError(const std::string& err) : XLException(err) {};
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLEXCEPTION_HPP

/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#include "XLSharedStringTable.hpp"

namespace OpenXLSX
{
    XLMapSharedStringTable::XLMapSharedStringTable(std::vector<std::string> strings)
    {
        m_strings.reserve(strings.size());
        for (auto& s : strings) append(s);
    }

    int32_t XLMapSharedStringTable::append(std::string_view str)
    {
        auto it = m_index.find(std::string(str));
        if (it != m_index.end()) return it->second;

        const auto idx = static_cast<int32_t>(m_strings.size());
        m_strings.emplace_back(str);
        m_index.emplace(m_strings.back(), idx);
        return idx;
    }

    void XLMapSharedStringTable::clear()
    {
        m_strings.clear();
        m_index.clear();
    }

    int32_t XLMapSharedStringTable::stringCount() const { return static_cast<int32_t>(m_strings.size()); }

    int32_t XLMapSharedStringTable::getStringIndex(std::string_view str) const
    {
        auto it = m_index.find(std::string(str));
        return it != m_index.end() ? it->second : -1;
    }

    bool XLMapSharedStringTable::stringExists(std::string_view str) const
    {
        return m_index.find(std::string(str)) != m_index.end();
    }

    const char* XLMapSharedStringTable::getString(int32_t index) const
    {
        if (index < 0 || static_cast<size_t>(index) >= m_strings.size()) return "";
        return m_strings[static_cast<size_t>(index)].c_str();
    }

    std::string_view XLMapSharedStringTable::getStringView(int32_t index) const
    {
        if (index < 0 || static_cast<size_t>(index) >= m_strings.size()) return {};
        return m_strings[static_cast<size_t>(index)];
    }

}    // namespace OpenXLSX

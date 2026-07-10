#ifndef OPENXLSX_XLSHAREDSTRINGTABLE_HPP
#define OPENXLSX_XLSHAREDSTRINGTABLE_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

/**
 * @file XLSharedStringTable.hpp
 * @brief Read-only shared-string table (SST) port.
 *
 * @details Stream readers and other consumers that only resolve SST indices to
 *          text should depend on this interface, not on the full XLDocument
 *          façade or the mutable XLSharedStrings implementation details.
 *
 *          Write paths (getOrCreateStringIndex / appendString / clearString)
 *          remain on XLSharedStrings via document.sharedStrings().
 *
 *          Prefer XLXmlFile::sharedStringTable() over parentDoc().sharedStrings()
 *          when only lookups are required.
 */

#include "OpenXLSX-Exports.hpp"

#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

namespace OpenXLSX
{
    /**
     * @brief Abstract read-only view of the workbook shared string table.
     * @note Implemented by XLDocument (delegates to its XLSharedStrings).
     *       Lifetime must outlive consumers.
     */
    class OPENXLSX_EXPORT XLSharedStringTable
    {
    public:
        virtual ~XLSharedStringTable() = default;

        /** @return Number of entries currently in the SST cache. */
        [[nodiscard]] virtual int32_t stringCount() const = 0;

        /**
         * @brief Look up an existing string's index.
         * @return Index, or a negative value when the string is not present.
         */
        [[nodiscard]] virtual int32_t getStringIndex(std::string_view str) const = 0;

        /** @return true if @p str is already registered in the SST. */
        [[nodiscard]] virtual bool stringExists(std::string_view str) const = 0;

        /**
         * @brief Resolve an SST index to a null-terminated C string.
         * @note Pointer is valid while the document / SST remains alive.
         */
        [[nodiscard]] virtual const char* getString(int32_t index) const = 0;

        /**
         * @brief Resolve an SST index to a string_view into the arena/cache.
         * @note View is valid while the document / SST remains alive.
         */
        [[nodiscard]] virtual std::string_view getStringView(int32_t index) const = 0;
    };

    /**
     * @brief In-memory read-only SST for unit tests and offline consumers (no document required).
     * @details Mirrors XLMapEvaluationContext: a lightweight mock of the SST port.
     *          Appended strings keep stable storage so getString/getStringView remain valid.
     */
    class OPENXLSX_EXPORT XLMapSharedStringTable final : public XLSharedStringTable
    {
    public:
        XLMapSharedStringTable() = default;

        /** @brief Replace the table with an ordered list of strings (indices 0..n-1). */
        explicit XLMapSharedStringTable(std::vector<std::string> strings);

        /** @brief Append a string and return its index (duplicates reuse the existing index). */
        int32_t append(std::string_view str);

        void clear();

        [[nodiscard]] int32_t            stringCount() const override;
        [[nodiscard]] int32_t            getStringIndex(std::string_view str) const override;
        [[nodiscard]] bool               stringExists(std::string_view str) const override;
        [[nodiscard]] const char*        getString(int32_t index) const override;
        [[nodiscard]] std::string_view   getStringView(int32_t index) const override;

    private:
        std::vector<std::string>                 m_strings;
        std::unordered_map<std::string, int32_t> m_index;
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLSHAREDSTRINGTABLE_HPP

#ifndef OPENXLSX_XLSHAREDSTRINGS_HPP
#define OPENXLSX_XLSHAREDSTRINGS_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

#include <deque>
#include <functional>      // std::reference_wrapper
#include <limits>          // std::numeric_limits
#include <ostream>         // std::basic_ostream
#include <shared_mutex>    // std::shared_mutex
#include <string>
#include <ankerl/unordered_dense.h> // O(1) ankerl::unordered_dense string lookup

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLStringArena.hpp"
#include "XLXmlFile.hpp"
#include <string_view>
#include <vector>

namespace OpenXLSX
{
    struct StringViewHash {
        using is_transparent = void;
        auto operator()(std::string_view str) const noexcept -> uint64_t {
            return ankerl::unordered_dense::hash<std::string_view>{}(str);
        }
    };

    // Alias for flat hash map to allow easy substitution in the future
    template<typename Key, typename Value>
    using FlatHashMap = ankerl::unordered_dense::map<Key, Value, StringViewHash, std::equal_to<>>;

    constexpr size_t XLMaxSharedStrings = (std::numeric_limits<int32_t>::max)();    // pull request #261: wrapped max in parentheses to
                                                                                    // prevent expansion of windows.h "max" macro

    struct XLSharedStringsState {
        XLStringArena                               arena{};
        std::vector<std::string_view>               cache{};
        FlatHashMap<std::string_view, int32_t>      index{};
        std::unique_ptr<std::shared_mutex>          mutex{std::make_unique<std::shared_mutex>()};

        void clear() {
            arena.clear();
            cache.clear();
            index.clear();
            // mutex remains intact
        }
    };

    class XLSharedStrings;    // forward declaration
    typedef std::reference_wrapper<const XLSharedStrings> XLSharedStringsRef;

    extern const XLSharedStrings
        XLSharedStringsDefaulted;    // to be used for default initialization of all references of type XLSharedStrings

    /**
     * @brief This class encapsulate the Excel concept of Shared Strings. In Excel, instead of havig individual strings
     * in each cell, cells have a reference to an entry in the SharedStrings register. This results in smalle file
     * sizes, as repeated strings are referenced easily.
     */
    class OPENXLSX_EXPORT XLSharedStrings : public XLXmlFile
    {
        //---------- Friend Declarations ----------//
        friend class XLDocument;    // for access to protected function rewriteXmlFromCache

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief
         */
        XLSharedStrings() = default;

        /**
         * @brief
         * @param xmlData
         * @param state Pointer to the shared strings state
         */
        explicit XLSharedStrings(XLXmlData*            xmlData,
                                 XLSharedStringsState* state);

        /**
         * @brief Destructor
         */
        ~XLSharedStrings();

        /**
         * @brief
         * @param other
         */
        XLSharedStrings(const XLSharedStrings& other) = default;

        /**
         * @brief
         * @param other
         */
        XLSharedStrings(XLSharedStrings&& other) noexcept = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLSharedStrings& operator=(const XLSharedStrings& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLSharedStrings& operator=(XLSharedStrings&& other) noexcept = default;

        /**
         * @brief return the amount of shared string entries currently in the cache
         * @return
         */
        int32_t stringCount() const
        {
            if (!m_state) return 0;
            if (m_state->mutex) {
                std::shared_lock<std::shared_mutex> lock(*m_state->mutex);
                return static_cast<int32_t>(m_state->cache.size());
            }
            return static_cast<int32_t>(m_state->cache.size());
        }

        /**
         * @brief
         * @param str
         * @return
         */
        int32_t getStringIndex(std::string_view str) const;

        /**
         * @brief
         * @param str
         * @return
         */
        bool stringExists(std::string_view str) const;

        /**
         * @brief
         * @param index
         * @return
         */
        const char* getString(int32_t index) const;

        /**
         * @brief Append a new string to the list of shared strings.
         * @param str The string to append.
         * @return An int32_t with the index of the appended string
         */
        int32_t appendString(std::string_view str) const;

        /**
         * @brief Pre-reserve capacity in the string cache and index for n strings.
         * @details Call before bulk-inserting many strings to avoid incremental
         *          reallocation.  Does not affect the pugi XML DOM.
         * @param n Number of strings to reserve capacity for.
         */
        void reserveStrings(size_t n) const;

        /**
         * @brief Approximate memory used by the cache and hash index structures.
         * @return Bytes consumed by m_stringCache and m_stringIndex (not the arena).
         */
        size_t memoryUsageBytes() const noexcept;

        /**
         * @brief Get or create a string index in O(1) time.
         * @param str The string to look up or add.
         * @return The index of the string (existing or newly added).
         * @note This is the optimized method for setting cell string values.
         */
        int32_t getOrCreateStringIndex(std::string_view str) const;

        /**
         * @brief Clear the string at the given index.
         * @param index The index to clear.
         * @note There is no 'deleteString' member function, as deleting a shared string node will invalidate the
         * shared string indices for the cells in the spreadsheet. Instead use this member functions, which clears
         * the contents of the string, but keeps the XMLNode holding the string.
         */
        void clearString(int32_t index) const;

        /**
         * @brief print the XML contents of the shared strings document using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;

    protected:
        /**
         * @brief clear & rewrite the full shared strings XML from the shared strings cache
         * @return the amount of strings written to XML (should be equal to m_stringCache->size())
         */
        int32_t rewriteXmlFromCache();

        /**
         * @brief Rebuilds the shared string arena and cache using the provided index mapping,
         * keeping only the used strings, and rewrites the XML.
         * @param indexMap Mapping from old string indices to new string indices.
         * @param newStringCount The number of used strings (including empty string at index 0).
         */
        void rebuild(const std::vector<int32_t>& indexMap, int32_t newStringCount);

    private:
        XLSharedStringsState* m_state{}; /** < Pointer to the shared strings state (arena, cache, index, mutex) */

        /**
         * @brief Tracks whether the pugi DOM is behind the string cache.
         * @details When true, appendString() skips DOM mutation; the DOM is
         *          rebuilt from the cache in rewriteXmlFromCache() at save time.
         */
        mutable bool m_domDirty{false};
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLSHAREDSTRINGS_HPP

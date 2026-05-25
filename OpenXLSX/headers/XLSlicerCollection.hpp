#ifndef OPENXLSX_XLSLICERCOLLECTION_HPP
#define OPENXLSX_XLSLICERCOLLECTION_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLSlicer.hpp"
#include "XLTables.hpp"      // XLTable, XLSlicerOptions (legacy)
#include "XLPivotTable.hpp"  // XLPivotTable
#include <string>
#include <string_view>
#include <vector>

namespace OpenXLSX
{
    class XLWorksheet;

    // ─────────────────────────────────────────────────────────────────────────
    // XLSlicerBuilder
    //
    // Returned by XLSlicerCollection::add(). Accumulates configuration using a
    // fluent API, then writes the complete XML in one shot on destruction (RAII)
    // or when build() is called explicitly.
    // ─────────────────────────────────────────────────────────────────────────
    class OPENXLSX_EXPORT XLSlicerBuilder
    {
    public:
        // Internal: called only by XLSlicerCollection::add()
        XLSlicerBuilder(XLWorksheet*     worksheet,
                        std::string      cellRef,
                        const XLTable*   table,
                        const XLPivotTable* pivot,
                        std::string      columnName);
        ~XLSlicerBuilder();

        // Non-copyable, movable
        XLSlicerBuilder(const XLSlicerBuilder&) = delete;
        XLSlicerBuilder& operator=(const XLSlicerBuilder&) = delete;
        XLSlicerBuilder(XLSlicerBuilder&&) noexcept;
        XLSlicerBuilder& operator=(XLSlicerBuilder&&) = delete;

        // ── Fluent configuration ─────────────────────────────────────────────
        /// Override the slicer's internal unique name (defaults to columnName).
        XLSlicerBuilder& name(std::string_view n);
        /// Set the visible header caption (defaults to columnName).
        XLSlicerBuilder& caption(std::string_view c);
        /// Set style via strongly-typed enum.
        XLSlicerBuilder& style(XLSlicerStyle s);
        /// Set style via raw Excel style string.
        XLSlicerBuilder& styleRaw(std::string_view rawName);
        /// Set size in pixels.
        XLSlicerBuilder& size(uint32_t widthPx, uint32_t heightPx);
        /// Pre-select only these items (hide all others).
        XLSlicerBuilder& showOnly(const std::vector<std::string>& items);
        /// Set number of button columns in the slicer panel.
        XLSlicerBuilder& columnCount(int cols);
        /// Sort items descending (Z→A). Default: ascending.
        XLSlicerBuilder& sortDescending(bool desc = true);
        /// Lock the slicer position so it can't be moved interactively.
        XLSlicerBuilder& lockedPosition(bool locked = true);
        /// Pixel offset from anchor cell (fine-grained positioning).
        XLSlicerBuilder& offset(int32_t dx, int32_t dy);

        /// Explicitly commit and return the resulting XLSlicer handle.
        /// After this call, destruction is a no-op.
        XLSlicer build();

        /// Implicit conversion — allows: XLSlicer s = wks.slicers().add(...).caption("x");
        operator XLSlicer();

    private:
        void commit();

        struct State {
            std::string name;
            std::string caption;
            std::string styleRaw   = "";   // Empty = no style attribute (Excel default)
            uint32_t    widthPx    = 144;
            uint32_t    heightPx   = 200;
            int32_t     offsetX    = 0;
            int32_t     offsetY    = 0;
            int         columnCount = 0;  // 0 = unset
            bool        sortDesc   = false;
            bool        locked     = false;
            std::vector<std::string> selectedItems;
        };

        XLWorksheet*        m_worksheet;
        const XLTable*      m_table;
        const XLPivotTable* m_pivot;
        std::string         m_cellRef;
        std::string         m_columnName;
        State               m_state;
        bool                m_committed{false};
        XLSlicer            m_result;
    };

    // ─────────────────────────────────────────────────────────────────────────
    // XLSlicerCollection
    //
    // Manages all slicers on a single worksheet.  Returned by
    // XLWorksheet::slicers().  Provides STL-compatible iteration and named
    // / indexed access.  Replaces the old std::vector<XLSlicer>.
    // ─────────────────────────────────────────────────────────────────────────
    class OPENXLSX_EXPORT XLSlicerCollection
    {
    public:
        // ── Iterator ────────────────────────────────────────────────────────
        class Iterator
        {
        public:
            using iterator_category = std::forward_iterator_tag;
            using value_type        = XLSlicer;
            using difference_type   = std::ptrdiff_t;
            using pointer           = XLSlicer*;
            using reference         = XLSlicer&;

            explicit Iterator(std::vector<XLSlicer>* vec, size_t idx)
                : m_vec(vec), m_idx(idx) {}

            XLSlicer& operator*()  { return (*m_vec)[m_idx]; }
            XLSlicer* operator->() { return &(*m_vec)[m_idx]; }
            Iterator& operator++() { ++m_idx; return *this; }
            bool operator==(const Iterator& o) const { return m_idx == o.m_idx; }
            bool operator!=(const Iterator& o) const { return m_idx != o.m_idx; }
        private:
            std::vector<XLSlicer>* m_vec;
            size_t                 m_idx;
        };

        class ConstIterator
        {
        public:
            using iterator_category = std::forward_iterator_tag;
            using value_type        = XLSlicer;
            using difference_type   = std::ptrdiff_t;
            using pointer           = const XLSlicer*;
            using reference         = const XLSlicer&;

            explicit ConstIterator(const std::vector<XLSlicer>* vec, size_t idx)
                : m_vec(vec), m_idx(idx) {}

            const XLSlicer& operator*()  const { return (*m_vec)[m_idx]; }
            const XLSlicer* operator->() const { return &(*m_vec)[m_idx]; }
            ConstIterator& operator++() { ++m_idx; return *this; }
            bool operator==(const ConstIterator& o) const { return m_idx == o.m_idx; }
            bool operator!=(const ConstIterator& o) const { return m_idx != o.m_idx; }
        private:
            const std::vector<XLSlicer>* m_vec;
            size_t                       m_idx;
        };

        // ── Constructors ────────────────────────────────────────────────────
        XLSlicerCollection() = default;
        explicit XLSlicerCollection(XLWorksheet* worksheet);

        // ── Query ────────────────────────────────────────────────────────────
        [[nodiscard]] size_t count() const;
        [[nodiscard]] bool   empty() const;
        [[nodiscard]] bool   contains(std::string_view name) const;

        /// Returns a valid XLSlicer if found, or an invalid one (!valid()) if not.
        [[nodiscard]] XLSlicer  find(std::string_view name) const;
        [[nodiscard]] XLSlicer& operator[](size_t index);
        [[nodiscard]] XLSlicer& operator[](std::string_view name);
        [[nodiscard]] const XLSlicer& operator[](size_t index) const;

        // ── Iteration ────────────────────────────────────────────────────────
        Iterator      begin();
        Iterator      end();
        ConstIterator begin() const;
        ConstIterator end()   const;

        // ── STL compatibility (replaces old std::vector<XLSlicer>) ───────────
        [[nodiscard]] size_t size() const { return count(); }
        bool                 valid() const { return m_worksheet != nullptr; }

        // ── Add (new ergonomic API) ──────────────────────────────────────────
        /// Add a slicer filtering a table column.
        [[nodiscard]] XLSlicerBuilder add(std::string_view       cellRef,
                                          const XLTable&         table,
                                          std::string_view       columnName);
        /// Add a slicer filtering a pivot table field.
        [[nodiscard]] XLSlicerBuilder add(std::string_view       cellRef,
                                          const XLPivotTable&    pivotTable,
                                          std::string_view       fieldName);

        // ── Remove ────────────────────────────────────────────────────────────
        void remove(std::string_view name);
        void remove(size_t index);

        // ── Internal: called by XLSlicerBuilder::commit() ───────────────────
        void reload();

    private:
        XLWorksheet*         m_worksheet{nullptr};
        mutable std::vector<XLSlicer> m_slicers;
        mutable bool         m_loaded{false};

        void load() const;
    };

} // namespace OpenXLSX

#endif // OPENXLSX_XLSLICERCOLLECTION_HPP

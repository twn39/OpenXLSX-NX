#ifndef OPENXLSX_XLSLICER_HPP
#define OPENXLSX_XLSLICER_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include "XLXmlParser.hpp"
#include <string>
#include <string_view>
#include <vector>
#include <cstdint>

namespace OpenXLSX
{
    class XLWorksheet;
    class XLDocument;

    // ─────────────────────────────────────────────────────────────────────────
    // XLSlicerStyle — strongly-typed enum replacing magic strings
    // ─────────────────────────────────────────────────────────────────────────
    enum class XLSlicerStyle : uint8_t {
        // Light styles
        Light1 = 0, Light2, Light3, Light4, Light5, Light6,
        // Dark styles
        Dark1, Dark2, Dark3, Dark4, Dark5, Dark6,
        // Other styles
        Other1, Other2,
        // Custom (use setStyleRaw for arbitrary names)
        Custom = 255
    };

    /// Convert XLSlicerStyle enum to its Excel XML string representation.
    OPENXLSX_EXPORT std::string xlSlicerStyleToString(XLSlicerStyle style);
    /// Convert a raw style string to XLSlicerStyle enum (Custom if unrecognized).
    OPENXLSX_EXPORT XLSlicerStyle xlSlicerStyleFromString(std::string_view s);

    // ─────────────────────────────────────────────────────────────────────────
    // XLSlicer — three-source unified handle
    //
    // A slicer's full state lives across three XML parts:
    //   1. xl/slicers/slicerN.xml     → name, caption, style, showCaption …
    //   2. xl/slicerCaches/slicerCacheN.xml → selectedItems, sortOrder …
    //   3. xl/drawings/drawingM.xml  → position anchor (cellRef, width, height)
    //
    // This class holds references to all three so every property is readable
    // and writable from a single handle. All setters return *this for chaining.
    // ─────────────────────────────────────────────────────────────────────────
    class OPENXLSX_EXPORT XLSlicer
    {
    public:
        // ── Constructors ────────────────────────────────────────────────────
        XLSlicer() = default;

        /// Minimal constructor: wraps only the slicer XML (legacy compat).
        explicit XLSlicer(XLXmlData* slicerXml);

        /// Full three-source constructor used by XLSlicerCollection.
        XLSlicer(XLXmlData* slicerXml,
                 XMLNode    slicerNode,
                 XLXmlData* cacheXml,
                 XMLNode    anchorNode,
                 XLWorksheet* worksheet);

        XLSlicer(const XLSlicer&)            = default;
        XLSlicer(XLSlicer&&) noexcept        = default;
        XLSlicer& operator=(const XLSlicer&) = default;
        XLSlicer& operator=(XLSlicer&&)      = default;
        ~XLSlicer()                          = default;

        // ── Validity ────────────────────────────────────────────────────────
        [[nodiscard]] bool valid() const { return m_slicerXml != nullptr; }
        explicit operator bool() const   { return valid(); }

        // ── Source 1: slicer XML properties ─────────────────────────────────

        [[nodiscard]] std::string   name()         const;
        [[nodiscard]] std::string   caption()      const;
        [[nodiscard]] std::string   cache()        const;  ///< Cache name (e.g. "Slicer_Region")
        [[nodiscard]] XLSlicerStyle style()        const;
        [[nodiscard]] std::string   styleRaw()     const;  ///< Raw style string
        [[nodiscard]] bool          showCaption()  const;
        [[nodiscard]] int           columnCount()  const;  ///< Button columns (0 = unset → 1)
        [[nodiscard]] bool          lockedPosition() const;
        [[nodiscard]] int           rowHeight()    const;  ///< EMU per item row (0 = unset)

        void setName(std::string_view name);
        XLSlicer& setCaption(std::string_view caption);
        XLSlicer& setStyle(XLSlicerStyle style);
        XLSlicer& setStyleRaw(std::string_view rawStyleName); ///< Arbitrary style string
        XLSlicer& setShowCaption(bool show);
        XLSlicer& setColumnCount(int cols);        ///< Multi-column slicer layout
        XLSlicer& setLockedPosition(bool locked);
        XLSlicer& setRowHeight(int emuHeight);     ///< Raw EMU, 251883 ≈ 20 px

        // ── Source 2: cache XML — filter items ──────────────────────────────

        /// All available items in the slicer (read from pivot/table data).
        [[nodiscard]] std::vector<std::string> items()         const;
        /// Items that are currently SELECTED (shown, not filtered out).
        [[nodiscard]] std::vector<std::string> selectedItems() const;
        /// True if items are sorted descending (Z→A).
        [[nodiscard]] bool isSortDescending() const;

        /// Keep ONLY the listed items visible; hide all others.
        XLSlicer& showOnly(const std::vector<std::string>& itemsToShow);
        /// Clear all filters — show every item.
        XLSlicer& showAll();
        /// Hide the listed items (keep everything else visible).
        XLSlicer& hideItems(const std::vector<std::string>& itemsToHide);
        /// Set sort direction.
        XLSlicer& setSortDescending(bool desc);

        // ── Source 3: drawing anchor — position & size ───────────────────────

        /// Top-left anchor cell reference, e.g. "E2". Empty if anchor unknown.
        [[nodiscard]] std::string cellRef() const;
        /// Width in pixels (converted from EMU, 0 if anchor unknown).
        [[nodiscard]] uint32_t    width()   const;
        /// Height in pixels (converted from EMU, 0 if anchor unknown).
        [[nodiscard]] uint32_t    height()  const;

        /// Move the slicer anchor to a new cell (updates drawing XML).
        XLSlicer& moveTo(std::string_view cellRef);
        /// Resize the slicer (updates drawing XML, pixels).
        XLSlicer& resize(uint32_t widthPx, uint32_t heightPx);

        // ── Internals used by XLSlicerCollection ────────────────────────────
        XLXmlData*   slicerXml()  const { return m_slicerXml; }
        XLXmlData*   cacheXml()   const { return m_cacheXml;  }

    private:
        /// Returns the root <slicer> child node of m_slicerXml.
        XMLNode slicerNode() const;
        /// Returns the root element of m_cacheXml.
        XMLNode cacheRoot()  const;

        XLXmlData*   m_slicerXml{nullptr};
        XMLNode      m_slicerNode;
        XLXmlData*   m_cacheXml{nullptr};
        XMLNode      m_anchorNode;       ///< xdr:oneCellAnchor or xdr:twoCellAnchor
        XLWorksheet* m_worksheet{nullptr};
    };

} // namespace OpenXLSX

#endif // OPENXLSX_XLSLICER_HPP

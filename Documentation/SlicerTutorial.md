# Slicer Tutorial {#slicer_tutorial}

[TOC]

## Introduction

A **Slicer** is an interactive visual filtering control that lets end-users click buttons to filter the data displayed in an Excel Table or Pivot Table. OpenXLSX provides a complete, OOXML-compliant Slicer API that generates files compatible with Microsoft Excel 2013+.

This tutorial covers:
- Adding Table Slicers to filter Excel Tables
- Adding Pivot Slicers to filter Pivot Tables
- Multiple slicers on a single or multiple worksheets
- Styling, sizing, and positioning
- Programmatic pre-filtering (showOnly / hideItems)
- CRUD: reading, modifying, and deleting slicers

---

## 1. Concepts: Two Slicer Types

OpenXLSX supports two distinct slicer types that differ in their OOXML representation:

| Type | Connected to | XML format | Key OOXML URI |
|---|---|---|---|
| **Table Slicer** | `XLTable` (Excel Table) | `mc:Choice Requires="sle15"` | `{46BE6895…}` in workbook |
| **Pivot Slicer** | `XLPivotTable` | `mc:Choice Requires="sle"` | `{BBE1A952…}` in workbook |

Both types are created via `XLWorksheet::addTableSlicer()` and `XLWorksheet::addPivotSlicer()` respectively.

> **Note:** When a workbook contains **both** Pivot Slicers and Table Slicers, the internal `workbook.xml` registration order is critical. OpenXLSX automatically handles this: Pivot Slicer caches are always registered before Table Slicer caches, matching Excel's own output format.

---

## 2. Table Slicer (Filter an Excel Table)

### 2.1 Minimal Example

```cpp
#include <OpenXLSX.hpp>
using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("TableSlicerDemo.xlsx", XLForceOverwrite);

    auto wks = doc.workbook().worksheet("Sheet1");

    // 1. Fill data
    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Product";
    wks.cell("C1").value() = "Sales";
    wks.cell("A2").value() = "North";  wks.cell("B2").value() = "Widget";  wks.cell("C2").value() = 1200;
    wks.cell("A3").value() = "South";  wks.cell("B3").value() = "Gadget";  wks.cell("C3").value() = 850;
    wks.cell("A4").value() = "North";  wks.cell("B4").value() = "Widget";  wks.cell("C4").value() = 960;
    wks.cell("A5").value() = "East";   wks.cell("B5").value() = "Gadget";  wks.cell("C5").value() = 1100;

    // 2. Create an Excel Table over the data range
    auto& tables = wks.tables();
    auto table = tables.add("T_Sales", "A1:C5");

    // 3. Add a Table Slicer connected to the "Region" column
    //    Anchored at cell F1, default size (150 × 200 px)
    wks.addTableSlicer("F1", table, "Region");

    doc.save();
    doc.close();
    return 0;
}
```

### 2.2 With Full Options

```cpp
    XLSlicerOptions opts;
    opts.name      = "Slicer_Region";     // Internal XML name (must be unique per workbook)
    opts.caption   = "Filter by Region";  // Displayed title bar text
    opts.style     = XLSlicerStyle::Dark3;
    opts.width     = 180;                 // pixels
    opts.height    = 260;                 // pixels
    opts.columnCount = 2;                 // Show items in 2 columns
    opts.rowHeight   = 251883;            // EMU per item row (251883 ≈ 20px, default)
    opts.showCaption = true;

    wks.addTableSlicer("F1", table, "Region", opts);
```

### 2.3 Multiple Table Slicers on One Sheet

You can add multiple slicers to the same worksheet. Each slicer must have a **unique name** within the workbook.

```cpp
    // Slicer 1: Filter by Region
    XLSlicerOptions opts1;
    opts1.name    = "Slicer_Region";
    opts1.caption = "Region";
    opts1.style   = XLSlicerStyle::Light1;
    wks.addTableSlicer("F1", table, "Region", opts1);

    // Slicer 2: Filter by Product (placed to the right)
    XLSlicerOptions opts2;
    opts2.name    = "Slicer_Product";
    opts2.caption = "Product";
    opts2.style   = XLSlicerStyle::Dark3;
    wks.addTableSlicer("H1", table, "Product", opts2);
```

---

## 3. Pivot Slicer (Filter a Pivot Table)

A Pivot Slicer connects to an `XLPivotTable` and lets users interactively filter pivot data.

```cpp
    // --- Sheet1: Raw data ---
    auto rawWks = doc.workbook().worksheet("Sheet1");
    // ... [fill data] ...

    // --- Sheet2: Pivot Table ---
    doc.workbook().addWorksheet("Summary");
    auto summaryWks = doc.workbook().worksheet("Summary");

    auto ptOptions = XLPivotTableOptions("PT_Sales", "Sheet1!A1:C5", "B2")
        .addRowField("Region")
        .addDataField("Sales", "Total Sales", XLPivotSubtotal::Sum);
    auto pt = summaryWks.addPivotTable(ptOptions);

    // Add a Pivot Slicer filtering by "Region"
    XLSlicerOptions sOpts;
    sOpts.name    = "Region";       // Must match the pivot field source column name
    sOpts.caption = "地区筛选";
    sOpts.style   = XLSlicerStyle::Dark4;
    summaryWks.addPivotSlicer("F2", pt, "Region", sOpts);

    doc.save();
    doc.close();
```

> **Important:** The `name` in `XLSlicerOptions` for a Pivot Slicer refers to the **source column name** in the pivot data (i.e., the field name from the original data range). For Table Slicers, it is the slicer's own identity name.

---

## 4. Mixed: Table + Pivot Slicers in One Workbook

OpenXLSX fully supports workbooks that contain both Table Slicers and Pivot Slicers, even across multiple worksheets.

```cpp
    // Sheet1: Table with Table Slicers
    auto wks1 = doc.workbook().worksheet("Sheet1");
    // ... [fill data and create table] ...
    wks1.addTableSlicer("F1", tableOnSheet1, "Region", optsRegion);
    wks1.addTableSlicer("H1", tableOnSheet1, "Product", optsProduct);

    // Sheet2: Pivot Table with Pivot Slicer
    auto wks2 = doc.workbook().worksheet("Sheet2");
    // ... [create pivot table] ...
    wks2.addPivotSlicer("F2", pivotTable, "Region", optsPivotRegion);

    // Sheet3: Another Table with Table Slicer
    auto wks3 = doc.workbook().worksheet("Sheet3");
    // ... [fill data and create table] ...
    wks3.addTableSlicer("F1", tableOnSheet3, "Quarter", optsQuarter);

    doc.save();
    doc.close();
```

---

## 5. Slicer Styles

OpenXLSX provides a strongly-typed `XLSlicerStyle` enum matching all built-in Excel slicer styles.

```cpp
// Available styles:
XLSlicerStyle::Light1  // SlicerStyleLight1
XLSlicerStyle::Light2
XLSlicerStyle::Light3
XLSlicerStyle::Light4
XLSlicerStyle::Light5
XLSlicerStyle::Light6

XLSlicerStyle::Dark1   // SlicerStyleDark1
XLSlicerStyle::Dark2
XLSlicerStyle::Dark3
XLSlicerStyle::Dark4
XLSlicerStyle::Dark5
XLSlicerStyle::Dark6

XLSlicerStyle::Other1  // SlicerStyleOther1
XLSlicerStyle::Other2

XLSlicerStyle::Custom  // Use setStyleRaw() for arbitrary names
```

To apply a style at creation time:

```cpp
    XLSlicerOptions opts;
    opts.style = XLSlicerStyle::Dark4;
    wks.addTableSlicer("E1", table, "Region", opts);
```

To change the style after creation via the Slicer Collection API:

```cpp
    wks.slicers()["Slicer_Region"].setStyle(XLSlicerStyle::Light2);
```

---

## 6. Slicer Collection API (CRUD)

`XLWorksheet::slicers()` returns an `XLSlicerCollection` that provides full read/write access to all slicers on the sheet.

### 6.1 Iterate over all slicers

```cpp
    for (auto& slicer : wks.slicers()) {
        std::cout << "Slicer: " << slicer.name()
                  << ", caption: " << slicer.caption()
                  << ", style: " << slicer.styleRaw()
                  << ", at: " << slicer.cellRef()
                  << " (" << slicer.width() << "x" << slicer.height() << " px)\n";
    }
```

### 6.2 Access by name

```cpp
    auto& s = wks.slicers()["Slicer_Region"];
    if (s.valid()) {
        s.setCaption("Filter: Region")
         .setStyle(XLSlicerStyle::Dark1)
         .setColumnCount(3);
    }
```

### 6.3 Modify position and size

```cpp
    wks.slicers()["Slicer_Product"]
        .moveTo("J1")          // Move anchor to J1
        .resize(200, 300);     // Width=200px, Height=300px
```

### 6.4 Pre-filter items programmatically

```cpp
    // Show ONLY "North" and "East" (hide "South", "West", etc.)
    wks.slicers()["Slicer_Region"].showOnly({"North", "East"});

    // Hide specific items
    wks.slicers()["Slicer_Region"].hideItems({"South"});

    // Clear all filters (show everything)
    wks.slicers()["Slicer_Region"].showAll();
```

### 6.5 Delete a slicer

```cpp
    wks.deleteSlicer("Slicer_Product");
    // Or via collection:
    // wks.slicers().remove("Slicer_Product");
```

Deletion removes the slicer from the drawing XML, the slicer XML part, the slicer cache XML, all workbook relationships, and all worksheet relationships. Orphaned files are automatically removed from the package.

---

## 7. XLSlicerOptions Reference

```cpp
struct XLSlicerOptions {
    std::string    name;           // Internal slicer name (unique per workbook).
                                   // Default: auto-generated from column name ("Slicer_<col>")
    std::string    caption;        // Display title shown on slicer header.
                                   // Default: same as column name
    XLSlicerStyle  style{XLSlicerStyle::Light1}; // Visual style
    uint32_t       width{150};     // Slicer width in pixels
    uint32_t       height{200};    // Slicer height in pixels
    uint32_t       offsetX{0};     // X offset from anchor cell in pixels
    uint32_t       offsetY{0};     // Y offset from anchor cell in pixels
    int            columnCount{1}; // Number of item columns in the slicer panel
    int            rowHeight{251883}; // EMU per item row (251883 ≈ 20px)
    bool           showCaption{true}; // Show/hide the slicer title bar
    bool           lockedPosition{false}; // Prevent user from moving slicer
};
```

---

## 8. XLSlicer Property Reference

After creation, every property of a slicer is readable and writable via the `XLSlicer` handle.

### Getters

| Method | Return type | Description |
|---|---|---|
| `name()` | `std::string` | Internal slicer name |
| `caption()` | `std::string` | Displayed header text |
| `cache()` | `std::string` | Name of the linked slicer cache |
| `style()` | `XLSlicerStyle` | Enum style value |
| `styleRaw()` | `std::string` | Raw XML style string |
| `showCaption()` | `bool` | Whether header is shown |
| `columnCount()` | `int` | Number of item columns |
| `rowHeight()` | `int` | EMU height per item row |
| `lockedPosition()` | `bool` | Whether position is locked |
| `cellRef()` | `std::string` | Anchor cell, e.g. `"E2"` |
| `width()` | `uint32_t` | Width in pixels |
| `height()` | `uint32_t` | Height in pixels |
| `items()` | `vector<string>` | All available filter items |
| `selectedItems()` | `vector<string>` | Currently selected (visible) items |
| `isSortDescending()` | `bool` | Sort direction |

### Setters (fluent, return `XLSlicer&`)

| Method | Description |
|---|---|
| `setCaption(caption)` | Change displayed title |
| `setStyle(style)` | Change style enum |
| `setStyleRaw(name)` | Set arbitrary style name |
| `setShowCaption(bool)` | Show/hide title bar |
| `setColumnCount(int)` | Multi-column layout |
| `setRowHeight(emu)` | Item row height in EMU |
| `setLockedPosition(bool)` | Lock slicer position |
| `moveTo(cellRef)` | Reposition anchor cell |
| `resize(w, h)` | Resize in pixels |
| `showOnly(items)` | Pre-filter: keep only listed items |
| `hideItems(items)` | Pre-filter: hide listed items |
| `showAll()` | Clear all filters |
| `setSortDescending(bool)` | Set sort direction |

---

## 9. OOXML Compatibility Notes

The following constraints are enforced internally by OpenXLSX to ensure 100% compatibility with Microsoft Excel 2013+:

1. **Workbook extLst order**: The `{BBE1A952}` Pivot slicer cache block must appear **before** the `{79F54976}` workbookPr block, which must appear **before** the `{46BE6895}` Table slicer caches block in `workbook.xml`. OpenXLSX handles this automatically.

2. **Drawing XML namespace scoping**: Table Slicer drawings use `xmlns:sle15` on `mc:Choice` and `xmlns:sle` directly on `sle:slicer`. Pivot Slicer drawings use `xmlns:sle` on `mc:AlternateContent`. The `mc:Fallback` element must **not** carry `xmlns=""` for Table Slicers.

3. **Shape IDs are drawing-local**: `cNvPr id` attributes must be unique within a single `drawingN.xml` file, not across the entire workbook.

4. **UUID format**: All `xr10:uid` attributes must use the full 8-4-4-4-**12** hex UUID format (e.g., `{00000000-0009-0000-0100-000001000000}`). The last segment is always 12 characters.

5. **Boolean attributes in slicer caches**: Use `"0"` and `"1"`, not `"false"` and `"true"`, for OOXML boolean attributes such as `showMissing` and `s` in `slicerCacheDefinition`.

6. **Slicer name uniqueness**: Each slicer name must be unique across the entire workbook (not just per worksheet). OpenXLSX automatically appends numeric suffixes (e.g., `Slicer_Region1`) when a name conflict is detected.

---

## 10. Complete Example: Mixed Dashboard

```cpp
#include <OpenXLSX.hpp>
using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("SlicerDashboard.xlsx", XLForceOverwrite);

    // ── Sheet1: Sales Data + Table Slicers ──────────────────────────────────
    auto wks1 = doc.workbook().worksheet("Sheet1");
    wks1.setName("Sales");

    // Fill data
    wks1.cell("A1").value() = "Region";
    wks1.cell("B1").value() = "Product";
    wks1.cell("C1").value() = "Quarter";
    wks1.cell("D1").value() = "Sales";
    // ... [add data rows A2:D9] ...

    // Create table
    auto& tables = wks1.tables();
    auto table = tables.add("T_Sales", "A1:D9");

    // Add two Table Slicers side by side
    wks1.addTableSlicer("F1", table, "Region",
        XLSlicerOptions{"Slicer_Region", "Region", XLSlicerStyle::Light1, 150, 220});
    wks1.addTableSlicer("H1", table, "Product",
        XLSlicerOptions{"Slicer_Product", "Product", XLSlicerStyle::Dark3, 150, 220});

    // ── Sheet2: Pivot Table + Pivot Slicer ──────────────────────────────────
    doc.workbook().addWorksheet("PivotView");
    auto wks2 = doc.workbook().worksheet("PivotView");

    auto ptOpts = XLPivotTableOptions("PT_Sales", "Sales!A1:D9", "B2")
        .addRowField("Region")
        .addDataField("Sales", "Total", XLPivotSubtotal::Sum)
        .setRefreshOnLoad(true);
    auto pt = wks2.addPivotTable(ptOpts);

    XLSlicerOptions psOpts;
    psOpts.name    = "Region";
    psOpts.caption = "Filter Region";
    psOpts.style   = XLSlicerStyle::Dark4;
    psOpts.width   = 160;
    psOpts.height  = 240;
    wks2.addPivotSlicer("G2", pt, "Region", psOpts);

    // ── Save ─────────────────────────────────────────────────────────────────
    doc.save();
    doc.close();
    return 0;
}
```

Open the generated `SlicerDashboard.xlsx` in Excel — both sheets will display interactive slicers with no file-corruption warnings.

# OpenXLSX XLStreamWriter Tutorial & Best Practices

The `XLStreamWriter` is a highly optimized, low-allocation module designed for generating massive Excel worksheets (millions of cells) with a small memory footprint. It bypasses the standard `pugixml` DOM engine and writes XML data directly to a temporary file stream.

## Core Concepts & Limitations

Because `XLStreamWriter` streams data directly to disk:

1. **No In-Memory DOM**: Cells written via `XLStreamWriter` **do not exist** in the `XLDocument`'s memory. You cannot read them back using `wks.cell(...)` during the same session.
2. **Strictly Increasing Rows**: Write rows from top to bottom. Use `appendRow` for sequential rows, or `setRow` with a higher row number (excelize-compatible order). You cannot go back to a previous row.
3. **Exclusive Lock**: While a stream writer is active, do not modify the same worksheet with DOM methods (`wks.cell(...)`, `wks.row(...)`). Only one active stream writer is allowed per worksheet.
4. **Dimension**: On `close()` / destruction, `<dimension>` in the temporary worksheet XML is patched to the used range (DOM base + streamed extents).

## Lifecycle: `close()` Before `save()`

The writer keeps the XML stream open and buffers data. Call **`close()`** (or `flush()`, an alias) or destroy the writer so footers (`</sheetData>`, …) are written and `<dimension>` is patched.

**`doc.save()` while a stream writer is still open throws `XLInputError`** and does **not** write a corrupt package.

### ❌ Incorrect Usage (Throws)

```cpp
XLDocument doc;
auto wks = doc.workbook().worksheet("Sheet1");
auto stream = wks.streamWriter();

std::vector<XLCellValue> row = {1, 2, 3};
stream.appendRow(row);

doc.save(); // throws XLInputError — stream still active
```

### ✅ Correct Usage (Scope or explicit close)

```cpp
XLDocument doc;
auto wks = doc.workbook().worksheet("Sheet1");

{ // Scope so destructor closes the stream
    auto stream = wks.streamWriter();
    std::vector<XLCellValue> row = {1, 2, 3};
    stream.appendRow(row);
} // close + dimension patch

doc.save();
```

```cpp
auto stream = wks.streamWriter();
stream.appendRow({1, 2, 3});
stream.close(); // or stream.flush()
doc.save();
```

## Formulas, Row Options, and `setRow`

### Formulas (`XLStreamCell`)

```cpp
auto stream = wks.streamWriter();
stream.appendRow({
    XLStreamCell(1),
    XLStreamCell(2),
    XLStreamCell::withFormula("A1+B1", XLCellValue(3)), // formula + cached value
});
stream.appendRow({
    XLStreamCell::withFormula("SUM(A1:B1)"), // formula only
});
stream.close();
```

### Row options (`XLStreamRowOpts`)

```cpp
XLStreamRowOpts opts;
opts.height = 24.0;
opts.hidden = false;
// opts.outlineLevel = 1;
// opts.styleIndex = someStyle;

stream.appendRow({XLCellValue("Header")}, opts);
```

### Explicit coordinates (`setRow`)

```cpp
// Start at C10 (row 10, column 3); rows must increase vs previous writes
stream.setRow("C10", {XLCellValue(1), XLCellValue(2)});
stream.setRow(12, 1, {XLCellValue("next")}); // row 12, col A
// stream.setRow(11, ...) → throws (row must be > last written row)
```

## Streaming Data for Pivot Tables

When creating an `XLPivotTable` from a data range, `addPivotTable()` needs the **header row** in the DOM to name pivot cache fields.

If the header is only written via the stream writer, headers may appear as `"Field1"`, `"Field2"`, ….

### Hybrid DOM + Stream

```cpp
auto wksData = doc.workbook().worksheet("RawData");

// 1. Headers in DOM
wksData.cell("A1").value() = "Region";
wksData.cell("B1").value() = "Sales";

// 2. Bulk rows via stream
{
    auto stream = wksData.streamWriter();
    for (int i = 0; i < 1000000; ++i) {
        stream.appendRow({XLCellValue("North"), XLCellValue(500.0)});
    }
}

// 3. Pivot
auto pivotWks = doc.workbook().worksheet("PivotDashboard");
XLPivotTableOptions options("MyPivot", "RawData!A1:B1000001", "A3");
options.addRowField("Region").addDataField("Sales");
auto pt = pivotWks.addPivotTable(options);
pt.setRefreshOnLoad(true);

doc.save();
```

## Shared Strings

```cpp
// Default: inlineStr (lower process memory, larger files for repeated strings)
auto stream = wks.streamWriter();

// Optional SST with unique-string cap (fallback to inlineStr when exceeded)
auto streamSst = wks.streamWriter(/*useSharedStrings=*/true, /*maxUniqueStrings=*/100000);

// Recommended: also set a hard SST byte budget (default 64 MiB) so unique long
// strings cannot grow memory unboundedly (excelize StreamWriter OOM class of bugs).
XLStreamWriteOptions opts;
opts.useSharedStrings = true;
opts.maxUniqueStrings = 100000;
opts.maxSstBytes      = 64 * 1024 * 1024; // 0 = unlimited
auto streamBudget     = wks.streamWriter(opts);
```

## Freeze panes (structured)

```cpp
auto stream = wks.streamWriter();
stream.freezePanes("B2");              // freeze row 1 + column A
// or:
XLStreamPanesOptions panes;
panes.freezeCols = 1;
panes.freezeRows = 1;
stream.setPanes(panes);
stream.appendRow({XLCellValue("H1"), XLCellValue("H2")});
```

## Date / time cells

```cpp
stream.appendRow({
    XLStreamCell::withDateTime(XLDateTime::fromString("2024-06-15 12:30:00"))
});
// Applies a cached default number format (yyyy-mm-dd hh:mm:ss) unless a style is passed.
```

## Hyperlinks, AutoFilter, shared/array formulas (P2)

```cpp
stream.appendRow({XLCellValue("Name"), XLCellValue("Go")});
stream.addHyperlink("B1", "https://example.com", "tooltip");
stream.addInternalHyperlink("A2", "Sheet1!A1");
stream.setAutoFilter("A1:B10");

stream.appendRow({
    XLStreamCell::withArrayFormula("SUM(A1:A2)", "A3"),
    XLStreamCell::withSharedFormula(/*si=*/0, "B1+1", "B1:B2"),
});
stream.appendRow({
    XLCellValue(1),
    XLStreamCell::withSharedFormula(0), // dependent shared cell
});
```

## Close path (memory)

On `close()` / destruction the writer:

1. Writes the sheet footer (`</sheetData>`, merges, page breaks, table parts).
2. **Patches `<dimension>` in-place** using a fixed-width placeholder (`A1:XFD1048576` slot) — **no full-file reload** of the worksheet XML.
3. **Package-only materializes** to a temp file or memory buffer used by `doc.save()`; streamed sheets skip DOM `updateDimension()`.

```cpp
XLStreamWriteOptions opts;
opts.memorySpillThreshold = 0; // always spill rows to temp file
opts.dimensionMode        = XLStreamDimensionMode::FixedSlot; // default
// opts.dimensionMode = XLStreamDimensionMode::OmitPatch; // leave max-range placeholder
```

## Supported Types

`XLStreamWriter` uses `std::to_chars` / stack buffers for numeric types and supports Integers, Floats, Booleans, Strings, `XLRichText`, Errors, per-cell styles (`XLStreamCell` + `XLStyleIndex`), and formulas as above.

## Related: streaming read

See [StreamReaderTutorial.md](StreamReaderTutorial.md) for `XLStreamReader` (formulas, styles, empty-row policy, number formats).

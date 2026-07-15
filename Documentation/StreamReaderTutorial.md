# OpenXLSX XLStreamReader Tutorial

`XLStreamReader` streams worksheet XML from the ZIP package with a SAX-style parser. It does **not** build a per-row DOM tree, so large sheets can be scanned with bounded memory.

## Basic usage

```cpp
XLDocument doc;
doc.open("large.xlsx");
auto wks    = doc.workbook().worksheet("Sheet1");
auto reader = wks.streamReader();

while (reader.hasNext()) {
    auto row = reader.nextRow();           // vector<XLCellValue>
    uint32_t r = reader.currentRow();      // 1-based
    // ...
}
reader.close();   // optional if fully consumed (EOF closes the zip entry)
doc.close();
```

**Lifetime:** Keep the `XLWorksheet` (and document) alive for the lifetime of the reader. Prefer `reader.close()` before `doc.close()` if the stream was not fully consumed.

## Options (`XLStreamReadOptions`)

```cpp
XLStreamReadOptions opts;
opts.emptyRows = XLStreamEmptyRowPolicy::EmitEmptyRows; // or SkipMissingRows (default)
opts.applyNumberFormats = true;                         // for nextRowStrings / formattedValue
opts.sharedStringMode = XLStreamSharedStringMode::Indexed; // optional large-SST path
opts.sstSpillThreshold = 16 * 1024 * 1024;              // spill Indexed SST payload to temp file
opts.maxRowBytes = 64 * 1024 * 1024;                    // guard against pathological wide rows

auto reader = wks.streamReader(opts);
```

### Empty-row policy

| Policy | Behaviour when XML has row 1 then row 3 |
|---|---|
| `SkipMissingRows` (default) | Second `nextRow()` returns row 3 |
| `EmitEmptyRows` | Second call returns empty row 2 (`currentRowOpts().isSyntheticEmpty == true`), then row 3 |

## Detailed cells (formula + style)

```cpp
while (reader.hasNext()) {
    auto cells = reader.nextRowDetailed();  // vector<XLStreamCellView>
    for (const auto& c : cells) {
        // c.value, c.formula, c.styleIndex, c.column (1-based)
    }
    auto rowOpts = reader.currentRowOpts(); // height / hidden / outlineLevel / styleIndex
}
```

## Display strings (excelize-like)

```cpp
XLStreamReadOptions opts;
opts.applyNumberFormats = true;
auto reader = wks.streamReader(opts);

auto strings = reader.nextRowStrings(); // formatted when style has a number format
// or:
auto cells = reader.nextRowDetailed();
auto s = reader.formattedValue(cells[0]);
```

When `applyNumberFormats` is false, strings are raw conversions of the stored value (similar to excelize `RawCellValue`).

## Column projection (wide tables)

```cpp
XLStreamReadOptions opts;
opts.columnFilter = {1, 3, 5};    // 1-based columns A, C, E
auto reader = wks.streamReader(opts);

while (reader.hasNext()) {
    auto cells = reader.nextRowProjected();   // only filtered columns
    // or:
    // reader.forEachCellInNextRow([](const XLStreamCellView& c) { ... });
}
```

Keep the `XLWorksheet` alive for the lifetime of the reader.

## Rich text, formula metadata, columns (P2)

```cpp
XLStreamReadOptions opts;
opts.parseRichText = true;
auto reader = wks.streamReader(opts);
auto row = reader.nextRowDetailed();
// row[i].formulaType / formulaRef / formulaSi / formulaCa
// row[i].richText  when inlineStr had <r> runs

// Column-oriented pass (one full scan into columnar buffers; excelize Cols-like)
auto cols = wks.streamReader().streamColumns();
while (cols.next()) {
    uint16_t c = cols.currentColumn(); // 1-based
    const auto& vals = cols.values();  // top-to-bottom
}
```

## Compatibility notes

- `nextRow()` remains the stable, type-preserving API (legacy behaviour for empty rows).
- Shared strings (`t="s"`), booleans, numbers, inlineStr, errors, and formulas (`<f>`) are supported.
- Number formatting uses workbook styles + built-in ECMA numFmtId table when custom formats are absent.
- Reader is not concurrent with mutations of the same worksheet package entry.

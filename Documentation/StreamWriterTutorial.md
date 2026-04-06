# Streaming Large Files: XLStreamWriter {#stream_writer_tutorial}

[TOC]

## Introduction

When generating very large Excel files (e.g., hundreds of thousands of rows), writing data cell-by-cell can consume a significant amount of memory. To solve this, OpenXLSX provides the `XLStreamWriter` interface.

The `XLStreamWriter` allows you to append rows to a worksheet sequentially. As rows are appended, the underlying XML is flushed directly to the disk/archive rather than being kept in memory, ensuring that the memory footprint remains small (O(1)) and constant, regardless of the number of rows.

This tutorial covers:
- Basic unstyled data streaming.
- Advanced streaming with rich styling and cell-level formatting.

---

## 1. Basic Data Streaming

For pure data extraction, you can simply stream vectors of `XLCellValue` objects.

```cpp
#include <OpenXLSX.hpp>
#include <iostream>
#include <vector>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("StreamingBasicTest.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Initialize the StreamWriter on the target worksheet
    auto stream = wks.streamWriter();

    // Append a header row
    std::vector<XLCellValue> headerRow = {"ID", "Name", "Score"};
    stream.appendRow(headerRow);

    // Append 100,000 rows of data with minimal memory usage
    for (int i = 1; i <= 100000; ++i) {
        std::vector<XLCellValue> row = {
            i, 
            "Player " + std::to_string(i), 
            i * 3.14
        };
        stream.appendRow(row);
    }

    // Always close the stream before saving the document!
    stream.close();

    doc.save();
    doc.close();
    
    return 0;
}
```

## 2. Advanced Streaming with Styles

If you need to generate professional reports, you can mix streaming with cell-level styling using `XLStreamCell`. `XLStreamCell` pairs an `XLCellValue` with an `XLStyle` ID.

First, pre-define all the styles you will need using `findOrCreateStyle()`.

```cpp
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("StreamingStyledTest.xlsx", XLForceOverwrite);
    
    // 1. Define a Header Style (Blue background, White bold text)
    XLStyle headerStyle;
    headerStyle.font.bold = true;
    headerStyle.font.color = XLColor("FFFFFF");
    headerStyle.fill.pattern = XLPatternSolid;
    headerStyle.fill.fgColor = XLColor("4F81BD");
    auto headerId = doc.styles().findOrCreateStyle(headerStyle);

    // 2. Define standard Currency and Percentage formatting styles
    XLStyle currencyStyle;
    currencyStyle.numberFormat = "$#,##0.00";
    auto currencyId = doc.styles().findOrCreateStyle(currencyStyle);

    XLStyle percentageStyle;
    percentageStyle.numberFormat = "0.00%";
    auto percentId = doc.styles().findOrCreateStyle(percentageStyle);

    // 3. Define an Alternate Row (Zebra-striping) Style
    XLStyle altRowStyle;
    altRowStyle.fill.pattern = XLPatternSolid;
    altRowStyle.fill.fgColor = XLColor("DCE6F1");
    auto altRowId = doc.styles().findOrCreateStyle(altRowStyle);

    auto wks = doc.workbook().worksheet("Sheet1");
    
    // You can set column widths before streaming
    wks.column(1).setWidth(15);
    wks.column(2).setWidth(12);
    wks.column(3).setWidth(12);

    auto stream = wks.streamWriter();

    // Write the styled header row using XLStreamCell
    stream.appendRow({
        XLStreamCell("Product", headerId),
        XLStreamCell("Price", headerId),
        XLStreamCell("Margin", headerId)
    });

    // Write data rows with zebra striping and number formats
    for (int i = 1; i <= 20; ++i) {
        
        // Determine base background color for zebra striping
        uint32_t styleId = (i % 2 == 0) ? altRowId : 0; 
        
        // We must compose the Alternate Row color with the specific Number Formats
        XLStyle rowCurrencyStyle = currencyStyle;
        if (i % 2 == 0) {
            rowCurrencyStyle.fill.pattern = XLPatternSolid;
            rowCurrencyStyle.fill.fgColor = XLColor("DCE6F1");
        }
        auto rowCurrencyId = doc.styles().findOrCreateStyle(rowCurrencyStyle);

        XLStyle rowPercentStyle = percentageStyle;
        if (i % 2 == 0) {
            rowPercentStyle.fill.pattern = XLPatternSolid;
            rowPercentStyle.fill.fgColor = XLColor("DCE6F1");
        }
        auto rowPercentId = doc.styles().findOrCreateStyle(rowPercentStyle);

        // Stream the fully styled row
        stream.appendRow({
            XLStreamCell("Product " + std::to_string(i), styleId),
            XLStreamCell(10.5 * i, rowCurrencyId),
            XLStreamCell(0.15 + (i * 0.01), rowPercentId)
        });
    }

    // Always close the stream before saving the document!
    stream.close();
    doc.save();
    doc.close();

    return 0;
}
```

### Important Streaming Considerations
1. **No going back:** Once `appendRow()` is called, that row is written to the archive. You cannot modify it later or go back to previous rows.
2. **`stream.close()` is mandatory:** You must call `.close()` on the `XLStreamWriter` object before calling `doc.save()`. This flushes and finalizes the XML stream.
3. **Style Combination:** Because styles in Excel are uniquely identified integer IDs representing a combination of attributes, if you want a cell to be both "Currency Formatted" AND have a "Blue Background", you must find/create a style ID that combines both attributes beforehand, as shown in the zebra-striping example above.
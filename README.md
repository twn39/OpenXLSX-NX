# OpenXLSX

[![OpenXLSX CI](https://github.com/twn39/OpenXLSX/actions/workflows/ci.yml/badge.svg)](https://github.com/twn39/OpenXLSX/actions/workflows/ci.yml)

OpenXLSX is a high-performance C++ library for reading, writing, creating, and modifying Microsoft Excel® files in the `.xlsx` format. It is designed to be fast, cross-platform, and has minimal external dependencies.

## 🚀 Key Features

- **Modern C++**: Built with C++17, ensuring type safety and modern abstractions.
- **Formula Engine**: Built-in lightweight AST parser and evaluation engine (`XLFormulaEngine`) supporting 45+ standard Excel functions (Math, Logic, Text, etc.) without requiring external Excel calculation.
- **Advanced Charts**: Fluent Builder API for generating complex 2D/3D charts (`XLChart`), supporting series styling, legends, and strict OOXML schema compliance.
- **Dynamic Structural Mutation**: Insert or delete rows and columns programmatically (`insertRow`, `deleteColumn`) with automatic coordinate shifting.
- **Smart Style Deduplication**: O(1) hash-based global style pool prevents `styles.xml` bloat and file corruption when applying bulk formatting to massive datasets.
- **High Performance**: Optimized for speed, capable of processing millions of cells per second.
- **Stream Reading/Writing**: Support for true streaming workflows (`XLStreamReader` & `XLStreamWriter`) with tiny memory footprint to handle gigabyte-level reports.
- **Zero-Dependency Core**: All dependencies (`libzip`, `pugixml`, `fmt`, `fast_float`) are integrated via CMake's `FetchContent` or standard library.
- **Unicode Support**: Consistent UTF-8 handling across all platforms using C++17 `std::filesystem`.
- **Ergonomic Facade APIs**: "Python/Go-like" C++ developer experience. Use declarative `XLStyle` builder to style cells instantly, auto-fit column widths, or freeze panes intuitively (`freezePanes("B2")`).
- **Comprehensive Image Support**: Insert images via simple APIs (`wks.insertImage("A1", "logo.png")`) with zero-dependency automatic resolution detection and dynamic aspect-ratio scaling (`XLImageOptions`).
- **Batch & Matrix Operations**: Bulk-fill 2D matrices (`std::vector<std::vector<T>>`) into regions using `XLCellRange::setValue()` without `for` loops, and apply bulk styling/border-outlining instantly.
- **Advanced Data Validation**: Fluent-API builder for generating dropdown lists, region subtraction, cross-sheet references, and error alerts (`requireList({"A", "B"})`).
- **Data Tables & AutoFilters**: High-level API for creating and managing Excel Tables (`XLTables`), including totals rows, column-level aggregate functions (`Sum`, `Average`, etc.), and custom logical AutoFilters.
- **AutoFilter & Names**: Set worksheet filters and manage workbook-level named ranges and constants.
- **Page Setup & Print**: Fine-grained control over margins, orientation, paper size, and print options (gridlines, headings, centering).
- **Sheet Protection**: Secure worksheets with passwords and granular permission controls.
- **Enhanced DateTime**: Robust date/time handling with `std::chrono` integration and string parsing/formatting.
- **Worksheet Management**: Create, clone, rename, and delete worksheets with validation.
- **Data Integrity**: Enforces OOXML standards, including strict XML declarations and metadata handling.
- **Modern Testing**: Integrated with **Catch2 v3.13.0** for robust verification.

## 🛠 Quick Start

```cpp
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("Example.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Easy styling and data writing
    wks.cell("A1").value() = "Hello, OpenXLSX!";
    
    XLStyle style;
    style.font.bold = true;
    style.font.color = XLColor("FF0000"); // Red text
    wks.cell("A1").setStyle(style);

    // AutoFit Column
    wks.autoFitColumn(1);

    doc.save();
    return 0;
}
```

### 📊 Advanced Feature: Data Tables & AutoFilters

```cpp
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("SalesTable.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // 1. Prepare data (Headers + 3 rows)
    wks.cell("B2").value() = "Product";
    wks.cell("C2").value() = "Sales";
    wks.cell("B3").value() = "Apples";  wks.cell("C3").value() = 5000;
    wks.cell("B4").value() = "Oranges"; wks.cell("C4").value() = 4200;
    wks.cell("B5").value() = "Bananas"; wks.cell("C5").value() = 7500;

    // 2. Create a Table and snap it to the data
    auto& table = wks.tables();
    table.setName("SalesTable");
    table.resizeToFitData(wks); // Automatically detects B2:C5
    
    // 3. Enable Totals Row and expand range (B2:C6)
    table.setRangeReference("B2:C6"); 
    table.setShowTotalsRow(true);
    table.setStyleName("TableStyleMedium9");

    // 4. Configure Column-level Aggregate Functions
    auto col1 = table.appendColumn("Product");
    col1.setTotalsRowLabel("Total:");
    wks.cell("B6").value() = "Total:";

    auto col2 = table.appendColumn("Sales");
    col2.setTotalsRowFunction(XLTotalsRowFunction::Sum);
    // Use SUBTOTAL formula for dynamic updates when filtering
    wks.cell("C6").formula() = "SUBTOTAL(109,SalesTable[Sales])";

    // 5. Apply a Custom AutoFilter (Sales > 4500)
    auto filter = table.autoFilter();
    filter.filterColumn(1).setCustomFilter("greaterThan", "4500");

    doc.save();
    return 0;
}
```

## 📦 Installation & Build

OpenXLSX uses a simplified CMake (3.15+) build system.

### Integration (FetchContent)
The recommended way to use OpenXLSX is via CMake's `FetchContent`:

```cmake
include(FetchContent)
FetchContent_Declare(
  OpenXLSX
  GIT_REPOSITORY https://github.com/twn39/OpenXLSX.git
  GIT_TAG        master # Or a specific tag
)
FetchContent_MakeAvailable(OpenXLSX)
target_link_libraries(my_project PRIVATE OpenXLSX::OpenXLSX)
```

### Manual Build
To build the library and tests locally:
```bash
mkdir build && cd build
cmake .. -DCMAKE_BUILD_TYPE=Release
cmake --build . -j8
```

## 🧪 Testing
The library features a comprehensive test suite. To run the tests after building:
```bash
# From the build directory
./OpenXLSXTests
```

## ⚠️ Development Conventions

### Unicode / UTF-8
**All string input and output must be in UTF-8 encoding.** OpenXLSX uses `std::filesystem` to handle cross-platform path conversion (including UTF-8 on Windows). Ensure your source files are saved in UTF-8.

### Indexing
- **Sheet Indexing**: 1-based (consistent with Excel).
- **Row/Column Indexing**: Generally 1-based where it follows Excel conventions.

### Performance & Optimizations
The build system includes platform-specific optimizations for `Release` builds (e.g., `/O2` on MSVC, `-O3` on GCC/Clang) and supports **LTO (Link-Time Optimization)** which can be toggled via `OPENXLSX_ENABLE_LTO`.

## 🤝 Credits
- [PugiXML](https://pugixml.org/) - Fast XML parsing.
- [libzip](https://libzip.org/) & [zlib-ng](https://github.com/zlib-ng/zlib-ng) - Fast and compatible ZIP archive handling.
- [fmt](https://github.com/fmtlib/fmt) - Modern formatting library.
- [fast_float](https://github.com/fastfloat/fast_float) - Fast floating-point parsing.

---
## 🚀 Advanced Ergonomic Features

### 1. Matrix Binding & Batch Styling (Like Pandas)
```cpp
auto range = wks.range("A1:C3");

// One-liner to dump a 2D matrix into the spreadsheet
std::vector<std::vector<int>> matrix = { {1,2,3}, {4,5,6}, {7,8,9} };
range.setValue(matrix);

// One-liner to extract a 2D matrix back to C++
auto extracted = range.getValue<int>();

// Draw a thick blue border ONLY around the outer perimeter of the 3x3 matrix
range.setBorderOutline(XLLineStyleThick, XLColor("0000FF"));
```

### 2. Stream Reading (For Multi-Gigabyte Files)
Never worry about `std::bad_alloc` again. `XLStreamReader` uses a micro-DOM sliding window to read massive files with virtually zero memory overhead.
```cpp
auto reader = doc.workbook().worksheet("MassiveData").streamReader();
while (reader.hasNext()) {
    auto row = reader.nextRow();
    std::cout << "Read row " << reader.currentRow() << " with " << row.size() << " columns." << std::endl;
}
```

### 3. Smart Image Insertion
Insert any `png`, `jpg`, or `gif` using natural cell coordinates. The engine automatically parses the binary header to detect the image dimensions without depending on OpenCV or libpng.
```cpp
XLImageOptions opts;
opts.scaleX = 0.5; // Scale to 50%
opts.positioning = XLImagePositioning::TwoCell; // Stretch with cell bounds

wks.insertImage("B2", "company_logo.png", opts);
```

### 4. Fluent Data Validation (Dropdowns)
Build complex dropdown lists and warnings with method chaining.
```cpp
wks.dataValidations().add("C2:C100")
   .requireList({"Pending", "Approved", "Rejected"})
   .setErrorAlert("Invalid", "Please select a valid state from the list.");
```

### 5. UI Behaviors (Freeze Panes & AutoFit)
```cpp
// Instantly freeze Row 1 and Column A
wks.freezePanes("B2"); 

// Make Column B width auto-adapt to fit the longest text inside it
wks.autoFitColumn(2); 

// Hide Column C
wks.column("C").setHidden(true);
```

### 6. Formulas & Merge Cells
```cpp
// Write a formula
wks.cell("E1").formula() = "SUM(A1:D1)";

// Merge multiple cells
wks.mergeCells("A1:D1");
```

### 7. Thread-Safe Concurrent Access
OpenXLSX features a thread-safe two-tier lock architecture that enables high-performance parallel processing, allowing multiple threads to write to **different** worksheets simultaneously without data races.

#### Thread Safety Guarantees
| Operation | Safety | Note |
|-----------|--------|------|
| Multiple threads writing to **different** worksheets | ✅ Safe | Each worksheet has its own DOM |
| `save()` during concurrent writes | ✅ Safe | Automatically blocks until writes finish |
| Structural changes (`addWorksheet`, `deleteSheet`) | ✅ Safe | Protected by document-level exclusive lock |
| Shared string table contention (text insertion) | ✅ Safe | Protected by dedicated SST mutex |
| Multiple threads writing to the **same** worksheet | ❌ Unsafe | Pugixml DOM limitation |

```cpp
// 1. Thread 1 writing to Sheet 1
std::thread t1([&]() {
    auto wks1 = doc.workbook().worksheet("Sheet1");
    wks1.cell("A1").value() = "Written from Thread 1";
});

// 2. Thread 2 writing to Sheet 2 concurrently!
std::thread t2([&]() {
    auto wks2 = doc.workbook().worksheet("Sheet2");
    wks2.cell("A1").value() = "Written from Thread 2";
});

t1.join();
t2.join();

// 3. Save operation safely blocks and waits for all writes to complete
doc.save();
```
*Note: The underlying Shared String Table is automatically protected by a dedicated mutex, making concurrent text insertions entirely thread-safe. Writing to the **same** worksheet from multiple threads is not supported.*

### 8. Built-in Formula Evaluation Engine
Evaluate formulas directly in C++ without needing MS Excel to recalculate the file.
```cpp
wks.cell("A1").value() = 10.5;
wks.cell("A2").value() = 20.2;
wks.cell("B1").formula() = "SUM(A1:A2)";
wks.cell("B2").formula() = "IF(A1>10, "High", "Low")";

XLFormulaEngine engine;
auto resolver = XLFormulaEngine::makeResolver(wks);

// Evaluates to 30.7
double sumResult = engine.evaluate(wks.cell("B1").formula().get(), resolver).get<double>(); 
// Evaluates to "High"
std::string logicResult = engine.evaluate(wks.cell("B2").formula().get(), resolver).getString();
```

### 9. Dynamic Row/Column Insertion
Insert or delete rows and columns on the fly. Existing data and coordinates shift automatically.
```cpp
// Insert 2 blank rows starting at row 5 (existing row 5 becomes row 7)
wks.insertRow(5, 2); 

// Delete column C (existing column D shifts left to become C)
wks.deleteColumn(3, 1); 
```

## 📜 Changelog
<details>
<summary><b>Detailed Change Log</b></summary>

### 2026-03-27: Formula Engine, Advanced Charts, and Style Optimization
- **Formula Engine (`XLFormulaEngine`)**: Introduced a standalone AST parser and evaluation engine supporting 45+ built-in functions (Math, Logic, Text). Calculates results in C++ without Excel.
- **Chart Builder API**: Added robust support for creating complex charts with custom series styling and strictly compliant OOXML rendering.
- **Dynamic Structural Mutation**: Implemented `insertRow`, `deleteRow`, `insertColumn`, and `deleteColumn` with full coordinate shifting logic.
- **O(1) Style Deduplication**: Implemented a canonical XML fingerprinting and hash-based caching mechanism (`findOrCreateStyle`) across all format sub-pools (Fonts, Fills, Borders) to eliminate `styles.xml` bloat during bulk operations.
- **Concurrency**: Fixed race conditions in multi-threaded read environments.

### 2026-03-18: Enhanced Data Tables & AutoFilter
- **Table High-level API**: Implemented a comprehensive `XLTables` interface for creating and managing Excel Tables.
- **Totals Row Support**: Added support for `showTotalsRow`, `totalsRowCount`, and column-level aggregate functions (`Sum`, `Average`, `Count`, etc.) via `XLTableColumn`.
- **Advanced AutoFilters**: Enhanced `XLAutoFilter` to support custom logical rules (e.g., `greaterThan`) and value-based filtering.
- **Ergonomic Resizing**: Added `resizeToFitData()` to automatically snap table boundaries to contiguous worksheet data.
- **OOXML Compliance Fixes**: Standardized boolean XML attributes to `1`/`0` and optimized node ordering to ensure 100% compatibility with MS Excel's strict validation.
- **Performance & Testing**: Added comprehensive Catch2 unit tests and OOXML structure verification tests.

### 2026-03-16: Architectural Migration & Privacy Cleanup
- **Modernized File Creation**: Completely abandoned the hardcoded, 7.7KB hex-encoded binary `.xlsx` template. Migrated to dynamic, `constexpr std::string_view` XML string templates (inspired by Excelize), significantly improving code readability and maintainability without sacrificing C++ performance.
- **Privacy & Metadata Scrubbing**: Removed legacy workaround code that was previously necessary to scrub the original author's local file paths and revision histories embedded within the old binary payload. The new XML templates guarantee a 100% clean and pristine initial state for all generated documents.
- **Zip Archive Reliability**: Fixed a bug where files dynamically injected into memory during creation were occasionally discarded upon `save()` due to missing internal modification flags in `XLZipArchive`.
- **Enhanced Validation Compliance**: Embedded strict MS Excel required namespaces and `xl/theme/theme1.xml` to prevent any "file corruption / recovery" warnings when opening programmatically generated files.

### 2026-03-15: Feature Expansion & Robustness
- **Implemented Rich Text**: Added `XLRichText` and `XLRichTextRun` for multi-format text segments within cells, including support for font colors and styles.
- **AutoFilter Support**: New API to set and manage worksheet filters.
- **Workbook Defined Names**: Implemented `XLDefinedNames` for managing global and local named ranges.
- **Page Setup & Print Options**: Added comprehensive control over margins, orientation, paper size, and print-specific settings.
- **Granular Sheet Protection**: Enhanced protection with granular control over user permissions (sorting, formatting, filtering).
- **DateTime Overhaul**: Enhanced `XLDateTime` with `std::chrono` support, `fromString`/`toString` methods, and a convenient `now()` function.
- **Internal OOXML Verification**: Migrated structure verification tests to use internal APIs, removing external dependencies like `unzip` for testing.
- **Stability Fixes**: Resolved potential segmentation faults in protection property access and fixed symbol conflicts in test builds.

### 2026-03-14: Dependency Cleanup & Data Integrity
- **Removed `nowide` Dependency**: Migrated to C++17 `std::filesystem::u8path` for cross-platform UTF-8 path support, reducing the library footprint and simplifying build logic.
- **Fixed Document Properties**: Resolved issues where Title, Subject, and Creator were not correctly updated in `core.xml` due to OOXML namespace handling.
- **XML Format Optimization**: Switched to standard XML formatting to ensure XML declarations are preserved, improving compatibility with Excel and other spreadsheet viewers.
- **CI Enhancement**: Added caching for dependencies in GitHub Actions to speed up build times.

### 2026-02-28: Major Refactor & Feature Update
- **Unified Build System**: Consolidated all sub-module CMake configurations into a single root `CMakeLists.txt`.
- **Optimization Suite**: Added LTO support and platform-specific Release optimizations (`/O2`, `-O3`, dead-code stripping).
- **Image Support**: Implemented `XLDrawing` and aspect-ratio aware image insertion.
- **Data Validation**: Enhanced data validation support with full CRUD operations and Excel-compliant XML serialization.
- **Enhanced Testing**: Merged test suite into main build flow with automatic test data handling.

</details>

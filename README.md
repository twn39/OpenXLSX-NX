# OpenXLSX

[![OpenXLSX CI](https://github.com/twn39/OpenXLSX/actions/workflows/ci.yml/badge.svg)](https://github.com/twn39/OpenXLSX/actions/workflows/ci.yml)

OpenXLSX is a high-performance C++ library for reading, writing, creating, and modifying Microsoft Excel® files in the `.xlsx` format. It is designed to be fast, cross-platform, and has minimal external dependencies.

## 🚀 Key Features

- **Modern C++**: Built with C++17, ensuring type safety and modern abstractions.
- **High Performance**: Optimized for speed, capable of processing millions of cells per second.
- **LTO Support**: Optional Link-Time Optimization (LTO/IPO) for maximum execution performance.
- **Comprehensive Image Support**: Insert and read images (PNG/JPEG) with automatic dimension detection and aspect ratio preservation.
- **Rich Formatting**: Support for styles, fonts, fills, borders, and conditional formatting.
- **Worksheet Management**: Create, clone, rename, and delete worksheets with validation.
- **Data Integrity**: Enforces OOXML standards, including strict XML declarations and metadata handling.
- **Advanced Iterators**: Efficient cell and row iterators that minimize XML overhead.
- **Modern Testing**: Integrated with **Catch2 v3.13.0** for robust verification.
- **Simplified Build**: A unified single-root CMake configuration for easy integration.

## 🛠 Quick Start

```cpp
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("Example.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Writing data
    wks.cell("A1").value() = "Hello, OpenXLSX!";
    wks.cell("B1").value() = 42;

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
**All string input and output must be in UTF-8 encoding.** OpenXLSX does not perform internal encoding conversion. Ensure your source files are saved in UTF-8.

### Indexing
- **Sheet Indexing**: 1-based (consistent with Excel).
- **Row/Column Indexing**: Generally 1-based where it follows Excel conventions.

### Performance & Optimizations
The build system includes platform-specific optimizations for `Release` builds (e.g., `/O2` on MSVC, `-O3` on GCC/Clang) and supports **LTO (Link-Time Optimization)** which can be toggled via `OPENXLSX_ENABLE_LTO`.

## 🤝 Credits
- [PugiXML](https://pugixml.org/) - Fast XML parsing.
- [libzip](https://libzip.org/) & [zlib-ng](https://github.com/zlib-ng/zlib-ng) - Fast and compatible ZIP archive handling.
- [std::filesystem] - UTF-8 file system support on Windows (C++17).

---

<details>
<summary><b>Detailed Change Log (Feb 2026)</b></summary>

### 2026-02-28: Major Refactor & Feature Update
- **Unified Build System**: Consolidated all sub-module CMake configurations into a single root `CMakeLists.txt`.
- **Optimization Suite**: Added LTO support and platform-specific Release optimizations (`/O2`, `-O3`, dead-code stripping).
- **Image Support**: Implemented `XLDrawing` and aspect-ratio aware image insertion.
- **Cleanup**: Removed deprecated `Examples`, legacy `Makefile.GNU`, and redundant config files.
- **Enhanced Testing**: Merged test suite into main build flow with automatic test data handling.

</details>

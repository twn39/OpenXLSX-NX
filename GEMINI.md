# GEMINI.md - OpenXLSX Context

This file provides context and instructions for AI agents working on the OpenXLSX project.

## Project Overview
OpenXLSX is a high-performance C++ library for reading, writing, creating, and modifying Microsoft Excel® files in the `.xlsx` format. It is designed to be fast, cross-platform, and has minimal external dependencies.

### Key Technologies
- **Language**: C++17 (required for `std::variant`, `std::optional`, and other features).
- **Build System**: Unified CMake 3.15+ (Single root `CMakeLists.txt`).
- **XML Parsing**: [PugiXML](https://pugixml.org/) (DOM-based, fast).
- **ZIP Handling**: [Zippy](https://github.com/troldal/Zippy) (wrapper around `miniz`).
- **Unicode**: [Boost.Nowide](https://github.com/boostorg/nowide) (UTF-8 support on Windows).
- **Testing & Benchmarking**: [Catch2 v3](https://github.com/catchorg/Catch2) (Integrated via `FetchContent`).

## Building and Running

### Building with CMake
The project uses a simplified, single-root CMake configuration.
```bash
mkdir build && cd build
cmake .. -DCMAKE_BUILD_TYPE=Release
cmake --build . -j8
```

### Running Tests
Tests are built as a standalone executable in the build directory root:
```bash
./OpenXLSXTests
```

### Running Benchmarks
Benchmarks are also powered by Catch2 v3 and can be enabled via CMake:
```bash
cmake -DOPENXLSX_BUILD_BENCHMARKS=ON ..
cmake --build . --target OpenXLSXBenchmark
./OpenXLSXBenchmark "[.benchmark]"
```

## Development Conventions

### Coding Standards
- **Naming**: Member functions use `camelCase` (e.g., `worksheetCount()`). Classes use `PascalCase` (e.g., `XLDocument`).
- **Unicode**: All string input and output **MUST be in UTF-8 encoding**.
- **Indexing**: Sheet indexing is **1-based** (consistent with Excel). Row and column indexing is also generally 1-based.
- **Performance**: OpenXLSX is optimized for speed using a DOM-based XML parser. 
- **Modern API**: Prefer the explicit `doc.create(filename, XLForceOverwrite)` over the deprecated single-argument version.

### Error Handling
- Uses exceptions defined in `headers/XLException.hpp`.
- Common exceptions: `XLInputError`, `XLInternalError`, `XLPropertyError`, `XLCellAddressError`.

## Key Files and Architecture

### Core Library (`OpenXLSX/`)
- `OpenXLSX.hpp`: Main umbrella header.
- `headers/XLDocument.hpp`: Primary entry point.
- `headers/XLWorkbook.hpp`: Manages sheets within a document.
- `headers/XLSheet.hpp`: Represents individual worksheets.
- `headers/XLCell.hpp`: Cell value, formula, and formatting access.
- `sources/`: Implementation files, globally tracked in root `CMakeLists.txt`.

### Testing & Benchmarking
- `Tests/`: Catch2 v3 test suite (`testXL*.cpp`).
- `Benchmarks/`: Catch2 v3 benchmark suite (`Benchmark.cpp`).
- Test data (images) is automatically copied to the build directory during the build process.

### Third Party (`third_party/`)
- Local versions of `nowide` and `zippy`. 
- `pugixml`, `fmt`, and `fast_float` are managed via `FetchContent` in the root `CMakeLists.txt`.

## Usage Tips for AI
- **Single Build Entry**: Do not look for `CMakeLists.txt` in subdirectories; everything is in the root.
- **Clean Code**: The project has been refactored to remove all compiler warnings (except for third-party libs). Ensure new code maintains this standard.
- **Testing**: Always run `OpenXLSXTests` after changes. For performance-related changes, run `OpenXLSXBenchmark`.
- **API Preference**: Use `XLForceOverwrite` explicitly when creating or saving files in tests.

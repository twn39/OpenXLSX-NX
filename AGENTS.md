# GEMINI.md - OpenXLSX Context

This file provides context and instructions for AI agents working on the OpenXLSX project.

## Project Overview
OpenXLSX is a high-performance C++ library for reading, writing, creating, and modifying Microsoft Excel® files in the `.xlsx` format. It is designed to be fast, cross-platform, and has minimal external dependencies.

### Key Technologies
- **Language**: C++17 (required for `std::variant`, `std::optional`, and other features).
- **Build System**: Unified CMake 3.15+ (Single root `CMakeLists.txt`).
- **XML Parsing**: [PugiXML](https://pugixml.org/) (DOM-based, fast).
- **ZIP Handling**: [minizip-ng](https://github.com/zlib-ng/minizip-ng) with [zlib-ng](https://github.com/zlib-ng/zlib-ng) (SIMD optimized, Zip64 support).
- **Testing & Benchmarking**: [Catch2 v3](https://github.com/catchorg/Catch2) (Integrated via `FetchContent`).
- **Safety**: [Microsoft GSL v4.2.1](https://github.com/microsoft/GSL) (C++ Core Guidelines Support Library).

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

### Safety & GSL (Guidelines Support Library)
The project utilizes Microsoft GSL to enforce C++ Core Guidelines and ensure memory safety without runtime performance penalties. 

**Core Rules for GSL Usage:**
1.  **Pointer Safety**: Use `gsl::not_null<T*>` for pointers that must never be null (e.g., in constructors). This moves the null check to the call site and documents the requirement in the type system.
2.  **Bounds Safety**: Use `gsl::span<T>` for all raw array operations, buffer manipulations, and binary data parsing. Avoid raw `T* + size` combinations.
3.  **Safe Conversions**: Use `gsl::narrow<T>` when casting between numeric types (e.g., `size_t` to `int32_t`) where truncation or overflow is possible. It will throw if the value cannot be represented. Use `gsl::narrow_cast<T>` only for intentional, documented truncation.
4.  **Contract Programming**: Use `Expects()` for function preconditions (e.g., verifying a buffer is large enough) and `Ensures()` for postconditions.
5.  **Safe Indexing**: Prefer `gsl::at()` for container access where bounds checking is critical and cannot be verified at compile-time.

**Best Practices in OpenXLSX:**
- **Backward Compatibility**: When updating legacy utility functions, provide a new `gsl::span`-based core implementation and wrap it with the original pointer-based function for compatibility.
- **Excel Limits**: Since Excel has strict limits (1,048,576 rows and 16,384 columns), always use `gsl::narrow` when converting external data or large integers to row/column indices.
- **Zero-Cost Abstractions**: Use GSL types like `span` and `not_null` internally to help the compiler optimize while providing high safety guarantees.

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
- Local version of `pugixml`.
- `fmt`, `fast_float`, `Microsoft GSL`, `zlib-ng`, and `minizip-ng` are managed via `FetchContent` in the root `CMakeLists.txt`.

## Usage Tips for AI
- **Single Build Entry**: Do not look for `CMakeLists.txt` in subdirectories; everything is in the root.
- **GSL First**: When implementing new logic or refactoring, prioritize GSL safety (span, not_null, narrow). Do not use raw `memcpy` or `strcpy`; use `gsl::copy` or `span` slices.
- **Clean Code**: The project has been refactored to remove all compiler warnings (except for third-party libs). Ensure new code maintains this standard.
- **Testing**: Always run `OpenXLSXTests` after changes. For performance-related changes, run `OpenXLSXBenchmark`.
- **API Preference**: Use `XLForceOverwrite` explicitly when creating or saving files in tests.

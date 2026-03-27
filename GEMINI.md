# GEMINI.md - OpenXLSX Context

This file provides context and instructions for AI agents working on the OpenXLSX project.

## Project Overview
OpenXLSX is a high-performance C++ library for reading, writing, creating, and modifying Microsoft Excel® files in the `.xlsx` format. It is designed to be fast, cross-platform, and has minimal external dependencies.

### Key Technologies
- **Language**: C++17 (required for `std::variant`, `std::optional`, and other features).
- **Build System**: Unified CMake 3.15+ (Single root `CMakeLists.txt`).
- **XML Parsing**: [PugiXML](https://pugixml.org/) (DOM-based, fast).
- **ZIP Handling**: `libzip` and `zlib-ng` (fetched via CMake) for maximum Excel compatibility and zero Data Descriptor errors.
- **Unicode**: `std::filesystem` (UTF-8 support on Windows).
- **Testing & Benchmarking**: [Catch2 v3](https://github.com/catchorg/Catch2) (Integrated via `FetchContent`).

## Project Structure
```text
OpenXLSX/
├── CMakeLists.txt              # Root CMake configuration
├── GEMINI.md                   # AI Project Context (Internal use)
├── README.md                   # Project documentation
├── Benchmarks/                 # Benchmark source files
├── Documentation/              # Doxygen configuration
├── OpenXLSX/                   # Core library source and headers
│   ├── headers/                # Public and private headers
│   └── sources/                # Implementation files
├── Tests/                      # Catch2 test suite
└── third_party/                # External dependencies
    ├── excelize/               # Reference implementation in Go (Read-only)
    └── openpyxl/               # Reference implementation in Python (Read-only)
```

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
- **Document Creation**: New documents are dynamically constructed using clean, `constexpr std::string_view` XML templates (similar to `excelize`). No hardcoded binary payloads are used.
- **Modern API**: Prefer the explicit `doc.create(filename, XLForceOverwrite)` over the deprecated single-argument version.
- **Commenting**: Comment why, not how.

### Error Handling
- Uses exceptions defined in `headers/XLException.hpp`.
- Common exceptions: `XLInputError`, `XLInternalError`, `XLPropertyError`, `XLCellAddressError`.

## Git and Repository Rules
- **DO NOT COMMIT**: `GEMINI.md`, `third_party/excelize`, and `third_party/openpyxl` must never be committed to the Git repository.
- **NO .gitignore**: Do not add `GEMINI.md`, `third_party/excelize`, or `third_party/openpyxl` to `.gitignore`. They are intended to remain as untracked files in the local workspace.

## Reference Implementation
- **Excelize**: The code in `third_party/excelize` (a popular Go library) serves as a functional source code reference for implementing Excel features. When adding new functionality, consult the corresponding logic in Excelize to ensure compatibility and correctness.
- **Openpyxl**: The code in `third_party/openpyxl` (a popular Python library) also serves as a functional source code reference. It can be consulted alongside Excelize for additional implementation insights.

## Usage Tips for AI
- **Skills Requirement**: Always activate `gsl-reference` and `cpp-best-practices` skills before starting any implementation or refactoring task.
- **Single Build Entry**: Do not look for `CMakeLists.txt` in subdirectories; everything is in the root.
- **Clean Code**: The project has been refactored to remove all compiler warnings (except for third-party libs). Ensure new code maintains this standard.
- **Testing**: Always run `OpenXLSXTests` after changes. For performance-related changes, run `OpenXLSXBenchmark`.
- **API Preference**: Use `XLForceOverwrite` explicitly when creating or saving files in tests.


## Shell Compatibility Rule
You are operating in a environment where the default shell might be `fish`. 
- **DO NOT** use Heredoc syntax (e.g., `cat <<EOF`). It is not supported by `fish`.
- **To create/overwrite files**, use `printf` with single quotes to avoid unintended variable expansion. Example: `printf 'line1\nline2' > file.txt`.
- If you must use complex Bash-specific syntax, wrap the entire command in bash: `bash -c 'complex_command_here'`.
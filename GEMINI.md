# OpenXLSX-NX Project Context

This file provides foundational instructions and architectural context for Gemini CLI interactions with the **OpenXLSX-NX** codebase.

## 🚀 Project Overview
**OpenXLSX-NX** is a high-performance, modern C++17 library for reading, writing, and manipulating Microsoft Excel® (.xlsx) files. It emphasizes speed, enterprise-grade security, and cross-engine compatibility (openpyxl, excelize, Google Sheets, etc.).

### Key Technologies
- **Language**: C++17 (requires standard compliance).
- **Build System**: CMake (3.15+) with **Ninja** generator preferred.
- **Acceleration**: Integrated **ccache** and **Unity Build** support.
- **Dependencies**: Managed via **Git Submodules** in `third_party/`.
- **Testing**: **Catch2 v3.14.0** for unit/integration testing; **LLVM libFuzzer** for formula AST fuzzing.
- **Core Libraries**: `pugixml` (XML), `libzip` & `zlib-ng` (Compression), `fmt` (Formatting), `mbedtls` (Cryptography), `fast_float` (Numeric parsing).

## 🛠 Building and Running

### AI Context Management (GEMINI.md)
- **Do NOT add** `GEMINI.md` to `.gitignore`.
- **Do NOT commit** `GEMINI.md` to the Git repository. It serves as a local, volatile context for AI interactions.

### Build Prerequisites
- Git, CMake, and a C++17 compliant compiler (GCC 8+, Clang 7+, MSVC 2019+).
- Optional but recommended: `Ninja`, `ccache`.

### Commands
- **Initialize Dependencies**:
  ```bash
  git submodule update --init --recursive
  ```
- **Configure & Build (Optimized)**:
  ```bash
  mkdir build && cd build
  cmake .. -G Ninja -DCMAKE_BUILD_TYPE=Release
  ninja
  ```
- **Run Full Test Suite**:
  ```bash
  ./OpenXLSXTests
  ```
- **Run Benchmarks**:
  ```bash
  # Re-configure with benchmarks enabled
  cmake .. -G Ninja -DOPENXLSX_BUILD_BENCHMARKS=ON
  ninja
  ./OpenXLSXBenchmark
  ```

## 🏗 Architecture & Conventions

### Directory Structure
- `OpenXLSX/headers/`: Public and internal header files.
- `OpenXLSX/sources/`: Implementation files.
- `Tests/`: Comprehensive test suite and test fixtures.
- `third_party/`: External dependencies as git submodules.
- `vendors/`: Functional reference code from other libraries. **Do NOT add** to `.gitignore` and **do NOT commit** to the Git repository.
- `.github/workflows/`: CI/CD pipelines (Ninja enabled).

### Development Mandates
- **Deep Analysis & Confirmation**: When instructed to perform a "detailed analysis" or "in-depth analysis" of a specific feature or issue, **DO NOT modify code immediately**. First, confirm the feature/issue, provide a detailed analysis, and propose a concrete, elegant implementation or fix. Wait for the user's explicit confirmation before proceeding with any code modifications.
- **Encoding**: All string inputs/outputs **MUST** be UTF-8.
- **Indexing**: Follow Excel's **1-based** indexing for sheets, rows, and columns.
- **Memory Safety**: Hard zero-leak policy. Use RAII and modern C++ smart pointers. Avoid manual memory management.
- **Thread Safety**: Supports concurrent writes to **different** worksheets. Writing to the same worksheet is unsafe.
- **Style**: Adhere to `.clang-format` rules. Use PascalCase for classes and camelCase for functions/variables.
- **Error Handling**: Use `XLException` for library-specific errors. Anticipate corrupt XML inputs.

### Maintenance Tasks
- When adding new files, update the `file(GLOB ...)` logic in `CMakeLists.txt` or manually list them if required by specific architectural patterns.
- After updating submodules, always ensure `.gitmodules` and the submodule commit pointers are staged together.
- For performance-sensitive changes, verify against `Benchmarks/Benchmark.cpp`.

## 🧪 Testing Guidelines
- **Always** add a test case in `Tests/` for bug fixes or new features.
- Use Catch2's `GENERATE` for data-driven testing where applicable.
- Ensure all 220,000+ assertions pass before confirming a task.
- Fuzzing: Use `OPENXLSX_BUILD_FUZZERS` to validate AST parser changes.

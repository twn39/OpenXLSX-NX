cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX-NX VERSION 1.10.0 LANGUAGES CXX C)

# Set C++ standard
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Ensure all targets (including static dependencies) are compiled with -fPIC
set(CMAKE_POSITION_INDEPENDENT_CODE ON)

# Build options
option(OPENXLSX_BUILD_TESTS "Build and run library tests" ON)
option(OPENXLSX_BUILD_BENCHMARKS "Build benchmark programs" OFF)
option(OPENXLSX_BUILD_FUZZERS "Build libFuzzer targets (Clang only)" OFF)
option(OPENXLSX_CREATE_DOCS "Build library documentation" OFF)
option(OPENXLSX_ENABLE_LTO "Enable Link-Time Optimization (LTO/IPO) if supported" OFF)
option(OPENXLSX_ENABLE_SANITIZERS "Enable ASan and UBSan for memory/undefined behavior analysis (GCC/Clang only)" OFF)
option(OPENXLSX_OPTIMIZE_NATIVE "Enable architecture-native optimizations (-march=native or /arch:AVX2) for peak performance" OFF)

set(OPENXLSX_LIBRARY_TYPE "STATIC" CACHE STRING "Library type (STATIC or SHARED)")
option(OPENXLSX_ENABLE_UNITY_BUILD "Enable Unity Build (faster compilation)" ON)

#=======================================================================================================================
# Build Accelerators (ccache)
#=======================================================================================================================
find_program(CCACHE_PROGRAM ccache)
if(CCACHE_PROGRAM)
    set(CMAKE_CXX_COMPILER_LAUNCHER "${CCACHE_PROGRAM}")
    set(CMAKE_C_COMPILER_LAUNCHER "${CCACHE_PROGRAM}")
    message(STATUS "ccache found and enabled")
endif()

#=======================================================================================================================
# External Libraries
#=======================================================================================================================

# Set up fast_float
set(fast_float_SOURCE_DIR ${CMAKE_CURRENT_SOURCE_DIR}/third_party/fast_float)
add_subdirectory(third_party/fast_float)

# Set up GSL (Microsoft Guidelines Support Library)
add_subdirectory(third_party/GSL)

# Set up fmt
set(fmt_SOURCE_DIR ${CMAKE_CURRENT_SOURCE_DIR}/third_party/fmt)
add_subdirectory(third_party/fmt)

# Set up pugixml
set(PUGIXML_XPATH OFF CACHE BOOL "" FORCE)
add_subdirectory(third_party/pugixml)

# Set up Catch2 (if tests or benchmarks are enabled)
if(OPENXLSX_BUILD_TESTS OR OPENXLSX_BUILD_BENCHMARKS)
    add_subdirectory(third_party/Catch2)
endif()

# Set up zlib-ng (libzip dependency)
set(ZLIB_ENABLE_TESTS OFF CACHE BOOL "" FORCE)
set(BUILD_TESTING OFF CACHE BOOL "" FORCE)
set(ZLIB_COMPAT ON CACHE BOOL "" FORCE)
set(WITH_SANITIZER OFF CACHE STRING "" FORCE)
add_subdirectory(third_party/zlib)
set(zlib_SOURCE_DIR ${CMAKE_CURRENT_SOURCE_DIR}/third_party/zlib)
set(zlib_BINARY_DIR ${CMAKE_CURRENT_BINARY_DIR}/third_party/zlib)

# Create the alias target that many projects (like libzip) expect
if(NOT TARGET ZLIB::ZLIB)
    if(TARGET zlib)
        add_library(ZLIB::ZLIB ALIAS zlib)
    elseif(TARGET zlibstatic)
        add_library(ZLIB::ZLIB ALIAS zlibstatic)
    endif()
endif()

# Explicitly set ZLIB variables for libzip's find_package(ZLIB) call
# Note: On Windows, zconf.h is generated in the binary directory.
set(ZLIB_INCLUDE_DIR "${zlib_SOURCE_DIR};${zlib_BINARY_DIR}" CACHE PATH "ZLIB include directory" FORCE)
set(ZLIB_LIBRARY zlib CACHE FILEPATH "ZLIB library" FORCE)
set(ZLIB_LIBRARIES zlib CACHE FILEPATH "ZLIB libraries" FORCE)
set(ZLIB_FOUND TRUE CACHE BOOL "ZLIB found" FORCE)
set(ZLIB_VERSION "1.3.1" CACHE STRING "ZLIB version" FORCE)

# Set up libzip
set(ENABLE_BZIP2 OFF CACHE BOOL "" FORCE)
set(ENABLE_LZMA OFF CACHE BOOL "" FORCE)
set(ENABLE_ZSTD OFF CACHE BOOL "" FORCE)
set(ENABLE_GNUTLS OFF CACHE BOOL "" FORCE)
set(ENABLE_MBEDTLS OFF CACHE BOOL "" FORCE)
set(ENABLE_OPENSSL OFF CACHE BOOL "" FORCE)
set(BUILD_TOOLS OFF CACHE BOOL "" FORCE)
set(BUILD_REGRESS OFF CACHE BOOL "" FORCE)
set(BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
set(BUILD_DOC OFF CACHE BOOL "" FORCE)
set(BUILD_OSSFUZZ OFF CACHE BOOL "" FORCE)
set(LIBZIP_DO_INSTALL OFF CACHE BOOL "" FORCE)
add_subdirectory(third_party/libzip)

# Set up mbedtls for cryptography (AES, SHA)
set(ENABLE_TESTING OFF CACHE BOOL "" FORCE)
set(ENABLE_PROGRAMS OFF CACHE BOOL "" FORCE)
add_subdirectory(third_party/mbedtls)
add_subdirectory(third_party/unordered_dense)


# Sources
file(GLOB OPENXLSX_SOURCES OpenXLSX/sources/*.cpp)

# Library Target
add_library(OpenXLSX ${OPENXLSX_LIBRARY_TYPE} ${OPENXLSX_SOURCES})
add_library(OpenXLSX::OpenXLSX ALIAS OpenXLSX)

if(OPENXLSX_ENABLE_UNITY_BUILD)
    set_target_properties(OpenXLSX PROPERTIES UNITY_BUILD ON)
endif()

target_include_directories(OpenXLSX
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/OpenXLSX>
        $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/OpenXLSX/headers>
        $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>
    SYSTEM PUBLIC
        $<BUILD_INTERFACE:${fast_float_SOURCE_DIR}/include>
        $<BUILD_INTERFACE:${fmt_SOURCE_DIR}/include>
        $<BUILD_INTERFACE:${zlib_SOURCE_DIR}>
        $<BUILD_INTERFACE:${zlib_BINARY_DIR}>
)

target_link_libraries(OpenXLSX PUBLIC pugixml libzip::zip zlib Microsoft.GSL::GSL mbedcrypto unordered_dense::unordered_dense)
target_compile_definitions(OpenXLSX PUBLIC FMT_HEADER_ONLY)

if ("${OPENXLSX_LIBRARY_TYPE}" STREQUAL "STATIC")
    target_compile_definitions(OpenXLSX PUBLIC OPENXLSX_STATIC_DEFINE)
endif ()

# Enable Link-Time Optimization (LTO/IPO) if supported
if(OPENXLSX_ENABLE_LTO)
    include(CheckIPOSupported)
    check_ipo_supported(RESULT result OUTPUT output)
    if(result)
        set_property(TARGET OpenXLSX PROPERTY INTERPROCEDURAL_OPTIMIZATION TRUE)
    endif()
endif()

# Export Header (Required by headers)
include(GenerateExportHeader)
generate_export_header(OpenXLSX
    BASE_NAME openxlsx
    EXPORT_FILE_NAME OpenXLSX-Exports.hpp
    EXPORT_MACRO_NAME OPENXLSX_EXPORT
)

# Basic Compiler Flags & Release Optimizations
if(MSVC)
    target_compile_options(OpenXLSX PUBLIC /W4 /permissive- /utf-8 $<$<CONFIG:Release>:/O2 /Ot /Gy /Zc:inline>)
    target_compile_definitions(OpenXLSX PUBLIC _CRT_SECURE_NO_WARNINGS)
    
    if(OPENXLSX_ENABLE_SANITIZERS)
        target_compile_options(OpenXLSX PRIVATE /fsanitize=address)
    endif()
else()
    target_compile_options(OpenXLSX PRIVATE -Wall -Wextra $<$<CONFIG:Release>:-O3 -ffunction-sections -fdata-sections>)
    if(APPLE)
        target_link_options(OpenXLSX PRIVATE $<$<CONFIG:Release>:-Wl,-dead_strip>)
    else()
        target_link_options(OpenXLSX PRIVATE $<$<CONFIG:Release>:-Wl,--gc-sections>)
    endif()

    if(OPENXLSX_ENABLE_SANITIZERS)
        target_compile_options(OpenXLSX PRIVATE -fsanitize=address,undefined -fno-omit-frame-pointer)
        target_link_options(OpenXLSX PRIVATE -fsanitize=address,undefined)
    endif()

    if(OPENXLSX_OPTIMIZE_NATIVE)
        if(MSVC)
            target_compile_options(OpenXLSX PRIVATE $<$<CONFIG:Release>:/arch:AVX2>)
        elseif(APPLE AND CMAKE_SYSTEM_PROCESSOR MATCHES "arm64")
            target_compile_options(OpenXLSX PRIVATE $<$<CONFIG:Release>:-mcpu=apple-m1>)
        else()
            target_compile_options(OpenXLSX PRIVATE $<$<CONFIG:Release>:-march=native>)
        endif()
    endif()
endif()

#=======================================================================================================================
# TESTS (Merged configuration)
#=======================================================================================================================
if(OPENXLSX_BUILD_TESTS)
    enable_testing()
    file(GLOB TEST_SOURCES Tests/testXL*.cpp)
    # Exclude fuzzer from regular Catch2 tests
    list(FILTER TEST_SOURCES EXCLUDE REGEX "testXL.*Fuzzer\\.cpp")
    
    add_executable(OpenXLSXTests ${TEST_SOURCES})
    include(Catch)
    catch_discover_tests(OpenXLSXTests)
    target_link_libraries(OpenXLSXTests PRIVATE OpenXLSX::OpenXLSX Catch2::Catch2WithMain)
    
    # Disable LTO for Catch2 test executable to prevent GCC LTO SIGILL bugs during static initialization
    set_property(TARGET OpenXLSXTests PROPERTY INTERPROCEDURAL_OPTIMIZATION FALSE)

    if(OPENXLSX_ENABLE_SANITIZERS AND NOT MSVC)
        target_compile_options(OpenXLSXTests PRIVATE -fsanitize=address,undefined -fno-omit-frame-pointer)
        target_link_options(OpenXLSXTests PRIVATE -fsanitize=address,undefined)
    endif()

    # Copy test data to build directory
    add_custom_command(TARGET OpenXLSXTests POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:OpenXLSXTests>/Tests
        COMMAND ${CMAKE_COMMAND} -E make_directory ${CMAKE_BINARY_DIR}/Tests
        COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:OpenXLSXTests>/Tests/Fixtures
        COMMAND ${CMAKE_COMMAND} -E make_directory ${CMAKE_BINARY_DIR}/Tests/Fixtures
        COMMAND ${CMAKE_COMMAND} -E copy_if_different 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.png 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.jpg
            ${CMAKE_CURRENT_LIST_DIR}/Tests/shared_formula_test.xlsx
            ${CMAKE_CURRENT_LIST_DIR}/Tests/excelize_macro.xlsm
            $<TARGET_FILE_DIR:OpenXLSXTests>/Tests/
        COMMAND ${CMAKE_COMMAND} -E copy_if_different 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.png 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.jpg
            ${CMAKE_CURRENT_LIST_DIR}/Tests/shared_formula_test.xlsx
            ${CMAKE_CURRENT_LIST_DIR}/Tests/excelize_macro.xlsm
            ${CMAKE_BINARY_DIR}/Tests/
        COMMAND ${CMAKE_COMMAND} -E copy_if_different 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/openpyxl_generated.xlsx
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/excelize_generated.xlsx
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/google_spreadsheets.xlsx
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/tencent_sheet.xlsx
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/Encrypted_Agile.xlsx
            ${CMAKE_BINARY_DIR}/Tests/Fixtures/
        COMMAND ${CMAKE_COMMAND} -E copy_if_different 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/openpyxl_generated.xlsx
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/excelize_generated.xlsx
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/google_spreadsheets.xlsx
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/tencent_sheet.xlsx
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/Encrypted_Agile.xlsx
            $<TARGET_FILE_DIR:OpenXLSXTests>/Tests/Fixtures/
    )
endif()

#=======================================================================================================================
# FUZZERS (Requires Clang)
#=======================================================================================================================
if(OPENXLSX_BUILD_FUZZERS)
    add_executable(OpenXLSXFuzzer Tests/testXLFormulaFuzzer.cpp)
    add_executable(OpenXLSXOpenFuzzer Tests/testXLDocumentOpenFuzzer.cpp)
    target_link_libraries(OpenXLSXFuzzer PRIVATE OpenXLSX::OpenXLSX)
    target_link_libraries(OpenXLSXOpenFuzzer PRIVATE OpenXLSX::OpenXLSX)
    target_compile_options(OpenXLSXFuzzer PRIVATE -fsanitize=fuzzer,address,undefined -fno-omit-frame-pointer)
    target_compile_options(OpenXLSXOpenFuzzer PRIVATE -fsanitize=fuzzer,address,undefined -fno-omit-frame-pointer)
    target_link_options(OpenXLSXFuzzer PRIVATE -fsanitize=fuzzer,address,undefined)
    target_link_options(OpenXLSXOpenFuzzer PRIVATE -fsanitize=fuzzer,address,undefined)
endif()

if(OPENXLSX_BUILD_BENCHMARKS)
    add_executable(OpenXLSXBenchmark Benchmarks/Benchmark.cpp)
    target_link_libraries(OpenXLSXBenchmark PRIVATE OpenXLSX::OpenXLSX Catch2::Catch2WithMain)
    set_property(TARGET OpenXLSXBenchmark PROPERTY INTERPROCEDURAL_OPTIMIZATION FALSE)
endif()

#=======================================================================================================================
# DOCUMENTATION (Doxygen)
#=======================================================================================================================
if(OPENXLSX_CREATE_DOCS)
    find_package(Doxygen)
    if(DOXYGEN_FOUND)
        set(DOXYGEN_IN ${CMAKE_CURRENT_SOURCE_DIR}/Documentation/Doxyfile.in)
        set(DOXYGEN_OUT ${CMAKE_CURRENT_BINARY_DIR}/Doxyfile)
        configure_file(${DOXYGEN_IN} ${DOXYGEN_OUT} @ONLY)
        add_custom_target(OpenXLSXDocs ALL
            COMMAND ${DOXYGEN_EXECUTABLE} ${DOXYGEN_OUT}
            WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
            COMMENT "Generating API documentation with Doxygen"
            VERBATIM
        )
    endif()
endif()

#=======================================================================================================================
# INSTALLATION
#=======================================================================================================================
include(GNUInstallDirs)

install(TARGETS OpenXLSX
    EXPORT OpenXLSXTargets
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
)

install(FILES OpenXLSX/OpenXLSX.hpp
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX
)

install(DIRECTORY OpenXLSX/headers/
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX/headers
    FILES_MATCHING PATTERN "*.hpp"
)

install(FILES ${CMAKE_CURRENT_BINARY_DIR}/OpenXLSX-Exports.hpp
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX
)


install(DIRECTORY third_party/fmt/include/fmt
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)

install(DIRECTORY third_party/GSL/include/gsl
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)

install(DIRECTORY third_party/unordered_dense/include/ankerl
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)

install(TARGETS zip
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
)

install(FILES ${CMAKE_CURRENT_BINARY_DIR}/third_party/libzip/zipconf.h
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)

install(FILES ${CMAKE_CURRENT_SOURCE_DIR}/third_party/libzip/lib/zip.h
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)

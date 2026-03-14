cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX VERSION 0.4.1 LANGUAGES CXX)

# Set C++ standard
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Build options
option(OPENXLSX_BUILD_TESTS "Build and run library tests" ON)
option(OPENXLSX_BUILD_BENCHMARKS "Build benchmark programs" OFF)
option(OPENXLSX_CREATE_DOCS "Build library documentation" OFF)
option(OPENXLSX_ENABLE_NOWIDE "Enable Nowide for UTF-8 support on Windows" ON)
option(OPENXLSX_ENABLE_LTO "Enable Link-Time Optimization (LTO/IPO) if supported" ON)

set(OPENXLSX_LIBRARY_TYPE "STATIC" CACHE STRING "Library type (STATIC or SHARED)")

#=======================================================================================================================
# External Libraries
#=======================================================================================================================
include(FetchContent)

# Set up fast_float
FetchContent_Declare(fast_float GIT_REPOSITORY https://github.com/fastfloat/fast_float.git GIT_TAG v8.2.3)
FetchContent_MakeAvailable(fast_float)

# Set up fmt
FetchContent_Declare(fmt GIT_REPOSITORY https://github.com/fmtlib/fmt.git GIT_TAG 12.1.0)
FetchContent_MakeAvailable(fmt)

# Set up pugixml
set(PUGIXML_XPATH OFF CACHE BOOL "" FORCE)
FetchContent_Declare(pugixml GIT_REPOSITORY https://github.com/zeux/pugixml.git GIT_TAG v1.15)
FetchContent_MakeAvailable(pugixml)

# Set up nowide
set(NOWIDE_INSTALL OFF CACHE BOOL "" FORCE)
set(NOWIDE_TESTING OFF CACHE BOOL "" FORCE)
FetchContent_Declare(nowide GIT_REPOSITORY https://github.com/boostorg/nowide.git GIT_TAG v11.3.1)
FetchContent_MakeAvailable(nowide)

# Set up Catch2 (if tests or benchmarks are enabled)
if(OPENXLSX_BUILD_TESTS OR OPENXLSX_BUILD_BENCHMARKS)
    FetchContent_Declare(Catch2 GIT_REPOSITORY https://github.com/catchorg/Catch2.git GIT_TAG v3.13.0)
    FetchContent_MakeAvailable(Catch2)
endif()

if(WIN32)
    set(OPENXLSX_ENABLE_NOWIDE ON)
endif()

# Sources
file(GLOB OPENXLSX_SOURCES OpenXLSX/sources/*.cpp)

# Library Target
add_library(OpenXLSX ${OPENXLSX_LIBRARY_TYPE} ${OPENXLSX_SOURCES})
add_library(OpenXLSX::OpenXLSX ALIAS OpenXLSX)

target_include_directories(OpenXLSX
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/OpenXLSX>
        $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/OpenXLSX/headers>
        $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>
    SYSTEM PUBLIC
        $<BUILD_INTERFACE:${fast_float_SOURCE_DIR}/include>
        $<BUILD_INTERFACE:${fmt_SOURCE_DIR}/include>
        $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/third_party/zippy>
)

target_link_libraries(OpenXLSX PUBLIC pugixml)
target_compile_definitions(OpenXLSX PUBLIC FMT_HEADER_ONLY)

if (OPENXLSX_ENABLE_NOWIDE)
    target_link_libraries(OpenXLSX PRIVATE nowide::nowide)
    target_compile_definitions(OpenXLSX PRIVATE ENABLE_NOWIDE)
endif ()

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
else()
    target_compile_options(OpenXLSX PRIVATE -Wall -Wextra $<$<CONFIG:Release>:-O3 -ffunction-sections -fdata-sections>)
    if(APPLE)
        target_link_options(OpenXLSX PRIVATE $<$<CONFIG:Release>:-Wl,-dead_strip>)
    else()
        target_link_options(OpenXLSX PRIVATE $<$<CONFIG:Release>:-Wl,--gc-sections>)
    endif()
endif()

#=======================================================================================================================
# TESTS (Merged configuration)
#=======================================================================================================================
if(OPENXLSX_BUILD_TESTS)
    file(GLOB TEST_SOURCES Tests/testXL*.cpp Tests/main.cpp)
    add_executable(OpenXLSXTests ${TEST_SOURCES})
    target_link_libraries(OpenXLSXTests PRIVATE OpenXLSX::OpenXLSX Catch2::Catch2WithMain)

    # Copy test data to build directory
    add_custom_command(TARGET OpenXLSXTests POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:OpenXLSXTests>/Tests
        COMMAND ${CMAKE_COMMAND} -E make_directory ${CMAKE_BINARY_DIR}/Tests
        COMMAND ${CMAKE_COMMAND} -E copy_if_different 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.png 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.jpg
            ${CMAKE_CURRENT_LIST_DIR}/Tests/shared_formula_test.xlsx
            $<TARGET_FILE_DIR:OpenXLSXTests>/Tests/
        COMMAND ${CMAKE_COMMAND} -E copy_if_different 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.png 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/test.jpg
            ${CMAKE_CURRENT_LIST_DIR}/Tests/shared_formula_test.xlsx
            ${CMAKE_BINARY_DIR}/Tests/
    )
endif()

#=======================================================================================================================
# BENCHMARKS (Merged configuration using Catch2)
#=======================================================================================================================
if(OPENXLSX_BUILD_BENCHMARKS)
    add_executable(OpenXLSXBenchmark Benchmarks/Benchmark.cpp)
    target_link_libraries(OpenXLSXBenchmark PRIVATE OpenXLSX::OpenXLSX Catch2::Catch2WithMain)
endif()

#=======================================================================================================================
# DOCUMENTATION (Subdirectory configuration)
#=======================================================================================================================
if(OPENXLSX_CREATE_DOCS)
    add_subdirectory(Documentation)
endif()

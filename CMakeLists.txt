cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX-NX VERSION 1.5.0 LANGUAGES CXX)

# Set C++ standard
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Build options
option(OPENXLSX_BUILD_TESTS "Build and run library tests" ON)
option(OPENXLSX_BUILD_BENCHMARKS "Build benchmark programs" OFF)
option(OPENXLSX_CREATE_DOCS "Build library documentation" OFF)
option(OPENXLSX_ENABLE_LTO "Enable Link-Time Optimization (LTO/IPO) if supported" ON)
option(OPENXLSX_ENABLE_SANITIZERS "Enable ASan and UBSan for memory/undefined behavior analysis (GCC/Clang only)" OFF)

set(OPENXLSX_LIBRARY_TYPE "STATIC" CACHE STRING "Library type (STATIC or SHARED)")

#=======================================================================================================================
# External Libraries
#=======================================================================================================================
include(FetchContent)

# Set up fast_float
FetchContent_Declare(fast_float GIT_REPOSITORY https://github.com/fastfloat/fast_float.git GIT_TAG v8.2.3)
FetchContent_MakeAvailable(fast_float)

# Set up GSL (Microsoft Guidelines Support Library)
FetchContent_Declare(GSL GIT_REPOSITORY https://github.com/microsoft/GSL.git GIT_TAG v4.1.0)
FetchContent_MakeAvailable(GSL)

# Set up fmt
FetchContent_Declare(fmt GIT_REPOSITORY https://github.com/fmtlib/fmt.git GIT_TAG 12.1.0)
FetchContent_MakeAvailable(fmt)

# Set up pugixml
set(PUGIXML_XPATH OFF CACHE BOOL "" FORCE)
FetchContent_Declare(pugixml GIT_REPOSITORY https://github.com/zeux/pugixml.git GIT_TAG v1.15)
FetchContent_MakeAvailable(pugixml)

# Set up Catch2 (if tests or benchmarks are enabled)
if(OPENXLSX_BUILD_TESTS OR OPENXLSX_BUILD_BENCHMARKS)
    FetchContent_Declare(Catch2 GIT_REPOSITORY https://github.com/catchorg/Catch2.git GIT_TAG v3.13.0)
    FetchContent_MakeAvailable(Catch2)
endif()

# Set up zlib-ng (libzip dependency)
set(ZLIB_ENABLE_TESTS OFF CACHE BOOL "" FORCE)
set(ZLIB_COMPAT ON CACHE BOOL "" FORCE)
FetchContent_Declare(zlib GIT_REPOSITORY https://github.com/zlib-ng/zlib-ng.git GIT_TAG 2.2.4)
FetchContent_MakeAvailable(zlib)

# Help libzip find zlib-ng (Override find_package(ZLIB))
# We use CACHE INTERNAL to ensure these are picked up by libzip's find_package
set(ZLIB_FOUND TRUE CACHE INTERNAL "")
set(ZLIB_INCLUDE_DIR "${zlib_SOURCE_DIR};${zlib_BINARY_DIR}" CACHE INTERNAL "")
set(ZLIB_INCLUDE_DIRS "${zlib_SOURCE_DIR};${zlib_BINARY_DIR}" CACHE INTERNAL "")
set(ZLIB_LIBRARIES zlib CACHE INTERNAL "")
set(ZLIB_LIBRARY zlib CACHE INTERNAL "")

# Create alias for standard find_package compatibility
if(TARGET zlib AND NOT TARGET ZLIB::ZLIB)
    add_library(ZLIB::ZLIB ALIAS zlib)
endif()

# Set up libzip
set(ENABLE_BZIP2 OFF CACHE BOOL "" FORCE)
set(ENABLE_LZMA OFF CACHE BOOL "" FORCE)
set(ENABLE_ZSTD OFF CACHE BOOL "" FORCE)
set(ENABLE_GNUTLS OFF CACHE BOOL "" FORCE)
set(ENABLE_OPENSSL OFF CACHE BOOL "" FORCE)
set(BUILD_TOOLS OFF CACHE BOOL "" FORCE)
set(BUILD_REGRESS OFF CACHE BOOL "" FORCE)
set(BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
set(BUILD_DOC OFF CACHE BOOL "" FORCE)
set(BUILD_OSSFUZZ OFF CACHE BOOL "" FORCE)
set(LIBZIP_DO_INSTALL OFF CACHE BOOL "" FORCE)
FetchContent_Declare(libzip GIT_REPOSITORY https://github.com/nih-at/libzip.git GIT_TAG v1.11.4)
FetchContent_MakeAvailable(libzip)

# Third-party subdirectories

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
)

target_link_libraries(OpenXLSX PUBLIC pugixml libzip::zip zlib Microsoft.GSL::GSL)
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
endif()

#=======================================================================================================================
# TESTS (Merged configuration)
#=======================================================================================================================
if(OPENXLSX_BUILD_TESTS)
    file(GLOB TEST_SOURCES Tests/testXL*.cpp)
    add_executable(OpenXLSXTests ${TEST_SOURCES})
    target_link_libraries(OpenXLSXTests PRIVATE OpenXLSX::OpenXLSX Catch2::Catch2WithMain)
    
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
            $<TARGET_FILE_DIR:OpenXLSXTests>/Tests/Fixtures/
        COMMAND ${CMAKE_COMMAND} -E copy_if_different 
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/openpyxl_generated.xlsx
            ${CMAKE_CURRENT_LIST_DIR}/Tests/Fixtures/excelize_generated.xlsx
            ${CMAKE_BINARY_DIR}/Tests/Fixtures/
    )
endif()

#=======================================================================================================================
# BENCHMARKS (Merged configuration using Catch2)
#=======================================================================================================================
if(OPENXLSX_BUILD_BENCHMARKS)
    add_executable(OpenXLSXBenchmark Benchmarks/Benchmark.cpp)
    target_link_libraries(OpenXLSXBenchmark PRIVATE OpenXLSX::OpenXLSX Catch2::Catch2)
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
    else()
        message(WARNING "Doxygen not found, documentation will not be generated.")
    endif()
endif()

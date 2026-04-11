# Quick Start: Pivot Tables & Slicers {#pivottable_tutorial}

[TOC]

## Introduction

OpenXLSX provides a powerful, enterprise-grade engine for generating dynamic Pivot Tables and interactive Slicers. The library natively performs **O(1) memory deduplication** and **Cache Sharing**, generating 100% OOXML-compliant tables that render perfectly across Microsoft Excel, Google Sheets, and WPS Office without relying on delayed calculations.

This tutorial covers:
- Preparing the source data for a pivot table.
- Utilizing the Fluent Builder API to define rows, columns, and aggregations.
- Leveraging O(1) Pivot Cache Sharing for massive dashboards.
- Reading, modifying, and deleting existing pivot tables (CRUD).
- Adding interactive Slicers.

## 1. Preparing the Source Data

A pivot table requires a contiguous range of data with column headers in the first row. We start by creating a worksheet and filling it with our raw dataset.

```cpp
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("PivotDemo.xlsx", XLForceOverwrite);
    
    // Create the raw data sheet
    auto rawWks = doc.workbook().worksheet("Sheet1");
    rawWks.setName("RawData");

    // Add Headers
    rawWks.cell("A1").value() = "Date";
    rawWks.cell("B1").value() = "Region";
    rawWks.cell("C1").value() = "SalesRep";
    rawWks.cell("D1").value() = "Product";
    rawWks.cell("E1").value() = "Revenue";

    // Add some sample data rows...
    rawWks.cell("A2").value() = "2026-Q1";
    rawWks.cell("B2").value() = "North";
    rawWks.cell("C2").value() = "Alice";
    rawWks.cell("D2").value() = "Widget A";
    rawWks.cell("E2").value() = 15000.0;
    
    // ... [Assume more data is added up to row 13] ...
```

## 2. Creating a Pivot Table (Fluent API)

To create a pivot table, use the `XLPivotTableOptions` builder. The API is designed to be highly ergonomic, enforcing mandatory parameters (Name, Source, Target) in the constructor, and allowing you to chain configurations fluently.

```cpp
    doc.workbook().addWorksheet("Summary");
    auto pivotWks = doc.workbook().worksheet("Summary");

    // 1. Build the Pivot Table configuration fluently
    auto options = XLPivotTableOptions("PivotSummary", "RawData!A1:E13", "B3")
        .addFilterField("Region")                                          // Page/Report Filter
        .addRowField("SalesRep")                                           // Row grouping
        .addColumnField("Date")                                            // Column grouping
        .addDataField("Revenue", "Total Revenue", XLPivotSubtotal::Sum, 4) // Data aggregation (4 = '#,##0.00' format)
        .setPivotTableStyle("PivotStyleMedium14")                          // Built-in Excel Theme
        .setShowRowStripes(true)                                           // Banded rows
        .setRowGrandTotals(false);                                         // Hide row grand totals

    // 2. Apply and generate the Pivot Table
    auto pt = pivotWks.addPivotTable(options);
    
    // Optional: Tell Excel to refresh calculations on the next open
    pt.setRefreshOnLoad(true); 
```

## 3. O(1) Pivot Cache Sharing (Dashboards)

When building Dashboards with multiple Pivot Tables pulling from the *exact same data source*, OpenXLSX automatically detects the matching `sourceRange` and reuses the underlying `pivotCacheDefinition`.

This **O(1) Cache Sharing** prevents memory duplication and drastically reduces the final `.xlsx` file size.

```cpp
    // Create a second pivot table looking at the SAME source data
    auto opts2 = XLPivotTableOptions("PivotByProduct", "RawData!A1:E13", "H3")
        .addRowField("Product")
        .addDataField("Revenue", "Product Revenue", XLPivotSubtotal::Sum);
        
    // OpenXLSX natively reuses the cache from "PivotSummary" since the source matches!
    pivotWks.addPivotTable(opts2);
```

## 4. Reading, Updating, and Deleting (CRUD)

OpenXLSX provides comprehensive CRUD operations to manage existing Pivot Tables, making it perfect for automated report generation where you need to keep a template's styles but swap out its data.

```cpp
    // 1. Read existing tables
    auto allPivots = pivotWks.pivotTables();
    for (auto& pivot : allPivots) {
        std::cout << "Found Pivot: " << pivot.name() << " at " << pivot.targetCell() << std::endl;
    }
    
    // 2. Update (Change Data Source Range Dynamically)
    // If you added 100 new rows to RawData, simply update the existing pivot table's bounds:
    allPivots[0].changeSourceRange("RawData!A1:E113");
    
    // 3. Delete
    // Safely remove a pivot table and its associated XML relationships
    pivotWks.deletePivotTable("PivotByProduct");
```

## 5. Adding Slicers

A Slicer is a visual filtering control that connects to a Pivot Table. Adding a slicer in OpenXLSX is extremely straightforward once the Pivot Table is created.

```cpp
    // Configure the Slicer
    XLSlicerOptions sOpts;
    sOpts.name = "Region";           // The exact name of the source data column
    sOpts.caption = "Select Region"; // The display title of the slicer window

    // Add the slicer to the worksheet, anchoring it at cell "E5"
    pivotWks.addPivotSlicer("E5", pt, "Region", sOpts);
```

## 6. Finalizing the Document

Save and close the document as usual.

```cpp
    doc.save();
    doc.close();
    
    return 0;
}
```

When you open the generated `PivotDemo.xlsx` file, the native deduplication engine ensures your pivot tables and slicers are fully populated, styled, and interactive immediately!

# Quick Start: Adding Charts {#charts_tutorial}

[TOC]

## Introduction

OpenXLSX allows you to programmatically generate and insert a wide variety of native Excel charts directly into your worksheets or as standalone chartsheets. 

These charts are fully dynamic—meaning they are linked to cell ranges in your workbook and will automatically update in Excel if the underlying data changes. 

This tutorial covers:
- Creating a basic Column or Bar chart.
- Binding data series (values, labels, and names) to cell ranges.
- Creating advanced chart types (Line, Scatter, Pie, Area).
- Advanced financial and scientific charts (Stock OHLC, 3D Surface).
- Generating dedicated Chartsheet pages.

## 1. Preparing the Data

Before adding a chart, you need data in your worksheet. A chart series typically consists of:
- **Values:** The numerical data points to plot.
- **Categories (X-Axis):** The labels for each data point (e.g., Dates, Names).
- **Series Name:** The legend name for the data.

```cpp
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("ChartDemo.xlsx", XLForceOverwrite);
    
    auto wks = doc.workbook().worksheet("Sheet1");
    wks.setName("Data");

    // Adding headers
    wks.cell("A1").value() = "Month";
    wks.cell("B1").value() = "Sales";
    wks.cell("C1").value() = "Profit";

    // Adding data
    wks.cell("A2").value() = "Jan"; wks.cell("B2").value() = 150; wks.cell("C2").value() = 45;
    wks.cell("A3").value() = "Feb"; wks.cell("B3").value() = 180; wks.cell("C3").value() = 60;
    wks.cell("A4").value() = "Mar"; wks.cell("B4").value() = 220; wks.cell("C4").value() = 90;
    wks.cell("A5").value() = "Apr"; wks.cell("B5").value() = 280; wks.cell("C5").value() = 120;
```

## 2. Adding a Basic Chart

You can insert a chart into a worksheet using `wks.addChart()`. You must specify:
1. `XLChartType` (e.g., `Column`, `Bar`, `Line`)
2. Title of the chart
3. Row and Column anchor (top-left position of the chart object)
4. Width and Height in pixels

```cpp
    // Create a Column Chart anchored at row 2, column 5 (Cell E2), 450x300 pixels
    auto colChart = wks.addChart(XLChartType::Column, "Monthly Performance", 2, 5, 450, 300);
    
    // Add the first series (Sales)
    // Signature: addSeries(ValuesRange, SeriesName, CategoryRange)
    colChart.addSeries("Data!$B$2:$B$5", "Sales", "Data!$A$2:$A$5");
    
    // Add the second series (Profit)
    colChart.addSeries("Data!$C$2:$C$5", "Profit", "Data!$A$2:$A$5");
```

## 3. Common Chart Types

OpenXLSX supports dozens of chart types via the `XLChartType` enumeration.

### Line and Area Charts

Great for showing trends over time.

```cpp
    auto lineChart = wks.addChart(XLChartType::Line, "Profit Trend", 18, 5, 450, 300);
    lineChart.addSeries("Data!$C$2:$C$5", "Profit", "Data!$A$2:$A$5");

    auto areaChart = wks.addChart(XLChartType::Area, "Sales Area", 18, 12, 450, 300);
    areaChart.addSeries("Data!$B$2:$B$5", "Sales", "Data!$A$2:$A$5");
```

### Pie and Doughnut Charts

Used for showing distributions or proportions.

```cpp
    auto pieChart = wks.addChart(XLChartType::Pie, "Sales Distribution", 34, 5, 450, 300);
    pieChart.addSeries("Data!$B$2:$B$5", "Sales", "Data!$A$2:$A$5");
```

### Stacked Charts

Used to show parts of a whole over time. Types include `BarStacked`, `ColumnStacked`, and `BarPercentStacked` (100% Stacked).

```cpp
    auto stackedChart = wks.addChart(XLChartType::BarStacked, "Stacked Analysis", 34, 12, 450, 300);
    stackedChart.addSeries("Data!$B$2:$B$5", "Sales", "Data!$A$2:$A$5");
    stackedChart.addSeries("Data!$C$2:$C$5", "Profit", "Data!$A$2:$A$5");
```

## 4. Advanced Chart Types

### Scatter Charts (XY)

Scatter charts map numerical X values to numerical Y values, rather than using categorical text labels.
Use `ScatterSmoothMarker` for curved lines with dots, or `ScatterLine` for straight lines.

```cpp
    auto scatterChart = wks.addChart(XLChartType::ScatterSmoothMarker, "XY Relationship", 50, 5, 450, 300);
    // For scatter charts, the 3rd parameter acts as the continuous X-axis values
    scatterChart.addSeries("Data!$C$2:$C$5", "Profit vs Sales", "Data!$B$2:$B$5");
```

### Bubble Charts

Bubble charts are a variation of scatter charts where the size of each marker represents a third dimension of data. Because bubble charts require three data ranges (X values, Y values, and Bubble Sizes), you must use the specialized `addBubbleSeries()` method.

```cpp
    auto bubbleChart = wks.addChart(XLChartType::Bubble, "Market Bubbles", 50, 12, 450, 300);
    // Signature: addBubbleSeries(xValuesRange, yValuesRange, sizesRange, SeriesName)
    bubbleChart.addBubbleSeries("Data!$A$2:$A$5", "Data!$B$2:$B$5", "Data!$C$2:$C$5", "Market Share");
```

### Stock (OHLC) Financial Charts

Stock charts require data to be added in a very specific order to represent Open, High, Low, and Close values.

```cpp
    // Assuming you have daily data in F (Date), G (Open), H (High), I (Low), J (Close)
    auto stockChart = wks.addChart(XLChartType::StockOHLC, "Stock Performance", 50, 12, 450, 300);
    
    // Must be added in this exact order:
    stockChart.addSeries("Data!$G$2:$G$6", "Open", "Data!$F$2:$F$6");
    stockChart.addSeries("Data!$H$2:$H$6", "High", "Data!$F$2:$F$6");
    stockChart.addSeries("Data!$I$2:$I$6", "Low", "Data!$F$2:$F$6");
    stockChart.addSeries("Data!$J$2:$J$6", "Close", "Data!$F$2:$F$6");
```

### 3D Surface Charts

Surface charts are useful for finding optimum combinations between two sets of data. As in a topographic map, colors and patterns indicate areas that are in the same range of values. Surface charts require three-dimensional data (X, Y, and Z):

- **Series Names (Y-axis / Depth):** Represent the different column categories.
- **Categories (X-axis / Width):** Represent the row labels.
- **Values (Z-axis / Height):** The actual data points plotted on the surface.

```cpp
    // Assuming a 3x3 data matrix:
    //      | Col1 | Col2 | Col3
    // Row1 |  1   |  2   |  3
    // Row2 |  4   |  5   |  6
    // Row3 |  7   |  8   |  9

    auto surfChart = wks.addChart(XLChartType::Surface3D, "3D Surface Map", 66, 5, 450, 300);
    
    // Add each column of data as a series. 
    // Signature: addSeries(ValuesRange, SeriesName, CategoryRange)
    surfChart.addSeries("Data!$B$2:$B$4", "Data!$B$1", "Data!$A$2:$A$4");
    surfChart.addSeries("Data!$C$2:$C$4", "Data!$C$1", "Data!$A$2:$A$4");
    surfChart.addSeries("Data!$D$2:$D$4", "Data!$D$1", "Data!$A$2:$A$4");
```

## 5. Dedicated Chartsheets

If you want a chart to occupy an entire tab by itself without any worksheet grid lines, you can create a Chartsheet.

```cpp
    // Create a new chartsheet
    doc.workbook().addChartsheet("Dashboard Chart");
    auto cs = doc.workbook().chartsheet("Dashboard Chart");
    
    // Add the chart (No row/col anchoring needed as it fills the sheet)
    auto csChart = cs.addChart(XLChartType::Column3D, "3D Performance Overview");
    csChart.addSeries("Data!$B$2:$B$5", "Sales", "Data!$A$2:$A$5");
    csChart.addSeries("Data!$C$2:$C$5", "Profit", "Data!$A$2:$A$5");

    doc.save();
    doc.close();
    
    return 0;
}
```

This generates a professional workbook where the data drives multiple visualizations across embedded worksheets and full-page chart tabs.

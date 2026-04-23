# OpenXLSX Missing Excel Functions Checklist

This document serves as a comprehensive roadmap for expanding the `XLFormulaEngine` in the OpenXLSX C++ library. Based on a cross-reference between Microsoft Excel's core formula categories and the currently registered 114 functions in OpenXLSX, the following high-priority functions are missing and categorized by implementation difficulty and business value.

## 1. Lookup & Reference (查找和引用)
This category is the most requested by data engineers. Currently, OpenXLSX lacks dynamic array and spatial reference capabilities.

### Legacy High Priority
- [ ] `OFFSET` (Returns a reference offset from a starting cell)
- [ ] `INDIRECT` (Returns a reference specified by a text string - **Hard**)
- [ ] `ROW` / `COLUMN` (Returns the row/column number of a reference)
- [ ] `ROWS` / `COLUMNS` (Returns the number of rows/cols in a reference)
- [ ] `ADDRESS` (Creates a cell address as text)
- [ ] `CHOOSE` (Chooses a value from a list of values)
- [ ] `GETPIVOTDATA` (Returns data stored in a PivotTable report)

### Modern Dynamic Arrays (O365)
- [ ] `FILTER` (Filters a range of data based on criteria)
- [ ] `UNIQUE` (Returns a list of unique values)
- [ ] `SORT` / `SORTBY` (Sorts the contents of a range or array)
- [ ] `TRANSPOSE` (Returns the transpose of an array)
- [ ] `VSTACK` / `HSTACK` (Appends arrays vertically/horizontally)

## 2. Math & Trigonometry (数学和三角函数)
Basic trigonometry and math are implemented, but advanced statistical aggregations and matrix operations are missing.

### High Value Aggregation
- [ ] `SUBTOTAL` (Returns a subtotal in a list or database - **Must Have**)
- [ ] `AGGREGATE` (Returns an aggregate in a list or database, ignoring errors/hidden rows)
- [x] `SUMSQ` / `SUMX2MY2` / `SUMX2PY2` / `SUMXMY2` (Sum of squares calculations)
- [x] `MROUND` / `CEILING.MATH` / `FLOOR.MATH`
- [x] `TRUNC` (Advanced rounding)

### Matrix & Linear Algebra
- [ ] `MMULT` (Matrix multiplication)
- [ ] `MINVERSE` (Matrix inverse)
- [ ] `MDETERM` (Matrix determinant)

## 3. Date & Time (日期和时间)
While basic extraction (YEAR, MONTH, DAY) is complete, business logic functions for calculating workdays and exact ages are missing.

### Business Logic
- [ ] `DATEDIF` (Calculates the number of days, months, or years between two dates)
- [x] `DAYS360` (Calculates days between two dates based on a 360-day year)
- [ ] `YEARFRAC` (Returns the year fraction representing the number of whole days)
- [x] `WEEKNUM` / `ISOWEEKNUM` (Converts a serial number to a number representing where the week falls numerically)
- [ ] `WORKDAY.INTL` / `NETWORKDAYS.INTL` (Calculates workdays with custom weekend parameters)

## 4. Text (文本)
String manipulation is mostly complete, but formatting and modern splitting functions are absent.

### Formatting & Encoding
- [ ] `TEXT` (Formats a number and converts it to text - **Very Hard**, requires custom format parser)
- [x] `CHAR` / `UNICHAR` (Returns the character specified by the code number)
- [x] `CODE` / `UNICODE` (Returns a numeric code for the first character)

### Modern String Manipulation (O365)
- [ ] `TEXTSPLIT` (Splits text into rows or columns using delimiters)
- [ ] `TEXTBEFORE` / `TEXTAFTER` (Returns text that occurs before/after a given character)

## 5. Logical & Information (逻辑与信息)
The core logic (IF, AND, OR, SWITCH) is solidly implemented. Missing are constants, environment variables, and modern closures.

### Constants & Error Handling
- [x] `TRUE` / `FALSE` (Returns the logical value TRUE/FALSE)
- [x] `ISERR` / `ISEVEN` / `ISODD`
- [ ] `ISREF` (Various type-checking functions)
- [ ] `CELL` / `INFO` (Returns information about formatting, location, or contents of a cell)

### Modern Programming (O365)
- [ ] `LET` (Assigns names to calculation results)
- [ ] `LAMBDA` (Creates custom, reusable functions)

## 6. Financial & Statistical (财务与进阶统计)
Basic time-value-of-money (PMT, FV, PV, NPV) is present. Real-world financial modeling requires deeper algorithms.

### Corporate Finance
- [ ] `IRR` / `MIRR` / `XIRR` (Internal rate of return - Requires Newton-Raphson iteration)
- [ ] `XNPV` (Net present value for a schedule of cash flows that is not necessarily periodic)
- [x] `RATE` / `NPER` (Returns the interest rate per period / number of periods)

### Depreciation
- [x] `DB` (Fixed-declining balance depreciation)
- [x] `DDB` (Double-declining balance depreciation)
- [x] `SLN` (Straight-line depreciation)
- [x] `SYD` (Sum-of-years' digits depreciation)

### Advanced Statistics
- [ ] `PERCENTILE.INC` / `PERCENTILE.EXC`
- [x] `QUARTILE.INC` / `QUARTILE.EXC`
- [x] `STDEV.P` / `VAR.P` (Population standard deviation and variance)
- [ ] `CORREL` / `COVARIANCE.P` / `COVARIANCE.S`

## 7. Additional Statistical Functions (统计函数补充)
Identified from Excel's Statistical category (excluding those already implemented or listed above).
- [x] `AVEDEV` / `DEVSQ`
- [x] `AVERAGEA`
- [x] `AVERAGEIFS`
- [ ] `BETA.DIST` / `BETA.INV`
- [ ] `BINOM.DIST` / `BINOM.DIST.RANGE` / `BINOM.INV`
- [ ] `CHISQ.DIST` / `CHISQ.DIST.RT` / `CHISQ.INV` / `CHISQ.INV.RT` / `CHISQ.TEST`
- [ ] `CONFIDENCE.NORM` / `CONFIDENCE.T`
- [ ] `EXPON.DIST`
- [ ] `F.DIST` / `F.DIST.RT` / `F.INV` / `F.INV.RT` / `F.TEST`
- [x] `FISHER` / `FISHERINV`
- [ ] `FORECAST.ETS` / `FORECAST.ETS.CONFINT` / `FORECAST.ETS.SEASONALITY` / `FORECAST.ETS.STAT` / `FORECAST.LINEAR`
- [ ] `FREQUENCY`
- [ ] `GAMMA` / `GAMMA.DIST` / `GAMMA.INV`
- [ ] `NORM.INV` / `NORM.S.DIST` / `NORM.S.INV`
- [x] `PEARSON` / `RSQ`
- [ ] `PERCENTRANK.EXC` / `PERCENTRANK.INC`
- [x] `PERMUT` / `PERMUTATIONA`
- [ ] `PHI`
- [ ] `POISSON.DIST`
- [ ] `PROB`
- [ ] `RANK.AVG`
- [ ] `SKEW` / `SKEW.P`
- [ ] `SLOPE` / `STEYX` / `TREND`
- [x] `STANDARDIZE`
- [x] `STDEVA` / `STDEVPA`
- [ ] `T.DIST` / `T.DIST.2T` / `T.DIST.RT` / `T.INV` / `T.INV.2T` / `T.TEST`
- [ ] `TRIMMEAN`
- [x] `VARA` / `VARPA`
- [ ] `WEIBULL.DIST`
- [ ] `Z.TEST`

## 8. Compatibility Functions (兼容性函数)
Legacy functions retained by Excel for backward compatibility.
- [ ] `BETADIST` / `BETAINV`
- [ ] `BINOMDIST` / `CRITBINOM`
- [ ] `CHIDIST` / `CHIINV` / `CHITEST`
- [ ] `CONFIDENCE`
- [ ] `COVAR`
- [ ] `EXPONDIST`
- [ ] `FDIST` / `FINV` / `FTEST`
- [ ] `FORECAST`
- [ ] `GAMMADIST` / `GAMMAINV`
- [ ] `HYPGEOMDIST`
- [ ] `LOGINV` / `LOGNORMDIST`
- [ ] `MODE`
- [ ] `NEGBINOMDIST`
- [ ] `NORMDIST` / `NORMINV` / `NORMSDIST` / `NORMSINV`
- [ ] `PERCENTILE` / `PERCENTRANK`
- [ ] `POISSON`
- [ ] `QUARTILE`
- [x] `STDEVP`
- [ ] `TDIST` / `TINV` / `TTEST`
- [x] `VARP`
- [ ] `WEIBULL`
- [ ] `ZTEST`


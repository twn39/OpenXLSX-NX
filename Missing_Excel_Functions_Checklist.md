# OpenXLSX Missing Excel Functions Checklist

This document tracks functions **not yet implemented** in the `XLFormulaEngine`. All entries here are pending work items.

---

## 1. Lookup & Reference (查找和引用)

### Legacy
- [ ] `OFFSET` (Returns a reference offset from a starting cell)
- [ ] `INDIRECT` (Returns a reference specified by a text string — **Hard**)
- [ ] `ROWS` / `COLUMNS` (Returns the number of rows/cols in a reference)
- [ ] `ADDRESS` (Creates a cell address as text)
- [ ] `GETPIVOTDATA` (Returns data stored in a PivotTable report)

### Modern Dynamic Arrays (O365)
- [ ] `FILTER` (Filters a range of data based on criteria)
- [ ] `UNIQUE` (Returns a list of unique values)
- [ ] `SORT` / `SORTBY` (Sorts the contents of a range or array)
- [ ] `TRANSPOSE` (Returns the transpose of an array)
- [ ] `VSTACK` / `HSTACK` (Appends arrays vertically/horizontally)

---

## 2. Math & Trigonometry (数学和三角函数)

- [ ] `AGGREGATE` (Returns an aggregate in a list or database, ignoring errors/hidden rows)
- [ ] `MMULT` (Matrix multiplication)
- [ ] `MINVERSE` (Matrix inverse)
- [ ] `MDETERM` (Matrix determinant)

---

## 3. Date & Time (日期和时间)

- [ ] `YEARFRAC` (Returns the year fraction representing the number of whole days)
- [ ] `WORKDAY.INTL` / `NETWORKDAYS.INTL` (Calculates workdays with custom weekend parameters)

---

## 4. Text (文本)

### Modern String Manipulation (O365)
- [ ] `TEXTSPLIT` (Splits text into rows or columns using delimiters)
- [ ] `TEXTBEFORE` / `TEXTAFTER` (Returns text that occurs before/after a given character)

---

## 5. Logical & Information (逻辑与信息)

- [ ] `ISREF` (Returns TRUE if the value is a reference)
- [ ] `CELL` / `INFO` (Returns information about formatting, location, or contents of a cell)
- [ ] `LET` (Assigns names to calculation results — O365)
- [ ] `LAMBDA` (Creates custom, reusable functions — O365)

---

## 6. Financial (财务)

- [ ] `XIRR` (IRR for non-periodic cash flows)
- [ ] `XNPV` (Net present value for non-periodic cash flows)

---

## 7. Statistical (统计)

- [ ] `BETA.DIST` / `BETA.INV`
- [ ] `BINOM.DIST.RANGE` / `BINOM.INV`
- [ ] `CHISQ.TEST`
- [ ] `CONFIDENCE.NORM` / `CONFIDENCE.T`
- [ ] `F.DIST` / `F.DIST.RT` / `F.INV` / `F.INV.RT` / `F.TEST`
- [ ] `FORECAST.ETS` / `FORECAST.ETS.CONFINT` / `FORECAST.ETS.SEASONALITY` / `FORECAST.ETS.STAT`
- [ ] `FREQUENCY`
- [ ] `GAMMA` / `GAMMA.DIST` / `GAMMA.INV`
- [ ] `PERCENTRANK.EXC` / `PERCENTRANK.INC`
- [ ] `PHI`
- [ ] `PROB`
- [ ] `RANK.AVG`
- [ ] `SKEW.P` (Population skewness)
- [ ] `SLOPE` / `STEYX` / `TREND`
- [ ] `WEIBULL.DIST`
- [ ] `Z.TEST`

---

## 8. Compatibility Functions (兼容性函数)

Legacy functions not yet mapped to modern equivalents.

- [ ] `BETADIST` / `BETAINV`
- [ ] `CHITEST`
- [ ] `CONFIDENCE`
- [ ] `COVAR` (`COVARIANCE.P` is implemented; this legacy alias is not)
- [ ] `CRITBINOM`
- [ ] `FDIST` / `FINV` / `FTEST`
- [ ] `GAMMADIST` / `GAMMAINV`
- [ ] `HYPGEOMDIST`
- [ ] `LOGINV` / `LOGNORMDIST`
- [ ] `NEGBINOMDIST`
- [ ] `PERCENTRANK`
- [ ] `TTEST` / `ZTEST`
- [ ] `WEIBULL`

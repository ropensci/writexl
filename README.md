# writexl

<!-- badges: start -->
[![Project Status: Active – The project has reached a stable, usable state and is being actively developed.](http://www.repostatus.org/badges/latest/active.svg)](http://www.repostatus.org/#active)
[![CRAN_Status_Badge](http://www.r-pkg.org/badges/version/writexl)](http://cran.r-project.org/package=writexl)
[![CRAN RStudio mirror downloads](http://cranlogs.r-pkg.org/badges/writexl)](http://cran.r-project.org/web/packages/writexl/index.html)
[![badge](https://ropensci.r-universe.dev/badges/writexl)](https://ropensci.r-universe.dev)
[![R-CMD-check](https://github.com/ropensci/writexl/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/ropensci/writexl/actions/workflows/R-CMD-check.yaml)
[![Codecov test coverage](https://codecov.io/gh/ropensci/writexl/graph/badge.svg)](https://app.codecov.io/gh/ropensci/writexl)
<!-- badges: end -->

> Portable, light-weight data frame to xlsx exporter based on libxlsxwriter.  No Java or Excel required.

Wraps the [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter) library to create files
in Microsoft Excel 'xlsx' format.

## Installation

```r
install.packages("writexl")
```

## Getting started

```r
library(writexl)
library(readxl)
tmp <- write_xlsx(iris)
read_xlsx(tmp)
```
```
# A tibble: 150 × 5
   Sepal.Length Sepal.Width Petal.Length Petal.Width Species
          <dbl>       <dbl>        <dbl>       <dbl> <chr>
 1          5.1         3.5          1.4         0.2 setosa
 2          4.9         3            1.4         0.2 setosa
 3          4.7         3.2          1.3         0.2 setosa
 4          4.6         3.1          1.5         0.2 setosa
 5          5           3.6          1.4         0.2 setosa
# ℹ 145 more rows
```

A named list writes one sheet per element:

```r
write_xlsx(list(Flowers = iris, Cars = mtcars))
```

That is the whole of the common case. Everything below is optional.

## What else it can do

Each of these has a vignette; none is needed to write a workbook.

| | |
|---|---|
| **[Formatting cells](https://docs.ropensci.org/writexl/articles/formatting.html)** | Fonts, fills, borders, number formats and alignment, built from group constructors that combine with `+`. The same format object styles a cell, a column, a chart series or a conditional rule. Conditional formatting, colour scales, data bars and icon sets. |
| **[Worksheets and workbooks](https://docs.ropensci.org/writexl/articles/sheets.html)** | Column widths and row heights in characters or pixels, frozen and split panes, tab colours and visibility, outlines, protection, page setup and printing, document properties, and the row-streaming mode for large workbooks. |
| **[Charts and images](https://docs.ropensci.org/writexl/articles/charts.html)** | All 22 chart types Excel offers, with axes, markers, data labels, trendlines, error bars, legends and data tables; chartsheets; and pictures placed on a sheet, inside a cell, in a header or tiled behind the grid. |
| **[Formulas, tables and the rest](https://docs.ropensci.org/writexl/articles/more.html)** | Formulas including array and dynamic array formulas, hyperlinks, comments, rich strings, mixed-type columns, data validation, autofilters that actually hide the rows, worksheet tables and merged cells. |
| **[Everything at once](https://docs.ropensci.org/writexl/articles/showcase.html)** | One runnable script that exercises the whole package into two workbooks — the quickest way to see what it does without reading the other four. |

A taste of how they fit together:

```r
sheet <- xl_sheet(
  sales,
  cols   = xl_col_spec("revenue", format = xl_num_format("$#,##0")),
  freeze = "A2",
  chart  = xl_chart("column", xl_chart_series(values = list(cols = "revenue")))
)
write_xlsx(list(Sales = sheet), "report.xlsx")
```

## Round-tripping

Every column type survives a trip through `readxl` unchanged, with one
exception that is deliberate:

```r
library(nycflights13)
out <- readxl::read_xlsx(writexl::write_xlsx(flights))
keep <- setdiff(names(flights), "time_hour")
all.equal(out[keep], flights[keep])
## TRUE
```

`time_hour` is a `POSIXct` in `America/New_York`, and **Excel has no concept of
a time zone**. Rather than silently converting everything to UTC, writexl
decides once per workbook: when every datetime shares a zone it writes the
local wall-clock reading and drops the " UTC" suffix from the default datetime
format, so nothing is mislabelled; when the zones differ it converts to UTC and
warns. Either way the zone itself is not in the file for `readxl` to hand back.

## Performance

Writing `nycflights13::flights` — 336,776 rows by 19 columns:

```r
system.time(writexl::write_xlsx(flights, tmp <- tempfile(fileext = ".xlsx")))
##    user  system elapsed
##    7.4     0.3     7.7
file.size(tmp)
## 29157011   (27.8 MB)
```

For a workbook large enough to matter, `constant_memory = TRUE` streams each
row to disk instead of holding the sheet in memory; left at its default writexl
decides from the estimated cost. See
[Worksheets and workbooks](https://docs.ropensci.org/writexl/articles/sheets.html).

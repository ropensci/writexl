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
| **[Formatting cells](https://docs.ropensci.org/writexl/articles/b-formatting.html)** | Fonts, fills, borders, number formats and alignment, built from group constructors that combine with `+`. The same format object styles a cell, a column, a chart series or a conditional rule. Conditional formatting, colour scales, data bars and icon sets. |
| **[Worksheets and workbooks](https://docs.ropensci.org/writexl/articles/c-worksheets-workbooks.html)** | Column widths and row heights in characters or pixels, frozen and split panes, tab colours and visibility, outlines, protection, page setup and printing, document properties, and the row-streaming mode for large workbooks. |
| **[Charts and images](https://docs.ropensci.org/writexl/articles/d-charts-images.html)** | All 22 chart types Excel offers, with axes, markers, data labels, trendlines, error bars, legends and data tables; chartsheets; and pictures placed on a sheet, inside a cell, in a header or tiled behind the grid. |
| **[Formulas, tables and the rest](https://docs.ropensci.org/writexl/articles/e-formulas-and-more.html)** | Formulas including array and dynamic array formulas, hyperlinks, comments, rich strings, mixed-type columns, data validation, autofilters that actually hide the rows, worksheet tables and merged cells. |
| **[Everything at once](https://docs.ropensci.org/writexl/articles/f-everything-at-once.html)** | One runnable script that exercises the whole package into two workbooks — the quickest way to see what it does without reading the other four. |

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

Writing `nycflights13::flights` — 336,776 rows by 19 columns — is faster than
the `openxlsx2` implementation, for files of the same size:

```r
library(nycflights13)
bench <- function(f) median(replicate(2, {
  p <- tempfile(fileext = ".xlsx")
  on.exit(unlink(p))
  system.time(f(p))[["elapsed"]]
}))

bench(function(p) writexl::write_xlsx(flights, p))
## 12.3
bench(function(p) openxlsx2::write_xlsx(flights, p))
## 19.6
```

The output files are within a rounding error of each other:

```r
file.size(writexl::write_xlsx(flights, tempfile(fileext = ".xlsx")))
## 29132216   (27.8 MB)

p <- tempfile(fileext = ".xlsx"); openxlsx2::write_xlsx(flights, p); file.size(p)
## 29297097   (27.9 MB)
```

Measured on R 4.6.1, writexl 1.5.4.9000, openxlsx2 1.28.

### Memory

`constant_memory = TRUE` streams each row to disk instead of holding the whole
sheet, which is the difference between a workbook that fits in memory and one
that does not. Peak resident memory of the whole R process, writing the same
336,776 rows:

| | peak memory | file | time |
|---|---|---|---|
| `constant_memory = FALSE` | 894 MB | 26.8 MB | 7.4 s |
| `constant_memory = TRUE` | **208 MB** | 27.8 MB | 7.7 s |

A quarter of the memory, for a file 3.5% larger and about the same time.
Streaming produces slightly bigger files because a shared-string table cannot
be built for rows already flushed to disk.

Left at its default, writexl chooses per workbook: streaming is switched on
only when the estimated saving is worth it, and switched off automatically for
the features that cannot be written while streaming — worksheet tables,
multi-cell array formulas, merged ranges and embedded images. See
[Worksheets and workbooks](https://docs.ropensci.org/writexl/articles/c-worksheets-workbooks.html).

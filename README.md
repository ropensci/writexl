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

Currently the package only has `write_xlsx()` to export a data frame to xlsx. 

```r
library(writexl)
library(readxl)
tmp <- writexl::write_xlsx(iris)
readxl::read_xlsx(tmp)
```
```
# A tibble: 150 x 5
   Sepal.Length Sepal.Width Petal.Length Petal.Width Species
          <dbl>       <dbl>        <dbl>       <dbl>   <chr>
 1          5.1         3.5          1.4         0.2  setosa
 2          4.9         3.0          1.4         0.2  setosa
 3          4.7         3.2          1.3         0.2  setosa
 4          4.6         3.1          1.5         0.2  setosa
 5          5.0         3.6          1.4         0.2  setosa
 6          5.4         3.9          1.7         0.4  setosa
 7          4.6         3.4          1.4         0.3  setosa
 8          5.0         3.4          1.5         0.2  setosa
 9          4.4         2.9          1.4         0.2  setosa
10          4.9         3.1          1.5         0.1  setosa
# ... with 140 more rows
```

Most data types should roundtrip with `readxl`:

```r
library(nycflights13)
out <- readxl::read_xlsx(writexl::write_xlsx(flights))
all.equal(out, flights)
## TRUE
```

Performance is a bit better than `openxlsx` implementation:

```r
library(microbenchmark)
library(nycflights13)
microbenchmark(
  writexl = writexl::write_xlsx(flights, tempfile()),
  openxlsx = openxlsx::write.xlsx(flights, tempfile()),
  times = 5
)
## Unit: seconds
##      expr       min        lq      mean    median        uq       max neval
##   writexl  8.884712  8.904431  9.103419  8.965643  9.041565  9.720743     5
##  openxlsx 17.166818 18.072527 19.171003 18.669805 18.756661 23.189206     5
```

Also the output xlsx files are smaller:

```r
writexl::write_xlsx(flights, tmp1 <- tempfile())
file.info(tmp1)$size
## 29157282
```

```r
openxlsx::write.xlsx(flights, tmp2 <- tempfile())
file.info(tmp2)$size
## 35962067
```

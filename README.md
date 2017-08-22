# writexl

[![Build Status](https://travis-ci.org/ropensci/writexl.svg?branch=master)](https://travis-ci.org/ropensci/writexl)
[![AppVeyor Build Status](https://ci.appveyor.com/api/projects/status/github/ropensci/writexl?branch=master&svg=true)](https://ci.appveyor.com/project/jeroen/writexl)
[![Coverage Status](https://codecov.io/github/ropensci/writexl/coverage.svg?branch=master)](https://codecov.io/github/ropensci/writexl?branch=master)
[![CRAN_Status_Badge](http://www.r-pkg.org/badges/version/writexl)](http://cran.r-project.org/package=writexl)
[![CRAN RStudio mirror downloads](http://cranlogs.r-pkg.org/badges/writexl)](http://cran.r-project.org/web/packages/writexl/index.html)

> Portable, light-weight data frame to xlsx exporter based on libxlsxwriter.  No Java or Excel required.

Wraps the [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter) library to create files
in Microsoft Excel 'xslx' format.

## Installation

```r
devtools::install_github("ropensci/writexl")
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
```
```
TRUE
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
```
```
Unit: seconds
     expr      min       lq     mean   median       uq      max neval
  writexl 10.94430 11.01603 11.45755 11.41462 11.50409 12.40871     5
 openxlsx 18.20038 19.67410 19.53756 19.74198 19.75242 20.31890     5
```

Also the output xlsx files are smaller:

```r
writexl::write_xlsx(flights, tmp1 <- tempfile())
file.info(tmp1)$size
```
```
28229493
```
```r
openxlsx::write.xlsx(flights, tmp2 <- tempfile())
file.info(tmp2)$size
```
```
35962067
```

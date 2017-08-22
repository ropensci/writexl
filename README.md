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

## Hello World

Currently the package only has `write_xlsx()` to export a data frame to xlsx:

```r
library(writexl)
library(readxl)
tmp <- writexl::write_xlsx(iris)
readxl::read_xlsx(tmp)
```

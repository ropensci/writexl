#' Excel Types
#'
#' Create special column types to write to a spreadsheet
#'
#' @family writexl
#' @param x character vector to be interpreted as formula
#' @export
xl_formula <- function(x){
  if(is.factor(x))
    x <- as.character(x)
  stopifnot(is.character(x))
  structure(x, class = c('xl_formula', 'xl_object'))
}

#' @export
print.xl_formula <- function(x, max = 10, ...){
  cat(" [:xl_formula:]\n")
  if(length(x) > max)
    x <- c(x[1:max], "...", sprintf("(total: %s)", length(x)))
  cat(x, sep = "\n")
}

#' @export
rep.xl_object <- function(x, ...){
  structure(rep(unclass(x), ...), class = class(x))
}

#' @export
`[.xl_object` <- function(x, ...){
  structure(`[`(unclass(x), ...), class = class(x))
}

#' @export
`[[.xl_object` <- function(x, ...){
  structure(`[[`(unclass(x), ...), class = class(x))
}

#' @export
c.xl_object <- function(x, ...){
  structure(c(unclass(x), ...), class = class(x))
}

#' @export
as.data.frame.xl_object <- function(x, ..., stringsAsFactors = FALSE){
  as.data.frame.character(x, ..., stringsAsFactors = FALSE)
}

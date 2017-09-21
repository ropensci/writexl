#' Excel Types
#'
#' Create special column types to write to a spreadsheet
#'
#' @family writexl
#' @param x character vector to be interpreted as formula
#' @export
#' @rdname xl_formula
xl_formula <- function(x){
  if(is.factor(x))
    x <- as.character(x)
  stopifnot(is.character(x))
  if(!all(grepl("^=",x)))
    stop("Formulas must start with '='")
  structure(x, class = c('xl_formula', 'xl_object'))
}

#' @rdname xl_formula
#' @export
#' @param url character vector of URLs
#' @param name character vector of friendly names
xl_hyperlink <- function(url, name = NULL){
  if(is.factor(url))
    url <- as.character(url)
  stopifnot(is.character(url))
  hyperlink <- dubquote(url)
  if(length(name)){
    hyperlink <- paste(hyperlink, dubquote(name), sep = ",")
  }
  out <- xl_formula(sprintf("=HYPERLINK(%s)", hyperlink))
  structure(out, class = c('xl_hyperlink', 'xl_formula', 'xl_object'))
}

#' @export
print.xl_formula <- function(x, max = 10, ...){
  cat(sprintf(" [:%s:]\n", class(x)[1]))
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

dubquote <- function(x){
  paste0('"', x, '"')
}

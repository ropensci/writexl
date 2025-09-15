#' Excel Types
#'
#' Create special column types to write to a spreadsheet. See `dubquote()` if
#' you need to escape quoted strings in the formula for potential double quotes
#' that may be included.
#'
#' @family writexl
#' @param x character vector to be interpreted as formula
#' @export
#' @rdname xl_formula
#' @seealso [dubquote()]
#' @examples
#' df <- data.frame(
#'   name = c("UCLA", "Berkeley", "Jeroen"),
#'   founded = c(1919, 1868, 2030),
#'   website = xl_hyperlink(c("http://www.ucla.edu", "http://www.berkeley.edu", NA), "homepage")
#' )
#' df$age <- xl_formula('=(YEAR(TODAY()) - INDIRECT("B" & ROW()))')
#' write_xlsx(df, 'universities.xlsx')
#'
#' # cleanup
#' unlink('universities.xlsx')
xl_formula <- function(x){
  if(is.factor(x))
    x <- as.character(x)
  stopifnot(is.character(x))
  if(!all(grepl("^=",x) | is.na(x)))
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
  out[is.na(url)] <- NA
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

#' Double-quote values with required escaping
#'
#' Excel requires double quotes (`"`) within formula to be escaped by doubling
#' them (`""`). This will automatically detect when that is required and do the escaping for you.
#'
#' @param x A character string (or vector of character strings) to enclose in
#'   double quotes and escape if necessary
#' @returns `x` with double quotes around it and with double quotes escaped, if
#'   necessary
#' @export
dubquote <- function(x){
  # If there are double-quotes, escape the double-quotes
  paste0('"', escape_dubquote(x), '"')
}

# Escape double-quotes to ensure that excel files load correctly (#89)
escape_dubquote <- function(x) {
  gsub(x = x, pattern = '"', replacement = '""')
}

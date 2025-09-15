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

#' Double-quote and split values for formula with required escaping
#'
#' Excel requires double quotes (`"`) within formula to be escaped by doubling
#' them (`""`). This will automatically detect when that is required and do the
#' escaping for you.
#'
#' Additionally, Excel allows a maximum of 255 characters in a formula string.
#' This will break to a maximum of 255 characters (accounting for double quotes,
#' too).
#'
#' @param x A character string (or vector of character strings) to enclose in
#'   double quotes and escape if necessary
#' @returns `x` with double quotes around it and with double quotes escaped, if
#'   necessary
#' @export
dubquote <- function(x){
  if (any(!is.na(x) & nchar(x) > 32767)) {
    # Based on
    # https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3,
    # downloaded on 2025-09-15
    stop("Excel cannot handle > 32767 characters in a single cell")
    # This is not a perfect check since double quotes, and other modifications
    # will be added
  }
  # Count double quotes (") twice since they will need to be escaped using ""
  nchar_with_dubquotes <- nchar(x) + nchar(gsub(x = x, pattern = '[^"]', replacement = ""))
  ret <- x
  mask_short <- is.na(x) | (nchar_with_dubquotes < 256)
  # Short strings can be simply double-quoted
  ret[mask_short] <- dubquote_simple(escape_dubquote(x[mask_short]))
  if (any(!mask_short)) {
    # Long strings require the use of CONCATENATE()
    ret[!mask_short] <- sapply(X = as.list(x[!mask_short]), FUN = concatenate_longstring)
  }
  ret
}

dubquote_simple <- function(x) {
  paste0('"', x, '"')
}

# Escape double-quotes to ensure that excel files load correctly (#89)
escape_dubquote <- function(x) {
  gsub(x = x, pattern = '"', replacement = '""')
}

# Break up a long string into concatenated values, accounting for double quotes
concatenate_longstring <- function(x) {
  # The break location will depend on the locations of double quotes
  dubquote_locations <- setdiff(gregexpr(text = x, pattern = '"')[[1]], -1)
  # Allow a maximum of 255 characters in a substring, accounting for
  # double-quote locations
  current_split <- 0L
  ret <- character()
  while (current_split < nchar(x)) {
    last_split <- current_split
    current_split <- last_split + 255L - min(128L, sum(dubquote_locations > last_split & dubquote_locations <= (last_split + 255L)))
    ret[length(ret) + 1] <- dubquote_simple(escape_dubquote(substr(x, last_split + 1, current_split)))
  }
  sprintf("CONCATENATE(%s)", paste(ret, collapse = ","))
}

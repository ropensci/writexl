#' Excel Types
#'
#' Create special column types to write to a spreadsheet.
#'
#' @description
#' * `xl_formula(x)` — wraps a character vector of Excel formulas (each must
#'   start with `"="`). The formulas are written to the xlsx file as-is and
#'   are recalculated by Excel on open.
#'
#' * `xl_hyperlink(url, name)` — convenience wrapper that builds an Excel
#'   `=HYPERLINK(url, name)` **formula** for each element.  Because the
#'   hyperlink is stored as a formula, it is readable by
#'   [readxl::read_xlsx()], which returns the formula text.  Display text is
#'   controlled by the `name` argument.
#'
#' * `xl_hyperlink_cell(url, value)` — creates a **native cell-level
#'   hyperlink** using `worksheet_write_url_opt()` from libxlsxwriter.  The
#'   URL is stored as metadata attached to the cell, not in the formula bar.
#'   An optional `value` argument provides the display text shown in the cell.
#'   A tooltip and further options can be set by passing a named list to
#'   `xl_cell_general()` directly.
#'   **Note:** [readxl::read_xlsx()] cannot read cell-level hyperlinks and
#'   returns `NA` for those cells.  Use `xl_hyperlink()` instead when
#'   round-tripping through readxl is required.
#'
#' @family writexl
#' @param x character vector to be interpreted as formula
#' @export
#' @rdname xl_formula
#' @examples
#' df <- data.frame(
#'   name = c("UCLA", "Berkeley", "Jeroen"),
#'   founded = c(1919, 1868, 2030),
#'   website = xl_hyperlink(c("http://www.ucla.edu", "http://www.berkeley.edu", NA), "homepage")
#' )
#' df$age <- xl_formula('=(YEAR(TODAY()) - INDIRECT("B" & ROW()))')
#' write_xlsx(df, 'universities.xlsx')
#'
#' # xl_hyperlink_cell() stores the URL as native cell metadata.
#' # readxl cannot read these cells, but they display cleanly in Excel.
#' df2 <- data.frame(
#'   name = c("UCLA", "Berkeley"),
#'   website = xl_hyperlink_cell(c("http://www.ucla.edu", "http://www.berkeley.edu"),
#'                                value = "homepage")
#' )
#' write_xlsx(df2, 'universities2.xlsx')
#'
#' # cleanup
#' unlink(c('universities.xlsx', 'universities2.xlsx'))
xl_formula <- function(x){
  if(is.factor(x))
    x <- as.character(x)
  stopifnot(is.character(x))
  if(!all(grepl("^=",x) | is.na(x)))
    stop("Formulas must start with '='")
  xl_cell_general(formula = x)
}

#' @rdname xl_formula
#' @export
#' @param url character vector of URLs.  Use `NA` to produce a blank cell.
#' @param name character vector of friendly display names shown in the cell
#'   instead of the raw URL.  When `NULL`, the URL is shown.  Ignored for
#'   `NA` URLs.
xl_hyperlink <- function(url, name = NULL){
  if(is.factor(url))
    url <- as.character(url)
  stopifnot(is.character(url))
  fmlas <- if(!is.null(name)){
    paste0("=HYPERLINK(", dubquote(url), ",", dubquote(name), ")")
  } else {
    paste0("=HYPERLINK(", dubquote(url), ")")
  }
  fmlas[is.na(url)] <- NA_character_
  xl_formula(fmlas)
}

#' @rdname xl_formula
#' @export
#' @param value character vector (or `NULL`) of display text shown in the
#'   cell.  For `xl_hyperlink_cell()`, when `NULL` the raw URL is shown.
#'   Recycled to the length of `url`.  Automatically set to `NA` for cells
#'   whose URL is `NA`.
xl_hyperlink_cell <- function(url, value = NULL){
  if(is.factor(url))
    url <- as.character(url)
  stopifnot(is.character(url))
  if(is.null(value)){
    xl_cell_general(hyperlink = url)
  } else {
    value_arg <- rep_len(as.character(value), length(url))
    value_arg[is.na(url)] <- NA_character_
    xl_cell_general(value = value_arg, hyperlink = url)
  }
}

# Wrap x in Excel double-quotes, doubling any internal double-quote characters.
dubquote <- function(x){
  paste0('"', gsub('"', '""', x, fixed = TRUE), '"')
}

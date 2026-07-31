# Coerce a character-ish argument, accepting a factor, and fail with a message
# that names the argument rather than stopifnot()'s expression dump.  Shared by
# the three cell shorthands below, which all take a character vector first.
.as_character_arg <- function(x, arg) {
  if (is.factor(x)) x <- as.character(x)
  if (!is.character(x))
    stop(sprintf("`%s` must be a character vector (got %s)", arg,
                 paste(class(x), collapse = "/")), call. = FALSE)
  x
}

# The one string an object will show in the workbook.  Anything classed is put
# through as.character(), so a cell built for a sheet -- xl_cell_general(value
# = "Total") -- can be reused wherever only its displayed value is wanted.  A
# rich string is refused rather than flattened: these callers write a single
# run, so its per-run fonts would disappear without a word.
.as_display_string <- function(x, arg, null_ok = TRUE) {
  if (is.null(x)) {
    if (null_ok) return(NULL)
    stop(sprintf("`%s` must be a single non-NA string", arg), call. = FALSE)
  }
  if (is_xl_rich_string(x) || .cell_holds_rich_string(x))
    stop(sprintf(paste0("`%s` cannot be a rich string: only one font is ",
                        "written here, so the per-run fonts would be lost"),
                 arg), call. = FALSE)
  if (!is.character(x) && !is.null(attr(x, "class"))) x <- as.character(x)
  if (!is.character(x) || length(x) != 1L || is.na(x))
    stop(sprintf(paste0("`%s` must be a single non-NA string, or an object ",
                        "with an as.character() method such as an ",
                        "xl_cell_general()"), arg), call. = FALSE)
  x
}

.cell_holds_rich_string <- function(x) {
  inherits(x, "xl_cell_general") &&
    any(vapply(x, function(el) is_xl_rich_string(el[["value"]]), logical(1)))
}

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
#' @family cell content
#' @param x character vector to be interpreted as formula
#' @param format An optional [xl_format] (or list of `xl_format`, one per
#'   element) applied to the cells.  See [xl_format].
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
xl_formula <- function(x, format = NULL){
  x <- .as_character_arg(x, "x")
  if(!all(grepl("^=",x) | is.na(x)))
    stop("Formulas must start with '='")
  xl_cell_general(formula = x, format = format)
}

#' @rdname xl_formula
#' @export
#' @param url character vector of URLs.  Use `NA` to produce a blank cell.
#' @param name **Deprecated.** The former spelling of `value`, kept for
#'   backward compatibility.  Supplying it warns and points at `value`;
#'   supplying both is an error, since they mean the same thing.  `value` has
#'   taken the argument position `name` used to occupy, so code that passed the
#'   display text positionally keeps working unchanged.
xl_hyperlink <- function(url, value = NULL, format = NULL, name = NULL){
  if(!is.null(name)){
    if(!is.null(value))
      stop("Give either `value` or the deprecated `name`, not both: they mean ",
           "the same thing", call. = FALSE)
    warning("The `name` argument of xl_hyperlink() is deprecated; use `value` ",
            "instead, which is what xl_hyperlink_cell() and xl_cell_general() ",
            "already call it.", call. = FALSE)
    value <- name
  }
  url <- .as_character_arg(url, "url")
  fmlas <- if(!is.null(value)){
    paste0("=HYPERLINK(", dubquote(url), ",", dubquote(value), ")")
  } else {
    paste0("=HYPERLINK(", dubquote(url), ")")
  }
  fmlas[is.na(url)] <- NA_character_
  xl_formula(fmlas, format = format)
}

#' @rdname xl_formula
#' @export
#' @param value character vector (or `NULL`) of display text shown in the cell
#'   instead of the URL.  When `NULL` the URL itself is shown.  Recycled to the
#'   length of `url`, and automatically `NA` for cells whose URL is `NA`.  The
#'   same argument name is used by [xl_hyperlink_cell()] and
#'   [xl_cell_general()].
xl_hyperlink_cell <- function(url, value = NULL, format = NULL){
  url <- .as_character_arg(url, "url")
  if(is.null(value)){
    xl_cell_general(hyperlink = url, format = format)
  } else {
    value_arg <- rep_len(as.character(value), length(url))
    value_arg[is.na(url)] <- NA_character_
    xl_cell_general(value = value_arg, hyperlink = url, format = format)
  }
}

# Wrap x in Excel double-quotes, doubling any internal double-quote characters.
dubquote <- function(x){
  paste0('"', gsub('"', '""', x, fixed = TRUE), '"')
}

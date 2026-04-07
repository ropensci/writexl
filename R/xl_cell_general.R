#' General cell objects for Excel writing
#'
#' @description
#' `xl_cell_general` creates a vector of cell objects, each optionally
#' containing a **value**, a **formula**, and/or a **hyperlink**. It is the
#' fundamental building block used internally by [xl_formula()] and
#' [xl_hyperlink()], and can be used directly for mixed-type columns or cells
#' that combine multiple features (e.g., a formula with a pre-calculated
#' result, or a URL with separate display text and tooltip).
#'
#' An `xl_cell_general` behaves like a vector: it has a `length()`, supports
#' `[`, `c()`, and `rep()`, and recycles automatically when assigned to a
#' data frame column of a different length (just like [xl_formula()]).
#'
#' @param value An atomic vector or a list of scalars, one per cell. `NA`
#'   means no value is written for that cell. A list enables mixed types
#'   across cells in the same column (e.g., `list(1.5, "text", TRUE)`). Date
#'   and POSIXct scalars are supported and formatted as in [write_xlsx()].
#' @param formula A character vector of Excel formulas (each must start with
#'   `"="`), or `NA` for cells with no formula.  When both `value` and
#'   `formula` are supplied for the same cell, `value` is used as a
#'   pre-calculated result stored alongside the formula via
#'   `worksheet_write_formula_num()` (numeric value) or
#'   `worksheet_write_formula_str()` (character value).  This allows static
#'   xlsx exports that display formula text in the formula bar but do not
#'   require Excel to recalculate on open.
#' @param hyperlink A character vector of URLs, or a list where each element
#'   is `NA`, a single character URL, or a named list with elements:
#'   \describe{
#'     \item{`url`}{(required) The target URL.}
#'     \item{`string`}{(optional) Display text shown in the cell instead of
#'       the raw URL.}
#'     \item{`tooltip`}{(optional) Tooltip text shown on hover.}
#'   }
#'   Hyperlinks take priority over formulas when both are set for the same
#'   cell (the hyperlink is written; the formula is ignored).
#'
#' @return An object of class `c("xl_cell_general", "xl_cell")`, which is a
#'   list of length `n` where each element is a named list with fields
#'   `value`, `formula`, and `hyperlink`.
#'
#' @family writexl
#' @seealso [xl_formula()], [xl_hyperlink()], [write_xlsx()]
#' @export
#' @examples
#' # Value-only cell
#' xl_cell_general(value = 42)
#'
#' # Formula with a pre-calculated numeric result (static export)
#' xl_cell_general(value = 42.0, formula = "=SUM(A1:A10)")
#'
#' # Hyperlink with display text and tooltip
#' xl_cell_general(
#'   hyperlink = list(
#'     url     = "https://example.com",
#'     string  = "Visit",
#'     tooltip = "Go to example.com"
#'   )
#' )
#'
#' # Vector of cells: value and formula cells in one column
#' cells <- c(
#'   xl_cell_general(value = 1.5),
#'   xl_cell_general(value = "note"),
#'   xl_cell_general(formula = "=A1+A2")
#' )
#'
#' # Used in a data frame (length-1 recycles to fill all rows, as with
#' # xl_formula())
#' df <- data.frame(x = 1:3)
#' df$formula_col <- xl_formula("=A1*2")   # backward-compatible shorthand
#' df$cell_col    <- xl_cell_general(value = 99L)  # all rows get 99
xl_cell_general <- function(value = NULL, formula = NULL, hyperlink = NULL) {

  # 0. Pre-normalise hyperlink: a named list with 'url' is a single hyperlink
  #    spec, not a list of multiple hyperlinks.  Wrap it so length() = 1.
  if (is.list(hyperlink) && !is.null(hyperlink[["url"]])) {
    hyperlink <- list(hyperlink)
  }

  # 1. Determine output length n -------------------------------------------
  n <- max(c(
    if (!is.null(value))     length(value)     else 0L,
    if (!is.null(formula))   length(formula)   else 0L,
    if (!is.null(hyperlink)) length(hyperlink) else 0L,
    1L
  ))

  # 2. Normalise value to a list of length n --------------------------------
  value_list <- if (is.null(value)) {
    rep(list(NA), n)
  } else {
    rep_len(if (is.list(value)) value else as.list(value), n)
  }

  # 3. Normalise formula to a character vector of length n ------------------
  formula_vec <- if (is.null(formula)) {
    rep(NA_character_, n)
  } else {
    formula <- as.character(formula)
    bad <- !is.na(formula) & !startsWith(formula, "=")
    if (any(bad))
      stop("All non-NA formulas must start with '='", call. = FALSE)
    rep_len(formula, n)
  }

  # 4. Normalise hyperlink to a list of length n ----------------------------
  hyperlink_list <- if (is.null(hyperlink)) {
    rep(list(NA), n)
  } else {
    h <- if (is.list(hyperlink)) hyperlink else as.list(hyperlink)
    for (i in seq_along(h)) {
      el <- h[[i]]
      if (is.null(el))                                              next
      if (identical(el, NA))                                        next
      # character NA -> normalise to logical NA sentinel
      if (is.character(el) && length(el) == 1L && is.na(el)) {
        h[[i]] <- NA
        next
      }
      if (is.character(el) && length(el) == 1L)                    next
      if (is.list(el) && !is.null(el[["url"]]) &&
          is.character(el[["url"]]) && length(el[["url"]]) == 1L)  next
      stop(sprintf(
        paste0("hyperlink[[%d]] must be NA, a single character URL, or a ",
               "named list with a character 'url' element"),
        i), call. = FALSE)
    }
    rep_len(h, n)
  }

  # 5. Build per-cell records -----------------------------------------------
  cells <- lapply(seq_len(n), function(i) {
    list(
      value     = value_list[[i]],
      formula   = formula_vec[[i]],
      hyperlink = hyperlink_list[[i]]
    )
  })

  structure(cells, class = c("xl_cell_general", "xl_cell"))
}

# --- S3 methods --------------------------------------------------------------

#' @export
length.xl_cell_general <- function(x) length(unclass(x))

#' @export
`[.xl_cell_general` <- function(x, i, ...) {
  structure(unclass(x)[i], class = class(x))
}

#' @export
c.xl_cell_general <- function(x, ...) {
  structure(c(unclass(x), ...), class = class(x))
}

#' @export
rep.xl_cell_general <- function(x, ...) {
  structure(rep(unclass(x), ...), class = class(x))
}

#' @export
as.data.frame.xl_cell_general <- function(x, ...) {
  # Store the xl_cell_general object as a single list column
  structure(
    list(x),
    names     = "x",
    class     = "data.frame",
    row.names = seq_len(length(x))
  )
}

#' @export
print.xl_cell_general <- function(x, max = 10L, ...) {
  n <- length(x)
  cat(sprintf("[xl_cell_general: %d cell%s]\n", n, if (n == 1L) "" else "s"))
  show <- seq_len(min(n, max))
  for (i in show) {
    cell  <- unclass(x)[[i]]
    parts <- character(0L)

    val <- cell[["value"]]
    has_val <- !is.null(val) && length(val) > 0L &&
               !all(is.na(val))
    if (has_val)
      parts <- c(parts, paste0("value=", format(val)))

    fml <- cell[["formula"]]
    if (!is.null(fml) && !is.na(fml))
      parts <- c(parts, paste0("formula=", fml))

    hlk <- cell[["hyperlink"]]
    hlk_set <- !is.null(hlk) && !identical(hlk, NA) &&
               !(is.character(hlk) && length(hlk) == 1L && is.na(hlk))
    if (hlk_set)
      parts <- c(parts, "hyperlink=<set>")

    cat(sprintf("  [%d] %s\n", i,
                if (length(parts)) paste(parts, collapse = ", ")
                else "<empty>"))
  }
  if (n > max)
    cat(sprintf("  ... (%d more)\n", n - max))
  invisible(x)
}

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
#' @param value An atomic vector or a list of scalars, one per cell. Use `NA`
#'   to write an explicit empty cell. A list enables mixed types across cells
#'   in the same column (e.g., `list(1.5, "text", TRUE)`). Date and POSIXct
#'   scalars are supported and formatted as in [write_xlsx()].  When
#'   `hyperlink` is also set for the same cell, a **character** `value` is
#'   used as the display text shown in the cell instead of the raw URL; all
#'   other types are ignored for hyperlink cells.
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
#'     \item{`tooltip`}{(optional) Tooltip text shown on hover.}
#'   }
#'   Supply a character `value` alongside `hyperlink` to show custom display
#'   text in the cell instead of the raw URL.  The hyperlink is written via
#'   `worksheet_write_url_opt()`.
#'
#' @param format An [xl_format] object (applied to every cell), or a list of
#'   `xl_format` objects (one per cell, recycled), or `NULL` for no formatting.
#'   Build formats with [xl_font()], [xl_fill()], [xl_border()], [xl_align()],
#'   [xl_num_format()] and [xl_protection()], combined with [xl_format()] or
#'   `+`.  When a cell's value is a date/time and its format sets no number
#'   format, the default date/time number format is applied automatically.
#' @param comment Cell comments (notes): a character vector of comment text
#'   (one per cell, recycled; `NA` for no comment), a single [xl_comment()]
#'   (recycled to every cell), or a list mixing strings / `xl_comment` / `NA`
#'   per cell.  `NULL` for no comments.
#' @param array,dynamic Logical (one per cell, recycled): how the cell's
#'   `formula` is stored.  `array = TRUE` writes a legacy *array* (Ctrl-Shift-
#'   Enter) formula; `dynamic = TRUE` writes a modern *dynamic array* formula,
#'   which Excel spills over as many cells as the result needs.  Neither can
#'   carry a character `value`, because Excel stores no cached string result for
#'   an array formula.  On a cell that has no `formula` the flags are inert, so a
#'   single `array = TRUE` can be recycled across a column that mixes formula
#'   and value cells.
#' @param array_range The range a legacy array formula covers, for the rare case
#'   where it must be declared: an Excel range string (`"C2:C11"`) or a
#'   `list(rows = , cols = )` spec, one per cell (`NA` for none). It must start
#'   at the cell holding the formula, and must extend into cells the sheet does
#'   not otherwise write --- the range is padded on write, so an overlap would
#'   have the padding and the sheet's own values overwrite each other.
#'
#'   Leave it unset for almost everything: a single-cell `dynamic` formula
#'   spills automatically in Excel 365 / 2021, and a single-cell `array` formula
#'   is the right spelling for a `SUMPRODUCT`-style aggregate. Supplying it
#'   forces the workbook out of the memory-efficient row-streaming mode, since
#'   libxlsxwriter cannot pad an array range while streaming.
#' @return An object of class `c("xl_cell_general", "xl_cell")`, which is a
#'   list of length `n` where each element is a named list with fields
#'   `value`, `formula`, `hyperlink`, `format`, `comment`, `array`, `dynamic`
#'   and `array_range`.
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
#' # Hyperlink with display text (value) and tooltip
#' xl_cell_general(
#'   value    = "Visit",
#'   hyperlink = list(url = "https://example.com", tooltip = "Go to example.com")
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
xl_cell_general <- function(value = NULL, formula = NULL, hyperlink = NULL,
                            format = NULL, comment = NULL, array = FALSE,
                            dynamic = FALSE, array_range = NULL) {

  # Require at least one content argument -------------------------------------
  if (is.null(value) && is.null(formula) && is.null(hyperlink) &&
      is.null(comment))
    stop("At least one of 'value', 'formula', 'hyperlink', or 'comment' must ",
         "be provided. Use value = NA for an explicit empty cell.", call. = FALSE)

  # Reject values writexl cannot write ----------------------------------------
  # Uses the same predicate as the column-level check in normalize_df(), so an
  # unsupported type cannot slip in by being wrapped in a cell object.  NULL
  # and NA are allowed: both mean "blank cell".
  if (!is.null(value)) {
    vals <- if (is.list(value)) value else list(value)
    keep <- !vapply(vals, is.null, logical(1))
    bad <- which(keep & !vapply(vals, .is_supported_column, logical(1)))
    if (length(bad)) {
      detail <- vapply(bad, function(k) {
        v <- vals[[k]]
        cls <- paste(class(v), collapse = "/")
        ty <- typeof(v)
        lab <- if (identical(cls, ty)) cls else sprintf("%s (%s)", cls, ty)
        if (is.list(value)) sprintf("value[[%d]]: %s", k, lab)
        else sprintf("value: %s", lab)
      }, character(1))
      stop("xl_cell_general() cannot write these value type(s):\n  ",
           paste(detail, collapse = "\n  "), call. = FALSE)
    }
  }

  # Pre-normalise the hyperlink argument --------------------------------------
  # A named list with a 'url' element is a single hyperlink spec, not a list of
  # multiple hyperlinks.  Wrap it so length() = 1.
  if (is.list(hyperlink) && !is.null(hyperlink[["url"]])) {
    hyperlink <- list(hyperlink)
  }

  # Pre-normalise the array_range argument ------------------------------------
  # Same wrapping problem as `hyperlink`: a list(rows = , cols = ) spec is one
  # range, not a list of ranges.
  if (is.list(array_range) &&
      (!is.null(array_range[["rows"]]) || !is.null(array_range[["cols"]]))) {
    array_range <- list(array_range)
  }

  # Determine the output length n ---------------------------------------------
  # A single xl_comment counts as one comment (not its number of fields).
  comment_n <- if (is.null(comment)) 0L
               else if (is_xl_comment(comment)) 1L
               else length(comment)
  n <- max(c(
    if (!is.null(value))     length(value)     else 0L,
    if (!is.null(formula))   length(formula)   else 0L,
    if (!is.null(hyperlink)) length(hyperlink) else 0L,
    if (!is.null(array_range)) length(array_range) else 0L,
    comment_n
  ))

  # Normalise value to a list of length n -------------------------------------
  value_list <- if (is.null(value)) {
    rep(list(NA), n)
  } else {
    rep_len(if (is.list(value)) value else as.list(value), n)
  }

  # Normalise formula to a character vector of length n -----------------------
  formula_vec <- if (is.null(formula)) {
    rep(NA_character_, n)
  } else {
    formula <- as.character(formula)
    bad <- !is.na(formula) & !startsWith(formula, "=")
    if (any(bad))
      stop("All non-NA formulas must start with '='", call. = FALSE)
    rep_len(formula, n)
  }

  # Normalise hyperlink to a list of length n ---------------------------------
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

  # Normalise format to a list of length n ------------------------------------
  format_list <- if (is.null(format)) {
    vector("list", n)
  } else if (is_xl_format(format)) {
    rep(list(format), n)
  } else if (is.list(format)) {
    for (el in format)
      if (!is.null(el) && !is_xl_format(el))
        stop("each element of `format` must be an xl_format object or NULL",
             call. = FALSE)
    rep_len(format, n)
  } else {
    stop("`format` must be an xl_format object or a list of xl_format objects",
         call. = FALSE)
  }

  # Normalise comment to a list of length n -----------------------------------
  # Each element is NULL (no comment) or a C-ready comment payload.  A
  # character vector gives per-cell text with default options; a single
  # xl_comment recycles to every cell; a list mixes strings/xl_comment/NA.
  comment_list <- if (is.null(comment)) {
    vector("list", n)
  } else if (is_xl_comment(comment)) {
    rep(list(unclass(comment)), n)
  } else if (is.character(comment) || is.list(comment)) {
    rep_len(lapply(comment, .comment_payload), n)
  } else {
    stop("`comment` must be a character vector, an xl_comment, or a list of ",
         "strings/xl_comment objects", call. = FALSE)
  }

  # Normalise the array / dynamic flags and their range -----------------------
  array_vec   <- .cell_flag_vec(array, "array", n)
  dynamic_vec <- .cell_flag_vec(dynamic, "dynamic", n)
  range_list  <- .cell_range_list(array_range, n)
  .check_array_args(value_list, formula_vec, array_vec, dynamic_vec, range_list,
                    formula_given = !is.null(formula))

  # Build the per-cell records ------------------------------------------------
  cells <- lapply(seq_len(n), function(i) {
    list(
      value       = value_list[[i]],
      formula     = formula_vec[[i]],
      hyperlink   = hyperlink_list[[i]],
      format      = format_list[[i]],
      comment     = comment_list[[i]],
      array       = array_vec[[i]],
      dynamic     = dynamic_vec[[i]],
      array_range = range_list[[i]]
    )
  })

  structure(cells, class = c("xl_cell_general", "xl_cell"))
}

# --- array / dynamic formula helpers -----------------------------------------

# Recycle a logical flag argument to length n, refusing NA (a flag is a
# decision, and NA would leave it ambiguous).
.cell_flag_vec <- function(x, arg, n) {
  if (is.null(x)) return(rep(FALSE, n))
  if (!is.logical(x) || !length(x) || anyNA(x))
    stop(sprintf("`%s` must be TRUE or FALSE (no NA)", arg), call. = FALSE)
  rep_len(x, n)
}

# Normalise array_range to a list of length n whose elements are NULL (no
# declared range) or a range spec to be resolved at write time.
.cell_range_list <- function(x, n) {
  if (is.null(x)) return(vector("list", n))
  r <- if (is.list(x)) x else as.list(x)
  r <- lapply(r, function(el) {
    if (is.null(el) || (length(el) == 1L && !is.list(el) && is.na(el)))
      return(NULL)
    if (is.character(el) && length(el) == 1L) return(el)
    if (is.list(el)) return(el)
    stop("each `array_range` must be NA, an Excel range string, or a ",
         "list(rows = , cols = ) spec", call. = FALSE)
  })
  rep_len(r, n)
}

# Validate the array/dynamic arguments.
#
# `array` / `dynamic` describe how a *formula* is stored, so on a cell with no
# formula they are simply inert -- which is what recycling a single TRUE across
# a column of mixed formula and value cells has to mean.  What is worth
# refusing is a flag that can never apply to anything, and a combination Excel
# cannot represent.
.check_array_args <- function(values, formulas, arrays, dynamics, ranges,
                              formula_given) {
  any_flag  <- any(arrays) || any(dynamics)
  any_range <- any(!vapply(ranges, is.null, logical(1)))
  if (any_flag && !formula_given)
    stop("`array`/`dynamic` describe how a formula is stored, but no `formula` ",
         "was supplied", call. = FALSE)
  if (any_range && !any_flag)
    stop("`array_range` applies only to an `array` or `dynamic` formula, and ",
         "neither is set", call. = FALSE)
  for (i in seq_along(formulas)) {
    if (is.na(formulas[[i]])) next          # flags are inert without a formula
    if (!(arrays[[i]] || dynamics[[i]])) next
    # libxlsxwriter has write_array_formula_num() but no _str() counterpart:
    # Excel stores no cached string result for an array formula.  Erroring beats
    # silently downgrading to a plain formula and losing the array semantics.
    v <- values[[i]]
    if (is.character(v) && length(v) == 1L && !is.na(v))
      stop(sprintf(paste0("cell %d combines a character `value` with ",
                          "`array`/`dynamic`; only a numeric pre-calculated ",
                          "result can be stored for an array formula"), i),
           call. = FALSE)
  }
  invisible(NULL)
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
as.data.frame.xl_cell_general <- function(x, row.names = NULL, optional = FALSE, ...) {
  # Omit names so that data.frame() uses the caller's argument name rather
  # than an arbitrary inner name.  This mirrors how I(list(...)) behaves:
  # names(xi) == NULL causes data.frame() to leave vnames[[i]] as the
  # argument name instead of overriding it.
  structure(
    list(x),
    class     = "data.frame",
    row.names = if (is.null(row.names)) seq_len(length(x)) else row.names
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

    if (is_xl_format(cell[["format"]]))
      parts <- c(parts, "format=<set>")

    if (!is.null(cell[["comment"]]))
      parts <- c(parts, "comment=<set>")

    cat(sprintf("  [%d] %s\n", i,
                if (length(parts)) paste(parts, collapse = ", ")
                else "<empty>"))
  }
  if (n > max)
    cat(sprintf("  ... (%d more)\n", n - max))
  invisible(x)
}

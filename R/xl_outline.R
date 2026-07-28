# =============================================================================
# Outline (grouping) display, and the error indicators Excel shows in cells
# =============================================================================
#
# Two small worksheet-level settings that share nothing but their size.
#
# Outline settings control how the grouping controls that xl_col_spec(level =)
# and xl_row_spec(level =) create are *drawn* -- not which rows are grouped.
#
# Ignoring errors turns off the green triangle Excel puts in a cell it thinks
# is wrong.  The common case is a number deliberately written as text: Excel
# flags it, and there is no other way to silence that.
# -----------------------------------------------------------------------------

#' Control how outline (grouping) symbols are drawn
#'
#' @description
#' Grouping itself comes from `level` in [xl_col_spec()] / [xl_row_spec()].
#' `xl_outline()` only changes how the controls are *displayed*, which is
#' rarely needed --- the defaults match Excel's own.
#'
#' @param visible Show the outline symbols at all.  `FALSE` keeps the grouping
#'   (and so the collapsing) but hides the +/- controls.
#' @param symbols_below Put the summary row *below* the detail rows, which is
#'   Excel's default.  `FALSE` puts it above.
#' @param symbols_right Put the summary column to the *right* of the detail
#'   columns, Excel's default.  `FALSE` puts it to the left.
#' @param auto_style Apply Excel's automatic outline styling to the grouped
#'   rows and columns.
#' @return An `xl_outline` object, for [xl_sheet()]'s `outline` argument.
#' @family writexl
#' @seealso [xl_sheet], [xl_colrow_spec]
#' @export
#' @examples
#' xl_outline(symbols_below = FALSE)
#' xl_outline(visible = FALSE)
xl_outline <- function(visible = TRUE, symbols_below = TRUE,
                       symbols_right = TRUE, auto_style = FALSE) {
  flag <- function(x, arg) {
    if (!is.logical(x) || length(x) != 1L || is.na(x))
      stop(sprintf("`%s` must be TRUE or FALSE", arg), call. = FALSE)
    as.integer(x)
  }
  structure(list(visible = flag(visible, "visible"),
                 symbols_below = flag(symbols_below, "symbols_below"),
                 symbols_right = flag(symbols_right, "symbols_right"),
                 auto_style = flag(auto_style, "auto_style")),
            class = "xl_outline")
}

#' @export
print.xl_outline <- function(x, ...) {
  p <- unclass(x)
  cat(sprintf("<xl_outline: visible=%s below=%s right=%s auto_style=%s>\n",
              p$visible == 1L, p$symbols_below == 1L, p$symbols_right == 1L,
              p$auto_style == 1L))
  invisible(x)
}

.resolve_outline <- function(el) {
  if (!inherits(el, "xl_sheet") || is.null(el$outline)) return(NULL)
  if (!inherits(el$outline, "xl_outline"))
    stop("`outline` must be an xl_outline object", call. = FALSE)
  unclass(el$outline)
}

# The error indicators Excel can be told to stop showing.  The names mirror
# libxlsxwriter's LXW_IGNORE_* constants, lowercased.
.LXW_IGNORE_TYPE <- c(
  number_stored_as_text = 1L, eval_error = 2L, formula_differs = 3L,
  formula_range = 4L, formula_unlocked = 5L, empty_cell_reference = 6L,
  list_data_validation = 7L, calculated_column = 8L, two_digit_text_year = 9L
)

# Resolve xl_sheet(ignore_errors =) to the C payloads.  Each element is named
# for the error type and holds a range, in any form the shared resolver takes.
.resolve_ignore_errors <- function(el, df, header_offset) {
  if (!inherits(el, "xl_sheet") || is.null(el$ignore_errors)) return(list())
  ie <- el$ignore_errors
  if (!is.list(ie)) ie <- as.list(ie)
  nms <- names(ie)
  if (is.null(nms) || any(!nzchar(nms)))
    stop("`ignore_errors` must be a named list, each name an error type ",
         "and each value a range", call. = FALSE)
  bad <- setdiff(nms, names(.LXW_IGNORE_TYPE))
  if (length(bad))
    stop(sprintf("`ignore_errors`: unknown error type(s): %s.\n  Available: %s",
                 paste(bad, collapse = ", "),
                 paste(names(.LXW_IGNORE_TYPE), collapse = ", ")),
         call. = FALSE)
  if (anyDuplicated(nms))
    stop("`ignore_errors`: each error type may be given only once; combine ",
         "the ranges into one entry", call. = FALSE)
  lapply(seq_along(ie), function(i) {
    rng <- .xl_resolve_range(ie[[i]], arg = sprintf("ignore_errors$%s", nms[i]),
                             df = df, header_offset = header_offset,
                             allow_cell = TRUE)
    list(kind = "ignore_errors",
         type = unname(.LXW_IGNORE_TYPE[[nms[i]]]),
         range = as.integer(rng))
  })
}

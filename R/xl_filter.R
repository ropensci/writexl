# =============================================================================
# Autofilter criteria, and the row hiding they imply
# =============================================================================
#
# Setting a filter criteria is not enough.  Excel stores the criteria and the
# row-hidden flags independently and does *not* apply the filter when the file
# is opened, so a file with criteria alone shows the filter dropdown set and
# every row still visible.  writexl therefore evaluates the filter against the
# data frame and hides the rows that do not match.
#
# That makes writexl responsible for reproducing Excel's matching rules
# exactly.  They were established by writing a workbook with criteria but no
# hidden rows, opening it in Excel and using Data > Reapply, which makes Excel
# compute the match itself:
#
#   * text comparison is case-INsensitive         ("Apple" matches "apple")
#   * "*" and "?" in a text value are wildcards   ("ap*" matches "apricot")
#   * blank means an empty cell OR an empty string, and non-blank is its exact
#     complement -- the two partition the column
#   * there is no type coercion: a numeric criteria matches no text cell
#   * list membership follows the same case-insensitive rule as equality
#
# The one rule inferred rather than observed is the mirror of the fourth: a
# text criteria is taken to match no numeric cell, by symmetry with the tested
# direction.
# -----------------------------------------------------------------------------

.LXW_FILTER_CRITERIA <- c(
  none = 0L, "==" = 1L, "!=" = 2L, ">" = 3L, "<" = 4L, ">=" = 5L, "<=" = 6L,
  blanks = 7L, "non-blanks" = 8L
)

.LXW_FILTER_AND_OR <- c(and = 0L, or = 1L)

# The criteria that compare magnitudes, and so need a numeric column.
.FILTER_NUMERIC_CRITERIA <- c(">", "<", ">=", "<=")

# Excel treats an absent cell and an empty string alike (verified).
.filter_is_blank <- function(x) {
  if (is.character(x)) is.na(x) | !nzchar(x) else is.na(x)
}

# Excel matches text case-insensitively and honours * and ? as wildcards.
.filter_text_match <- function(x, value) {
  pat <- utils::glob2rx(as.character(value))
  grepl(pat, as.character(x), ignore.case = TRUE)
}

# Evaluate one rule against a column, returning TRUE where the row is kept.
# A blank never matches a comparison in Excel, so NA propagates to FALSE.
.filter_rule_keep <- function(x, criteria, value) {
  blank <- .filter_is_blank(x)
  if (criteria == "blanks")     return(blank)
  if (criteria == "non-blanks") return(!blank)

  numeric_col <- is.numeric(x) || inherits(x, "Date") || inherits(x, "POSIXct")
  numeric_val <- is.numeric(value) || inherits(value, "Date") ||
                 inherits(value, "POSIXct")

  # no coercion in either direction: a criteria of the wrong kind matches
  # nothing at all
  if (criteria %in% .FILTER_NUMERIC_CRITERIA && !numeric_col)
    return(rep(FALSE, length(x)))
  if (numeric_val && !numeric_col) return(rep(FALSE, length(x)))
  if (!numeric_val && numeric_col) return(rep(FALSE, length(x)))

  keep <- if (numeric_col) {
    v <- as.numeric(value)
    n <- as.numeric(x)
    switch(criteria,
           "==" = n == v, "!=" = n != v, ">" = n > v,
           "<"  = n <  v, ">=" = n >= v, "<=" = n <= v)
  } else {
    m <- .filter_text_match(x, value)
    switch(criteria, "==" = m, "!=" = !m,
           stop(sprintf("`criteria = \"%s\"` needs a numeric column", criteria),
                call. = FALSE))
  }
  keep & !blank
}

#' Filter an autofilter column, hiding the rows that do not match
#'
#' @description
#' `xl_filter()` sets the criteria on one autofilter column *and* hides the rows
#' that do not match.  Both halves are necessary: Excel stores the criteria and
#' the hidden rows separately and does not apply a filter when a file is opened,
#' so criteria alone produce a sheet that looks filtered but shows every row.
#'
#' Because writexl decides which rows to hide, it reproduces Excel's own
#' matching rules: text comparison is case-insensitive, `*` and `?` are
#' wildcards, blank covers both an empty cell and an empty string, and there is
#' no coercion between text and numbers.
#'
#' @param col The column to filter: a name or a 1-based position.
#' @param criteria One of `"=="`, `"!="`, `">"`, `"<"`, `">="`, `"<="`,
#'   `"blanks"` or `"non-blanks"`.  The four magnitude comparisons need a
#'   numeric, `Date` or `POSIXct` column.
#' @param value The value to compare against.  For text columns `*` matches any
#'   run of characters and `?` any single one.
#' @param criteria2,value2 An optional second rule for the same column.
#' @param and_or How the two rules combine: `"and"` (default) or `"or"`.
#' @param list Instead of a criteria, keep only rows whose value is in this
#'   character vector.  Matched case-insensitively, as Excel does.
#' @return An `xl_filter` object.
#' @family writexl
#' @seealso [xl_sheet]
#' @export
#' @examples
#' xl_filter("qty", ">", 100)
#' xl_filter("qty", ">", 100, "<", 200)          # between, via two rules
#' xl_filter("fruit", "==", "ap*")               # wildcard
#' xl_filter("fruit", list = c("apple", "banana"))
xl_filter <- function(col, criteria = NA, value = NULL, criteria2 = NA,
                      value2 = NULL, and_or = "and", list = NULL) {
  if (missing(col) || is.null(col))
    stop("`col` must name or index the column to filter", call. = FALSE)
  crit  <- .val_enum(criteria, names(.LXW_FILTER_CRITERIA)[-1L], "criteria")
  crit2 <- .val_enum(criteria2, names(.LXW_FILTER_CRITERIA)[-1L], "criteria2")
  ao    <- .val_enum(and_or, names(.LXW_FILTER_AND_OR), "and_or")

  if (!is.null(list)) {
    if (!is.null(crit))
      stop("give either `list` or `criteria`, not both", call. = FALSE)
    if (!is.character(list) || !length(list) || anyNA(list))
      stop("`list` must be a character vector of values to keep", call. = FALSE)
  } else {
    if (is.null(crit))
      stop("`xl_filter()` needs a `criteria` or a `list`", call. = FALSE)
    if (!crit %in% c("blanks", "non-blanks") && is.null(value))
      stop(sprintf("`criteria = \"%s\"` needs a `value`", crit), call. = FALSE)
    if (!is.null(crit2) && !crit2 %in% c("blanks", "non-blanks") &&
        is.null(value2))
      stop(sprintf("`criteria2 = \"%s\"` needs a `value2`", crit2),
           call. = FALSE)
  }
  structure(.drop_null(base::list(col = col, criteria = crit, value = value,
                                  criteria2 = crit2, value2 = value2,
                                  and_or = ao %||% "and", list = list)),
            class = "xl_filter")
}

#' @export
print.xl_filter <- function(x, ...) {
  p <- unclass(x)
  cat(sprintf("<xl_filter: %s %s>\n", p$col,
              if (!is.null(p$list)) sprintf("in %d value(s)", length(p$list))
              else paste(p$criteria, format(p$value))))
  invisible(x)
}

# Which data rows one filter keeps.
.filter_keep <- function(f, df) {
  p <- unclass(f)
  idx <- .resolve_col_index(p$col, names(df), "filter col")
  x <- df[[idx]]
  if (inherits(x, "xl_cell_general"))
    stop(sprintf(paste0("cannot filter column \"%s\": its cells carry formulas ",
                        "or mixed content, so writexl cannot know which rows ",
                        "Excel would keep. Filter on a plain column instead."),
                 names(df)[idx]), call. = FALSE)
  if (!is.null(p$list))
    return(tolower(as.character(x)) %in% tolower(p$list) &
             !.filter_is_blank(x))
  keep <- .filter_rule_keep(x, p$criteria, p$value)
  if (!is.null(p$criteria2)) {
    k2 <- .filter_rule_keep(x, p$criteria2, p$value2)
    keep <- if (identical(p$and_or, "or")) keep | k2 else keep & k2
  }
  keep
}

# Resolve a sheet's filters: the C payloads, plus the 0-based sheet rows to
# hide because they do not match.
.resolve_filters <- function(el, df, header_offset) {
  if (!inherits(el, "xl_sheet") || is.null(el$filter))
    return(list(payloads = list(), hide = integer(0)))
  fs <- if (inherits(el$filter, "xl_filter")) list(el$filter) else el$filter
  if (!is.list(fs))
    stop("`filter` must be an xl_filter object or a list of them", call. = FALSE)

  keep <- rep(TRUE, nrow(df))
  payloads <- lapply(seq_along(fs), function(i) {
    f <- fs[[i]]
    if (!inherits(f, "xl_filter"))
      stop(sprintf("`filter[[%d]]` must be an xl_filter object", i),
           call. = FALSE)
    p <- unclass(f)
    # filters on different columns combine with AND, as they do in Excel
    keep <<- keep & .filter_keep(f, df)
    out <- list(kind = "filter",
                col = .resolve_col_index(p$col, names(df), "filter col") - 1L)
    if (!is.null(p$list)) {
      out$value_list <- as.character(p$list)
    } else {
      out$criteria <- unname(.LXW_FILTER_CRITERIA[[p$criteria]])
      out <- c(out, .filter_value_payload(p$value, ""))
      if (!is.null(p$criteria2)) {
        out$criteria2 <- unname(.LXW_FILTER_CRITERIA[[p$criteria2]])
        out <- c(out, .filter_value_payload(p$value2, "2"))
        out$and_or <- unname(.LXW_FILTER_AND_OR[[p$and_or]])
      }
    }
    out
  })
  list(payloads = payloads,
       hide = as.integer(which(!keep) - 1L + header_offset))
}

# A rule's value, as the number or string the struct carries.
.filter_value_payload <- function(v, suffix) {
  if (is.null(v)) return(list())
  if (is.character(v))
    return(stats::setNames(list(as.character(v)),
                           paste0("value_string", suffix)))
  stats::setNames(list(as.numeric(v)), paste0("value", suffix))
}

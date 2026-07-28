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
# compute the match itself.
#
# What the measurements showed is that the rules are not a property of the
# criteria alone: they are a property of the XML form libxlsxwriter chooses.
# _set_custom_filter() in worksheet.c picks one of two forms, and Excel matches
# them by entirely different rules:
#
#   <filters>        "==" and "blanks", unless the value holds a * or ?
#   <customFilters>  everything else, including "==" with a wildcard
#
#   <filters>        matches the cell's DISPLAYED TEXT, case-insensitively.
#                    The R type of `value` does not enter into it: `== 10`,
#                    `== "10"` and `list = "10"` all write val="10", and all
#                    three match the number 10 and the string "10" alike.
#   <customFilters>  matches by type.  If the written value parses as a number
#                    the comparison is numeric, and a text cell never satisfies
#                    it -- except "!=", which a text cell satisfies vacuously.
#                    Otherwise the comparison is textual, case-insensitive,
#                    with * and ? as wildcards, and a number cell never
#                    satisfies it.
#
# So `== "10"` keeps the number 10 (it is a value list) while `== "1*"` keeps
# nothing (it is a custom filter, and no number is text).  Only the string
# written to val matters, which is why the two disagree.
#
# Blank cells satisfy "blanks" and nothing else; "non-blanks" is its exact
# complement, so the two partition the column.
# -----------------------------------------------------------------------------

.LXW_FILTER_CRITERIA <- c(
  none = 0L, "==" = 1L, "!=" = 2L, ">" = 3L, "<" = 4L, ">=" = 5L, "<=" = 6L,
  blanks = 7L, "non-blanks" = 8L
)

.LXW_FILTER_AND_OR <- c(and = 0L, or = 1L)

# The criteria that compare magnitudes, and so need a numeric column.
.FILTER_NUMERIC_CRITERIA <- c(">", "<", ">=", "<=")

# Excel matches text case-insensitively and honours * and ? as wildcards.
.filter_text_match <- function(x, value) {
  pat <- utils::glob2rx(as.character(value))
  grepl(pat, as.character(x), ignore.case = TRUE)
}

# --- How a column looks to Excel ---------------------------------------------
# Every column is reduced to three parallel vectors: the numeric value of the
# cells Excel stores as numbers, the text of the cells it stores as strings,
# and which cells are blank.  A cell contributes to exactly one of the first
# two, which is what lets a mixed xl_cell_general column be matched per cell.
.filter_cells <- function(x, nm) {
  if (inherits(x, "xl_cell_general")) return(.filter_cells_general(x, nm))
  n <- length(x)
  none <- rep(NA_character_, n)
  if (inherits(x, "Date") || inherits(x, "POSIXct"))
    return(list(num = as.numeric(x), txt = none, blank = is.na(x),
                dated = TRUE))
  if (is.numeric(x))
    return(list(num = as.numeric(x), txt = none, blank = is.na(x),
                dated = FALSE))
  # Excel stores a logical as a boolean, displayed as TRUE/FALSE, and does not
  # compare it numerically -- so it behaves as text here.
  if (is.logical(x))
    return(list(num = rep(NA_real_, n),
                txt = ifelse(is.na(x), NA_character_, as.character(x)),
                blank = is.na(x), dated = FALSE))
  ch <- as.character(x)
  list(num = rep(NA_real_, n), txt = ch, blank = is.na(ch) | !nzchar(ch),
       dated = FALSE)
}

.filter_cells_general <- function(x, nm) {
  recs <- unclass(x)
  n <- length(recs)
  num <- rep(NA_real_, n)
  txt <- rep(NA_character_, n)
  dated <- FALSE
  for (i in seq_len(n)) {
    r <- recs[[i]]
    if (!is.null(r$formula) && length(r$formula) && !is.na(r$formula))
      stop(sprintf(paste0("cannot filter column \"%s\": row %d holds a ",
                          "formula, so writexl cannot know the value Excel ",
                          "would match. Filter on a plain column instead."),
                   nm, i), call. = FALSE)
    v <- r$value
    if (is_xl_rich_string(v)) {
      txt[i] <- format(v)
    } else if (is.null(v) || (length(v) == 1L && !is_xl_rich_string(v) &&
                              is.na(v))) {
      # no value: a hyperlink cell displays its URL
      hl <- r$hyperlink
      if (is.character(hl) && length(hl) == 1L && !is.na(hl)) txt[i] <- hl
      else if (is.list(hl) && !is.null(hl$url)) txt[i] <- as.character(hl$url)
    } else if (inherits(v, "Date") || inherits(v, "POSIXct")) {
      num[i] <- as.numeric(v)
      dated <- TRUE
    } else if (is.numeric(v)) {
      num[i] <- as.numeric(v)
    } else {
      txt[i] <- as.character(v)
    }
  }
  list(num = num, txt = txt,
       blank = is.na(num) & (is.na(txt) | !nzchar(txt)), dated = dated)
}

# The text Excel shows for a cell, which is what a <filters> list matches.
.filter_display <- function(cells) {
  out <- cells$txt
  ok <- is.na(out) & !is.na(cells$num)
  # format() per element: a shared format would pad 10 to "10.0" beside 20.5
  out[ok] <- vapply(cells$num[ok], function(z)
    format(z, trim = TRUE, scientific = FALSE, digits = 15), character(1))
  out
}

# --- Which XML form a rule takes ---------------------------------------------
# Mirrors _set_custom_filter() in worksheet.c.  It has to be mirrored rather
# than guessed, because the form decides the matching rules.
.filter_wildcard <- function(v) is.character(v) && any(grepl("[*?]", v))

.filter_is_custom <- function(p) {
  custom <- !(identical(p$criteria, "==") || identical(p$criteria, "blanks"))
  if (!is.null(p$criteria2) && identical(p$and_or, "and")) custom <- TRUE
  if (.filter_wildcard(p$value) || .filter_wildcard(p$value2)) custom <- TRUE
  custom
}

# Does the string libxlsxwriter writes to val parse as a number?  That, not the
# R type, is what decides whether Excel compares numerically.
.filter_numeric_value <- function(v) {
  if (is.null(v)) return(FALSE)
  if (.filter_wildcard(v)) return(FALSE)
  if (is.numeric(v) || inherits(v, "Date") || inherits(v, "POSIXct"))
    return(TRUE)
  !is.na(suppressWarnings(as.numeric(as.character(v))))
}

# --- Evaluation ---------------------------------------------------------------
# A value list: the displayed text, matched case-insensitively.
.filter_display_match <- function(cells, vals) {
  d <- tolower(.filter_display(cells))
  !is.na(d) & d %in% tolower(as.character(vals)) & !cells$blank
}

# One custom-filter rule.
.filter_custom_keep <- function(cells, criteria, value) {
  if (criteria == "blanks")     return(cells$blank)
  if (criteria == "non-blanks") return(!cells$blank)
  n <- cells$num
  if (.filter_numeric_value(value)) {
    v <- as.numeric(value)
    keep <- switch(criteria,
      # a text cell is not a number, so it is never equal to one -- which makes
      # "!=" true for it, and every other comparison false
      "==" = !is.na(n) & n == v, "!=" = is.na(n) | n != v,
      ">"  = !is.na(n) & n >  v, "<"  = !is.na(n) & n <  v,
      ">=" = !is.na(n) & n >= v, "<=" = !is.na(n) & n <= v)
  } else {
    if (criteria %in% .FILTER_NUMERIC_CRITERIA)
      stop(sprintf(paste0("`criteria = \"%s\"` with the text value \"%s\": ",
                          "writexl cannot predict how Excel orders text, so it ",
                          "will not guess which rows to hide. Compare against ",
                          "a number instead."), criteria, value), call. = FALSE)
    m <- !is.na(cells$txt) & .filter_text_match(cells$txt, value)
    keep <- if (criteria == "==") m else !m
  }
  keep & !cells$blank
}

.filter_keep_rules <- function(cells, p) {
  if (!.filter_is_custom(p)) {
    # a value list, possibly of two values, plus the blanks flag
    vals <- c(if (!identical(p$criteria, "blanks")) p$value,
              if (!is.null(p$criteria2) && !identical(p$criteria2, "blanks"))
                p$value2)
    keep <- if (length(vals)) .filter_display_match(cells, vals)
            else rep(FALSE, length(cells$blank))
    if (identical(p$criteria, "blanks") || identical(p$criteria2, "blanks"))
      keep <- keep | cells$blank
    return(keep)
  }
  keep <- .filter_custom_keep(cells, p$criteria, p$value)
  if (!is.null(p$criteria2)) {
    k2 <- .filter_custom_keep(cells, p$criteria2, p$value2)
    keep <- if (identical(p$and_or, "or")) keep | k2 else keep & k2
  }
  keep
}

# A value list matches displayed text, and writexl cannot know how Excel will
# format a date, so that combination is refused rather than guessed at.
.filter_check_dated <- function(cells, nm) {
  if (!cells$dated) return(invisible(NULL))
  stop(sprintf(paste0("cannot filter column \"%s\" by an exact value or a ",
                      "`list`: it holds dates, and that form of filter matches ",
                      "the text Excel displays, which depends on the number ",
                      "format. Use a comparison such as `\">=\"` instead."),
               nm), call. = FALSE)
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
#' matching rules, which were measured in Excel rather than assumed.  Those
#' rules depend on which of two forms the filter takes:
#'
#' * `"=="` (with no wildcard) and `"blanks"`, and any `list`, are written as a
#'   **value list**.  Excel matches these against the text a cell *displays*,
#'   case-insensitively --- so `"=="` with `10`, with `"10"`, or a `list` of
#'   `"10"` all match both the number `10` and the string `"10"`.
#' * every other criteria, including `"=="` with a `*` or `?` in it, is written
#'   as a **typed comparison**.  If the value is a number the comparison is
#'   numeric and a text cell never satisfies it (except `"!="`, which a text
#'   cell satisfies because it is not that number).  If the value is text the
#'   comparison is textual, case-insensitive, with `*` and `?` as wildcards,
#'   and a number cell never satisfies it.
#'
#' The consequence worth knowing is that `"=="` with `"10"` keeps the number
#' `10`, while `"=="` with `"1*"` keeps nothing on a numeric column: the
#' wildcard changes the form, and so the rule.
#'
#' Blank means an empty cell or an empty string; `"non-blanks"` is its exact
#' complement.  A mixed-type column built with [xl_cell_general()] is matched
#' cell by cell, each by its own type.
#'
#' @section Limitations:
#' A value list matches displayed text, which writexl can only predict for the
#' General format.  Filtering a `Date` or `POSIXct` column by `"=="` or `list`
#' is therefore refused --- use a comparison such as `">="`.  For the same
#' reason a numeric column carrying a custom number format (say two decimal
#' places, or a currency symbol) may display differently from what writexl
#' compares, so prefer a comparison there too.  Columns whose cells hold
#' formulas cannot be filtered at all, since writexl does not know what Excel
#' would compute.
#'
#' @param col The column to filter: a name or a 1-based position.
#' @param criteria One of `"=="`, `"!="`, `">"`, `"<"`, `">="`, `"<="`,
#'   `"blanks"` or `"non-blanks"`.  The four magnitude comparisons need a
#'   numeric value; writexl will not guess how Excel orders text.
#' @param value The value to compare against.  In a text comparison `*` matches
#'   any run of characters and `?` any single one.
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
#' xl_filter("fruit", "==", "ap*")               # wildcard: typed comparison
#' xl_filter("fruit", list = c("apple", "banana"))
#' xl_filter("qty", "==", "10")                  # value list: matches the
#'                                               # number 10 and the text "10"
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
  nm <- names(df)[idx]
  cells <- .filter_cells(df[[idx]], nm)
  if (!is.null(p$list)) {
    .filter_check_dated(cells, nm)
    return(.filter_display_match(cells, p$list))
  }
  if (!.filter_is_custom(p)) .filter_check_dated(cells, nm)
  .filter_keep_rules(cells, p)
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

# =============================================================================
# Range addressing: one resolver for every range-scoped worksheet feature
# =============================================================================
#
# `.xl_resolve_range()` is the single place where a user-supplied range is
# turned into the 0-based `c(first_row, first_col, last_row, last_col)` quad
# that C consumes.  Everything range-scoped goes through it -- the autofilter
# today, and merged cells / data validation / conditional formats / images /
# charts / tables as they arrive -- so there is exactly one set of accepted
# spellings and one set of error messages.
#
# Two spellings are accepted:
#
#   * an A1-style string:
#       "A1:D51"        a rectangle
#       "A2"            a single cell (a zero-span rectangle)
#       "$A$1:$D$51"    absolute references ($ is ignored)
#       "B:D"           whole columns, spanning the sheet's used rows
#       "2:10"          whole rows, spanning the sheet's used columns
#   * a data-frame-relative spec, `list(rows = , cols = )`, where `cols` names
#     or indexes columns of the sheet's data frame and `rows` gives 1-based
#     *data* row indices (row 1 is the first data row, ignoring the header).
#
# The header offset is applied here, so C never has to know whether a header
# row was written.
# -----------------------------------------------------------------------------

# The xlsx grid limits (mirroring LXW_ROW_MAX / LXW_COL_MAX in libxlsxwriter).
.XLSX_MAX_ROWS <- 1048576L
.XLSX_MAX_COLS <- 16384L

# A1 notation, with the `$` of an absolute reference optional and ignored.
.RE_CELL_REF <- "^[$]?([A-Za-z]{1,3})[$]?([0-9]+)$"
.RE_COL_REF  <- "^[$]?([A-Za-z]{1,3})$"
.RE_ROW_REF  <- "^[$]?([0-9]+)$"

# Convert column letters ("A", "AB") to a 0-based column index.
.col_letters_to_index <- function(letters_str) {
  chars <- utf8ToInt(toupper(letters_str)) - utf8ToInt("A") + 1L
  idx <- 0L
  for (v in chars) idx <- idx * 26L + v
  idx - 1L
}

# Validate a 1-based sheet row number and return it 0-based.
.check_row_number <- function(n, arg) {
  n <- suppressWarnings(as.numeric(n))
  if (is.na(n) || n < 1 || n > .XLSX_MAX_ROWS)
    stop(sprintf("`%s` row must be between 1 and %d", arg, .XLSX_MAX_ROWS),
         call. = FALSE)
  as.integer(n) - 1L
}

# Validate an already 0-based column index.
.check_col_index <- function(i, arg) {
  if (is.na(i) || i < 0L || i >= .XLSX_MAX_COLS)
    stop(sprintf("`%s` column is outside the xlsx grid (maximum column is XFD)",
                 arg), call. = FALSE)
  as.integer(i)
}

# Parse one A1-style cell reference ("A2", "$A$2") to 0-based c(row, col).
.parse_cell_ref <- function(s, arg) {
  m <- regmatches(s, regexec(.RE_CELL_REF, s))[[1]]
  if (length(m) != 3L)
    stop(sprintf('`%s` must be a cell reference like "A2"', arg), call. = FALSE)
  c(.check_row_number(m[3L], arg),
    .check_col_index(.col_letters_to_index(m[2L]), arg))
}

# The rows a whole-column reference ("B:D") spans: the sheet's used rows when a
# data frame is available, otherwise the whole grid.
.sheet_row_extent <- function(df, header_offset, arg) {
  if (is.null(df)) return(c(0L, .XLSX_MAX_ROWS - 1L))
  last <- as.integer(header_offset) + nrow(df) - 1L
  if (last < 0L)
    stop(sprintf("`%s` covers no rows: the sheet is empty", arg), call. = FALSE)
  c(0L, as.integer(last))
}

# The columns a whole-row reference ("2:10") spans.
.sheet_col_extent <- function(df, arg) {
  if (is.null(df)) return(c(0L, .XLSX_MAX_COLS - 1L))
  if (length(df) < 1L)
    stop(sprintf("`%s` covers no columns: the sheet has no columns", arg),
         call. = FALSE)
  c(0L, as.integer(length(df) - 1L))
}

# Reject an inverted or off-grid quad.  A zero-span range (a single cell) is
# legitimate and passes.
.check_range_span <- function(v, arg) {
  # Unreachable from today's callers -- every path into here has already been
  # through .check_row_number() / .check_col_index(), which reject anything
  # below the top-left cell.  Kept as a guard for the kinds still to come, so a
  # new producer of a quad cannot leak a negative index into C.
  # nocov start
  if (v[1L] < 0L || v[2L] < 0L)
    stop(sprintf("`%s` range starts before the top-left cell", arg),
         call. = FALSE)
  # nocov end
  if (v[3L] < v[1L] || v[4L] < v[2L])
    stop(sprintf(paste0("`%s` range is inverted: the last cell must not come ",
                        "before the first"), arg), call. = FALSE)
  as.integer(v)
}

# Resolve an A1-style range string.
.resolve_range_string <- function(rng, arg, df, header_offset, allow_cell) {
  malformed <- function()
    stop(sprintf('`%s` range must look like "A1:D51" (got "%s")', arg, rng),
         call. = FALSE)
  has_colon <- grepl(":", rng, fixed = TRUE)
  parts <- strsplit(rng, ":", fixed = TRUE)[[1L]]
  if (!has_colon) {
    # A bare cell is a zero-span range, but only where one makes sense.
    if (!allow_cell || !grepl(.RE_CELL_REF, parts[1L])) malformed()
    rc <- .parse_cell_ref(parts[1L], arg)
    return(c(rc[1L], rc[2L], rc[1L], rc[2L]))
  }
  if (length(parts) != 2L || !all(nzchar(parts))) malformed()
  a <- parts[1L]
  b <- parts[2L]
  if (grepl(.RE_CELL_REF, a) && grepl(.RE_CELL_REF, b)) {
    ra <- .parse_cell_ref(a, arg)
    rb <- .parse_cell_ref(b, arg)
    return(.check_range_span(c(ra[1L], ra[2L], rb[1L], rb[2L]), arg))
  }
  if (grepl(.RE_COL_REF, a) && grepl(.RE_COL_REF, b)) {
    ext <- .sheet_row_extent(df, header_offset, arg)
    return(.check_range_span(c(
      ext[1L], .check_col_index(.col_letters_to_index(sub("^[$]", "", a)), arg),
      ext[2L], .check_col_index(.col_letters_to_index(sub("^[$]", "", b)), arg)
    ), arg))
  }
  if (grepl(.RE_ROW_REF, a) && grepl(.RE_ROW_REF, b)) {
    ext <- .sheet_col_extent(df, arg)
    return(.check_range_span(c(
      .check_row_number(sub("^[$]", "", a), arg), ext[1L],
      .check_row_number(sub("^[$]", "", b), arg), ext[2L]
    ), arg))
  }
  malformed()
}

# The elements a data-frame-relative range spec may carry.
.RANGE_SPEC_FIELDS <- c("rows", "cols")

# Turn a 1-based index vector into the c(first, last) it spans, refusing a
# non-contiguous selection: a range is a rectangle, and quietly filling the
# holes would select cells the caller did not ask for.
.contiguous_span <- function(idx, arg, what) {
  u <- sort(unique(as.integer(idx)))
  if (length(u) > 1L && any(diff(u) != 1L))
    stop(sprintf("`%s` %s must select a contiguous block", arg, what),
         call. = FALSE)
  c(u[1L], u[length(u)])
}

# Resolve list(rows = , cols = ) against the sheet's data frame.
.resolve_range_spec <- function(spec, arg, df, header_offset) {
  nms <- names(spec)
  if (!length(spec) || is.null(nms) || !all(nzchar(nms)))
    stop(sprintf("`%s` list must be fully named, e.g. list(rows = 1:3, cols = \"a\")",
                 arg), call. = FALSE)
  if (anyDuplicated(nms))
    stop(sprintf("`%s` list has duplicated element(s): %s", arg,
                 paste(unique(nms[duplicated(nms)]), collapse = ", ")),
         call. = FALSE)
  unknown <- setdiff(nms, .RANGE_SPEC_FIELDS)
  if (length(unknown))
    stop(sprintf("unknown `%s` element(s): %s", arg,
                 paste(unknown, collapse = ", ")), call. = FALSE)
  if (is.null(df))
    stop(sprintf("`%s` given as list(rows = , cols = ) needs the sheet's data frame",
                 arg), call. = FALSE)

  if (is.null(spec$cols)) {
    if (length(df) < 1L)
      stop(sprintf("`%s` covers no columns: the sheet has no columns", arg),
           call. = FALSE)
    cols <- c(1L, length(df))
  } else {
    cols <- .contiguous_span(.resolve_col_index(spec$cols, names(df), arg),
                             arg, "cols")
  }

  if (is.null(spec$rows)) {
    if (nrow(df) < 1L)
      stop(sprintf("`%s` covers no rows: the sheet has no data rows", arg),
           call. = FALSE)
    rows <- c(1L, nrow(df))
  } else {
    r <- spec$rows
    if (!is.numeric(r) || !length(r) || anyNA(r))
      stop(sprintf("`%s` rows must be a numeric vector of 1-based data-row indices",
                   arg), call. = FALSE)
    if (any(r < 1))
      stop(sprintf("`%s` rows must be positive (1-based data-row indices)", arg),
           call. = FALSE)
    rows <- .contiguous_span(r, arg, "rows")
  }

  first_row <- .check_row_number(rows[1L] + header_offset, arg)
  last_row  <- .check_row_number(rows[2L] + header_offset, arg)
  .check_range_span(c(first_row, cols[1L] - 1L, last_row, cols[2L] - 1L), arg)
}

# Resolve any accepted range spelling to 0-based
# c(first_row, first_col, last_row, last_col).
#
# `arg` names the calling argument in every error message.  `df` and
# `header_offset` supply the sheet context needed by the data-frame-relative
# and whole-column/whole-row forms.  `allow_cell = FALSE` rejects a bare cell
# reference for arguments documented as taking a rectangle.
.xl_resolve_range <- function(x, arg = "range", df = NULL, header_offset = 1L,
                              allow_cell = TRUE) {
  if (is.list(x) && !is.data.frame(x))
    return(.resolve_range_spec(x, arg, df, header_offset))
  if (is.character(x) && length(x) == 1L && !is.na(x))
    return(.resolve_range_string(x, arg, df, header_offset, allow_cell))
  stop(sprintf(paste0('`%s` must be an Excel range string like "A1:D51" or a ',
                      "list(rows = , cols = ) spec"), arg), call. = FALSE)
}

# Parse an Excel range ("A1:D51") into 0-based c(first_row, first_col,
# last_row, last_col).  A thin wrapper over the shared resolver that keeps the
# rectangle-only contract of the arguments documented as taking a *range*.
.parse_range <- function(rng, arg = "autofilter", df = NULL,
                         header_offset = 1L) {
  .xl_resolve_range(rng, arg = arg, df = df, header_offset = header_offset,
                    allow_cell = FALSE)
}

# Parse a freeze specification into c(freeze_row, freeze_col) (counts), or
# c(-1, -1) for none.
#
# Note the different reading of `list(row =, col =)` here: for `freeze` these
# are *counts* of rows/columns to freeze, not indices, so only the cell
# reference form shares the range machinery.
.parse_freeze <- function(freeze) {
  if (is.null(freeze) || (length(freeze) == 1L && is.na(freeze)))
    return(c(-1L, -1L))
  if (is.list(freeze)) {
    r <- if (!is.null(freeze$row)) as.integer(freeze$row) else 0L
    cc <- if (!is.null(freeze$col)) as.integer(freeze$col) else 0L
    return(c(r, cc))
  }
  if (is.character(freeze) && length(freeze) == 1L)
    return(.parse_cell_ref(freeze, "freeze"))
  stop('`freeze` must be a cell reference ("A2") or list(row =, col =)',
       call. = FALSE)
}

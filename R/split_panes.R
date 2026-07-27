# =============================================================================
# Split panes: converting a cell reference into libxlsxwriter's units
# =============================================================================
#
# worksheet_split_panes() does not take a row and a column.  Its two arguments
# are distances, measured in the units Excel uses for row height and column
# width, and the two units differ from each other:
#
#     worksheet_split_panes(worksheet1, 15, 0);    // First row.
#     worksheet_split_panes(worksheet2, 0,  8.43); // First column.
#
# So `vertical` is a number of row-heights and positions the split *between
# rows*, while `horizontal` is a number of column-widths and positions it
# between columns.  Passing 1 for "one row" would put the split a fifteenth of
# the way down the first row, silently.
#
# writexl therefore takes a cell reference, as `freeze` does, and converts:
# the distance above the split is the summed height of the rows before it, and
# the distance to its left is the summed width of the columns before it.  Those
# use the sheet's real geometry -- header row height, xl_row_spec() heights,
# xl_col_spec() widths and auto_colwidth results -- rather than assuming the
# defaults, so a split lands where it was asked for on a sheet whose rows or
# columns have been resized.
#
# A raw list(vertical =, horizontal =) is still accepted for callers who want
# to specify the units directly.
# -----------------------------------------------------------------------------

# Excel's defaults, and the values libxlsxwriter's own examples use.
.DEFAULT_ROW_HEIGHT <- 15
.DEFAULT_COL_WIDTH  <- 8.43

# The height of one 0-based sheet row: an explicit xl_row_spec() height, else
# the sheet default, else Excel's.
.row_height_at <- function(i, header_offset, props, default_row_height,
                           row_row, row_height) {
  k <- which(row_row == i)
  if (length(k) && !is.na(row_height[k[1L]])) return(row_height[k[1L]])
  if (header_offset > 0L && i == 0L && !is.null(props$header_row_height))
    return(as.numeric(props$header_row_height))
  if (!is.na(default_row_height)) return(as.numeric(default_row_height))
  .DEFAULT_ROW_HEIGHT
}

# The width of one 1-based column.
.col_width_at <- function(j, col_width) {
  if (j <= length(col_width) && !is.na(col_width[j])) col_width[j]
  else .DEFAULT_COL_WIDTH
}

.resolve_split <- function(spec, df, header_offset, props, default_row_height,
                           row_row, row_height, col_width) {
  # raw units, for callers who want to place the split exactly
  if (is.list(spec) &&
      (!is.null(spec[["vertical"]]) || !is.null(spec[["horizontal"]]))) {
    bad <- setdiff(names(spec), c("vertical", "horizontal"))
    if (length(bad))
      stop("unknown `split` element(s): ", paste(bad, collapse = ", "),
           call. = FALSE)
    v <- if (is.null(spec$vertical)) 0 else as.numeric(spec$vertical)
    h <- if (is.null(spec$horizontal)) 0 else as.numeric(spec$horizontal)
    if (anyNA(c(v, h)) || any(c(v, h) < 0))
      stop("`split` units must be non-negative numbers", call. = FALSE)
    return(c(vertical = v, horizontal = h))
  }

  q <- .xl_resolve_range(spec, arg = "split", df = df,
                         header_offset = header_offset, allow_cell = TRUE)
  n_rows <- q[1L]   # 0-based row index == how many rows sit above the split
  n_cols <- q[2L]

  vertical <- 0
  if (n_rows > 0L)
    vertical <- sum(vapply(seq_len(n_rows) - 1L, .row_height_at, numeric(1),
                           header_offset = header_offset, props = props,
                           default_row_height = default_row_height,
                           row_row = row_row, row_height = row_height))
  horizontal <- 0
  if (n_cols > 0L)
    horizontal <- sum(vapply(seq_len(n_cols), .col_width_at, numeric(1),
                             col_width = col_width))

  c(vertical = vertical, horizontal = horizontal)
}

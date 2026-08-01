# Tests for the shared range resolver in R/xl_range.R.

df3 <- data.frame(a = 1:3, b = 4:6, c = 7:9)

test_that("column letters convert to 0-based indices", {
  expect_equal(.col_letters_to_index("A"), 0L)
  expect_equal(.col_letters_to_index("b"), 1L)
  expect_equal(.col_letters_to_index("Z"), 25L)
  expect_equal(.col_letters_to_index("AA"), 26L)
  expect_equal(.col_letters_to_index("XFD"), 16383L)   # the last xlsx column
})

test_that("cell references parse, including absolute form", {
  expect_equal(.parse_cell_ref("A1", "x"), c(0L, 0L))
  expect_equal(.parse_cell_ref("a2", "x"), c(1L, 0L))
  expect_equal(.parse_cell_ref("D51", "x"), c(50L, 3L))
  expect_equal(.parse_cell_ref("$D$51", "x"), c(50L, 3L))
  expect_equal(.parse_cell_ref("$D51", "x"), c(50L, 3L))
  expect_error(.parse_cell_ref("nonsense", "x"), "cell reference")
  expect_error(.parse_cell_ref("1A", "x"), "cell reference")
  expect_error(.parse_cell_ref("", "x"), "cell reference")
})

test_that("grid limits are enforced on cell references", {
  expect_error(.parse_cell_ref("A0", "x"), "row must be between 1")
  expect_error(.parse_cell_ref("A1048577", "x"), "row must be between 1")
  expect_equal(.parse_cell_ref("XFD1048576", "x"), c(1048575L, 16383L))
  expect_error(.parse_cell_ref("XFE1", "x"), "outside the xlsx grid")
})

test_that("A1 rectangles resolve, absolute or not", {
  expect_equal(.xl_resolve_range("A1:D51", "x"), c(0L, 0L, 50L, 3L))
  expect_equal(.xl_resolve_range("$A$1:$D$51", "x"), c(0L, 0L, 50L, 3L))
  expect_equal(.xl_resolve_range("B2:B2", "x"), c(1L, 1L, 1L, 1L))  # zero span
})

test_that("a bare cell is a zero-span range where one is allowed", {
  expect_equal(.xl_resolve_range("A2", "x"), c(1L, 0L, 1L, 0L))
  expect_equal(.xl_resolve_range("$C$4", "x"), c(3L, 2L, 3L, 2L))
  # ... and is rejected where the argument is documented as taking a rectangle
  expect_error(.xl_resolve_range("A2", "x", allow_cell = FALSE), "A1:D51")
})

test_that("whole-column references span the sheet's used rows", {
  expect_equal(.xl_resolve_range("B:D", "x", df3, 1L), c(0L, 1L, 3L, 3L))
  expect_equal(.xl_resolve_range("B:D", "x", df3, 0L), c(0L, 1L, 2L, 3L))
  expect_equal(.xl_resolve_range("$B:$B", "x", df3, 1L), c(0L, 1L, 3L, 1L))
  # with no data frame in hand, the whole grid
  expect_equal(.xl_resolve_range("B:D", "x"), c(0L, 1L, 1048575L, 3L))
})

test_that("whole-row references span the sheet's used columns", {
  expect_equal(.xl_resolve_range("2:10", "x", df3, 1L), c(1L, 0L, 9L, 2L))
  expect_equal(.xl_resolve_range("$2:$2", "x", df3, 1L), c(1L, 0L, 1L, 2L))
  expect_equal(.xl_resolve_range("2:10", "x"), c(1L, 0L, 9L, 16383L))
})

test_that("malformed range strings are rejected and name the argument", {
  for (bad in c("A1:", ":", ":B2", "A1::B2", "A1:9", "9:A1", "A1:B", "",
                "not a range", "A1:B2:C3")) {
    expect_error(.xl_resolve_range(bad, "merge"), "`merge` range must look like",
                 fixed = TRUE)
  }
})

test_that("non-string, non-list ranges are rejected", {
  expect_error(.xl_resolve_range(42, "merge"), "`merge` must be an Excel range",
               fixed = TRUE)
  expect_error(.xl_resolve_range(NA_character_, "merge"), "must be an Excel range")
  expect_error(.xl_resolve_range(c("A1:B2", "C1:C2"), "merge"),
               "must be an Excel range")
})

test_that("inverted ranges are an error, zero-span ones are not", {
  expect_error(.xl_resolve_range("D1:A1", "merge"), "inverted")
  expect_error(.xl_resolve_range("A5:A2", "merge"), "inverted")
  expect_error(.xl_resolve_range("D:A", "merge", df3, 1L), "inverted")
  expect_error(.xl_resolve_range("10:2", "merge", df3, 1L), "inverted")
  expect_equal(.xl_resolve_range("A1:A1", "merge"), c(0L, 0L, 0L, 0L))
})

test_that("data-frame-relative specs resolve rows and columns", {
  expect_equal(.xl_resolve_range(list(rows = 1:2, cols = c("a", "b")), "m", df3, 1L),
               c(1L, 0L, 2L, 1L))
  expect_equal(.xl_resolve_range(list(rows = 1:2, cols = c("a", "b")), "m", df3, 0L),
               c(0L, 0L, 1L, 1L))
  expect_equal(.xl_resolve_range(list(rows = 2, cols = 2), "m", df3, 1L),
               c(2L, 1L, 2L, 1L))
  # numeric column positions and unsorted input both work
  expect_equal(.xl_resolve_range(list(rows = c(3, 2), cols = c(3, 2)), "m", df3, 1L),
               c(2L, 1L, 3L, 2L))
})

test_that("an omitted rows/cols element means the whole data block", {
  expect_equal(.xl_resolve_range(list(cols = "b"), "m", df3, 1L), c(1L, 1L, 3L, 1L))
  expect_equal(.xl_resolve_range(list(cols = "b"), "m", df3, 0L), c(0L, 1L, 2L, 1L))
  expect_equal(.xl_resolve_range(list(rows = 1), "m", df3, 1L), c(1L, 0L, 1L, 2L))
})

test_that("range specs are validated", {
  expect_error(.xl_resolve_range(list(), "m", df3, 1L), "must be fully named")
  expect_error(.xl_resolve_range(list(1:2), "m", df3, 1L), "must be fully named")
  expect_error(.xl_resolve_range(list(rows = 1, rows = 2), "m", df3, 1L),
               "duplicated element")
  expect_error(.xl_resolve_range(list(row = 1), "m", df3, 1L),
               "unknown `m` element", fixed = TRUE)
  expect_error(.xl_resolve_range(list(cols = "a"), "m"), "needs the sheet's data frame")
  expect_error(.xl_resolve_range(list(cols = "z"), "m", df3, 1L), "unknown column")
  expect_error(.xl_resolve_range(list(cols = 9), "m", df3, 1L), "out of range")
  expect_error(.xl_resolve_range(list(rows = 0), "m", df3, 1L), "must be positive")
  expect_error(.xl_resolve_range(list(rows = "a"), "m", df3, 1L), "must be a numeric")
  expect_error(.xl_resolve_range(list(rows = numeric(0)), "m", df3, 1L),
               "must be a numeric")
  expect_error(.xl_resolve_range(list(rows = c(1, NA)), "m", df3, 1L),
               "must be a numeric")
})

test_that("a range must be a contiguous rectangle", {
  expect_error(.xl_resolve_range(list(cols = c("a", "c")), "m", df3, 1L),
               "cols must select a contiguous block")
  expect_error(.xl_resolve_range(list(rows = c(1, 3)), "m", df3, 1L),
               "rows must select a contiguous block")
})

test_that("degenerate sheets are reported, not silently accepted", {
  empty_rows <- data.frame(a = integer(0), b = integer(0))
  no_cols <- data.frame(a = 1:2)[, character(0), drop = FALSE]
  expect_error(.xl_resolve_range(list(cols = "a"), "m", empty_rows, 1L),
               "covers no rows")
  expect_error(.xl_resolve_range(list(rows = 1), "m", no_cols, 1L),
               "covers no columns")
  expect_error(.xl_resolve_range("1:2", "m", no_cols, 1L), "covers no columns")
  # with a header row a 0-row sheet still has one usable row
  expect_equal(.xl_resolve_range("A:A", "m", empty_rows, 1L), c(0L, 0L, 0L, 0L))
  expect_error(.xl_resolve_range("A:A", "m", empty_rows, 0L), "covers no rows")
})

test_that("rows beyond the xlsx grid are rejected", {
  expect_error(.xl_resolve_range(list(rows = 1048576), "m", df3, 1L),
               "row must be between 1")
})

test_that(".parse_range keeps its rectangle-only contract", {
  expect_equal(.parse_range("A1:D51"), c(0L, 0L, 50L, 3L))
  expect_error(.parse_range("A1"), "`autofilter` range must look like",
               fixed = TRUE)
  expect_error(.parse_range("A1:B2", "merge"), NA)
  expect_error(.parse_range("A1", "merge"), "`merge` range must look like",
               fixed = TRUE)
})

test_that("freeze specs reuse the cell-reference parser but keep count semantics", {
  expect_equal(.parse_freeze("$B$3"), c(2L, 1L))
  # list(row =, col =) here means counts of frozen rows/columns, not indices
  expect_equal(.parse_freeze(list(row = 2, col = 1)), c(2L, 1L))
  expect_equal(.parse_freeze(list(row = 3)), c(3L, 0L))
  expect_equal(.parse_freeze(list(col = 3)), c(0L, 3L))
  expect_equal(.parse_freeze(NA_character_), c(-1L, -1L))
})

# ── Sheet-qualified ranges ───────────────────────────────────────────────────

test_that("a quoted sheet name is split off, apostrophes and all", {
  # Excel quotes a sheet name holding spaces or punctuation and doubles an
  # apostrophe inside it
  expect_equal(.split_sheet_ref("Data!A1:B5"),
               list(sheet = "Data", rest = "A1:B5"))
  expect_equal(.split_sheet_ref("'My Sheet'!A1:B5"),
               list(sheet = "My Sheet", rest = "A1:B5"))
  expect_equal(.split_sheet_ref("'It''s'!A1"),
               list(sheet = "It's", rest = "A1"))
  expect_equal(.split_sheet_ref("A1:B5"), list(sheet = NULL, rest = "A1:B5"))
  # an opening quote with no closing one is not a sheet reference
  expect_equal(.split_sheet_ref("'unterminated"),
               list(sheet = NULL, rest = "'unterminated"))
})

test_that("a sheet name is refused where a range applies to one sheet only", {
  df <- data.frame(a = 1:3, b = 4:6)
  expect_error(.xl_resolve_range("Other!A1:B2", "range", df),
               "may not name a sheet")
  expect_error(.xl_resolve_range(list(sheet = "Other", cols = "a"), "range", df),
               "may not name a sheet")
  # and carried as an attribute where it is allowed
  got <- .xl_resolve_range(list(sheet = "Other", cols = "a"), "range", df,
                           allow_sheet = TRUE)
  expect_equal(attr(got, "sheet"), "Other")
  got <- .xl_resolve_range("'My Sheet'!A1:B2", "range", df, allow_sheet = TRUE)
  expect_equal(attr(got, "sheet"), "My Sheet")
})

test_that("a sheet name in a spec must be a single string", {
  df <- data.frame(a = 1:3)
  for (bad in list(1, c("a", "b"), NA_character_))
    expect_error(.xl_resolve_range(list(sheet = bad, cols = "a"), "range", df,
                                   allow_sheet = TRUE),
                 "must be a single sheet name")
})

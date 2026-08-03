# What gets written where a value has none.
#
# writexl has always left such a cell blank.  `na` substitutes something else,
# at three scopes -- a cell, a column, the workbook -- with the innermost one
# that is set winning.  The default, `na = NA`, must keep the blank, and the
# first test proves that byte for byte rather than by inspection: this is the
# whole compatibility story for a package whose reverse dependencies all write
# data frames with NAs in them.

na_df <- function() data.frame(
  num  = c(1.5, NA, NaN),
  int  = c(1L, NA, 3L),
  chr  = c("a", NA, "c"),
  lgl  = c(TRUE, NA, FALSE),
  date = as.Date(c("2024-01-01", NA, "2024-03-01")),
  time = as.POSIXct(c("2024-01-01 10:00:00", NA, "2024-03-01 12:00:00"),
                    tz = "UTC"),
  stringsAsFactors = FALSE
)

sheet_xml_of <- function(...) {
  p <- tempfile(fileext = ".xlsx")
  write_xlsx(..., path = p)
  xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE)
}

# A string cell holds an index into sharedStrings.xml, so counting occurrences
# of a substitute means counting them there -- once per distinct string, which
# is why these tests count cells that point at it rather than the string.
shared_has <- function(p, s)
  grepl(s, xlsx_part(p, "xl/sharedStrings.xml", raw = TRUE), fixed = TRUE)

# ── The default changes nothing ───────────────────────────────────────────────

test_that("na = NA writes exactly what writing no `na` at all writes", {
  df <- na_df()
  # every scope explicitly unset, which is what the defaults already are
  explicit <- sheet_xml_of(list(D = xl_sheet(
    df, cols = xl_col_spec("num", na = NA))), na = NA)
  implicit <- sheet_xml_of(list(D = df))
  expect_equal(explicit, implicit)
  # and a blank cell is still no cell: the row holds only the non-NA columns
  expect_false(grepl("<c r=\"A3\"", implicit, fixed = TRUE))
})

test_that("the whole file is unchanged by the default, not just one sheet", {
  # xl_properties() gained an argument; the properties payload must not have
  # grown a key that reaches the file when nothing asked for one.  `created` is
  # pinned so the only thing that could differ is the change under test rather
  # than the timestamp libxlsxwriter stamps on each run.
  df <- na_df()
  when <- as.POSIXct("2024-01-01 00:00:00", tz = "UTC")
  bytes <- function(...) {
    p <- write_tmp(xl_workbook(list(D = df),
                               properties = xl_properties(created = when, ...)))
    readBin(p, "raw", file.size(p))
  }
  expect_equal(bytes(na = NA), bytes())
})

# ── Every column type ─────────────────────────────────────────────────────────

test_that("a workbook-wide na fills in for every column type, NaN included", {
  p <- write_tmp(na_df(), na = "none")
  got <- as.data.frame(readxl::read_xlsx(p))
  # the substitute makes each column mixed, so readxl reports character
  for (col in names(got))
    expect_equal(got[[col]][2L], "none", info = col)
  # NaN is as unwritable as NA, so it is filled in too
  expect_equal(got$num[3L], "none")
  # and the values that do exist are untouched
  expect_equal(got$chr[1L], "a")
  expect_equal(got$int[3L], "3")
})

test_that("na keeps its own type rather than becoming a string", {
  p <- write_tmp(data.frame(x = c(1.5, NA)), na = 0)
  got <- as.data.frame(readxl::read_xlsx(p))
  # still a numeric column, because the substitute is a number
  expect_type(got$x, "double")
  expect_equal(got$x, c(1.5, 0))

  # a logical substitute stays a boolean cell
  x <- sheet_xml_of(list(D = data.frame(a = c("k", NA), stringsAsFactors = FALSE)),
                    na = TRUE)
  expect_match(x, 't="b"', fixed = TRUE)
})

test_that("Inf is unaffected: it has a representation already", {
  p <- write_tmp(data.frame(x = c(Inf, -Inf, NA)), na = "none")
  got <- as.data.frame(readxl::read_xlsx(p))
  expect_equal(got$x, c("Inf", "-Inf", "none"))
})

# ── Scopes ────────────────────────────────────────────────────────────────────

test_that("a column's na overrides the workbook's", {
  df <- data.frame(a = c(1, NA), b = c(1, NA), c = c(1, NA))
  p <- write_tmp(xl_workbook(
    list(D = xl_sheet(df, cols = list(xl_col_spec("a", na = -999),
                                      xl_col_spec("b", na = "gone")))),
    properties = xl_properties(na = "WB")))
  got <- as.data.frame(readxl::read_xlsx(p))
  # a number keeps the column numeric; a string makes it mixed
  expect_equal(got$a[2L], -999)     # column wins
  expect_equal(got$b[2L], "gone")   # column wins
  expect_equal(got$c[2L], "WB")     # no column setting, so the workbook's
})

test_that("a cell's na overrides the column's and the workbook's", {
  df <- data.frame(a = 1:3)
  df$mix <- xl_cell_general(value = list(1, NA, NA), na = c(NA, "CELL", NA))
  p <- write_tmp(xl_workbook(
    list(D = xl_sheet(df, cols = xl_col_spec("mix", na = "COL"))),
    properties = xl_properties(na = "WB")))
  got <- as.data.frame(readxl::read_xlsx(p))
  expect_equal(got$mix, c("1", "CELL", "COL"))
})

test_that("a cell reaches the workbook's na with no column between", {
  df <- data.frame(a = 1:2)
  df$mix <- xl_cell_general(value = list(1, NA))
  p <- write_tmp(xl_workbook(list(D = xl_sheet(df)),
                             properties = xl_properties(na = "WB")))
  expect_equal(as.data.frame(readxl::read_xlsx(p))$mix, c("1", "WB"))
})

# ── What na must not touch ────────────────────────────────────────────────────

test_that("a cell that has content keeps it", {
  df <- data.frame(a = 1:3)
  # a formula and a hyperlink are content, not missing values
  df$f <- xl_cell_general(formula = c("=A2+1", NA, NA))
  df$h <- xl_cell_general(hyperlink = c("https://example.com", NA, NA))
  p <- write_tmp(xl_workbook(list(D = xl_sheet(df)),
                             properties = xl_properties(na = "none")))
  x <- xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE)
  # the formula cell still holds its formula
  expect_match(x, "A2+1", fixed = TRUE)
  expect_true(shared_has(p, "none"))
  # the four cells that have neither formula nor value take the substitute,
  # and the two that have content do not
  got <- as.data.frame(readxl::read_xlsx(p))
  expect_equal(got$f[2:3], c("none", "none"))
  expect_equal(got$h[2:3], c("none", "none"))
  expect_false(identical(got$f[1L], "none"))
})

test_that("the header row is never substituted", {
  # a header is always written, so it has no missing value to stand in for --
  # this pins that `na` never reaches it
  df <- data.frame(none = c(1, NA))
  p <- write_tmp(list(D = df), na = "none")
  got <- as.data.frame(readxl::read_xlsx(p))
  expect_equal(names(got), "none")
  expect_equal(got$none, c("1", "none"))
})

test_that("a formatted blank keeps its format when substituted", {
  df <- data.frame(a = 1:2)
  df$b <- xl_cell_general(value = list(1, NA), format = xl_font(bold = TRUE),
                          na = "none")
  p <- write_tmp(list(D = xl_sheet(df)))
  x <- xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_true(shared_has(p, "none"))
  # the substituted cell carries the same style as the value cell above it,
  # which is the format the blank it replaces would have kept
  styles <- regmatches(x, gregexpr('<c r="B[23]" s="[0-9]+"', x))[[1L]]
  expect_length(styles, 2L)
  expect_equal(sub('.*s="', "", styles[1L]), sub('.*s="', "", styles[2L]))
  expect_match(xlsx_part(p, "xl/styles.xml", raw = TRUE), "<b/>", fixed = TRUE)
})

# ── Validation ────────────────────────────────────────────────────────────────

test_that("na must be a single value of some atomic type", {
  expect_error(xl_properties(na = c("a", "b")), "must be a single value")
  expect_error(xl_properties(na = list(1)), "must be a single value")
  expect_error(xl_col_spec("a", na = mean), "must be a single value")
  expect_error(xl_cell_general(value = 1, na = list(1, 2)),
               "must be length 1 or 1")
  expect_error(xl_cell_general(value = 1:3, na = c("a", "b")),
               "must be length 1 or 3")
  # a factor is taken as its label
  expect_equal(unclass(xl_properties(na = factor("gone")))$na, "gone")
  # NaN cannot stand in for a missing value either, so it reads as unset
  expect_null(unclass(xl_properties(na = NaN))$na)
  expect_null(unclass(xl_properties(na = NA))$na)
})

test_that("a per-cell na may leave some cells to inherit", {
  # NULL and NA both mean "nothing set here", so a list can mix cells that
  # override with cells that fall through to the column or the workbook
  df <- data.frame(a = 1:3)
  df$mix <- xl_cell_general(value = list(NA, NA, NA),
                            na = list("CELL", NULL, NA))
  p <- write_tmp(xl_workbook(list(D = xl_sheet(df)),
                             properties = xl_properties(na = "WB")))
  expect_equal(as.data.frame(readxl::read_xlsx(p))$mix,
               c("CELL", "WB", "WB"))
})

test_that("write_xlsx's na shorthand refuses to be silently ignored", {
  wb <- xl_workbook(list(D = data.frame(a = c(1, NA))))
  expect_error(write_tmp(wb, na = "none"), "already an xl_workbook")
  # without the shorthand the workbook's own setting applies
  expect_silent(write_tmp(wb))
})

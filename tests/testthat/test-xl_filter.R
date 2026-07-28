# Autofilter criteria, and the rows they hide.
#
# The expected results below are not derived from reading documentation: they
# were measured in Excel.  A workbook was written with criteria set and no rows
# hidden, opened, and Data > Reapply pressed so that Excel computed the match
# itself.  Each expectation records what Excel actually did, so a change in
# writexl's predicate that drifts away from Excel fails here.

# The sheet rows a filter hides (1-based, as they appear in the XML).
filter_hidden <- function(df, ...) {
  p <- write_tmp(list(D = xl_sheet(df, filter = list(...))))
  w <- xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE)
  rows <- regmatches(w, gregexpr('<row r="[0-9]+"[^>]*>', w))[[1]]
  n <- as.integer(sub('.*r="([0-9]+)".*', "\\1", rows))
  sort(n[grepl("hidden", rows)])
}

# Data row indices kept, derived from the predicate directly.
kept <- function(df, f) which(.filter_keep(f, df))

# ── Semantics measured in Excel ───────────────────────────────────────────────

test_that("text comparison is case-insensitive", {
  df <- data.frame(f = c("apple", "Apple", "APPLE", "banana"),
                   stringsAsFactors = FALSE)
  # Excel: apple, Apple and APPLE all stay visible
  expect_equal(kept(df, xl_filter("f", "==", "Apple")), 1:3)
  expect_equal(filter_hidden(df, xl_filter("f", "==", "Apple")), 5L)
})

test_that("* and ? are wildcards in a text value", {
  df <- data.frame(f = c("apple", "apricot", "banana", "ap"),
                   stringsAsFactors = FALSE)
  # Excel: apple, apricot and ap stay visible; banana goes
  expect_equal(kept(df, xl_filter("f", "==", "ap*")), c(1L, 2L, 4L))
  expect_equal(kept(df, xl_filter("f", "==", "ap?le")), 1L)
})

test_that("blank covers an empty cell and an empty string alike", {
  df <- data.frame(v = c("a", NA, "", "b"), stringsAsFactors = FALSE)
  # Excel: both middle rows stay visible under Blanks ...
  expect_equal(kept(df, xl_filter("v", "blanks")), 2:3)
  # ... and both disappear under Non-blanks, so the two partition the column
  expect_equal(kept(df, xl_filter("v", "non-blanks")), c(1L, 4L))
  expect_equal(filter_hidden(df, xl_filter("v", "blanks")), c(2L, 5L))
  expect_equal(filter_hidden(df, xl_filter("v", "non-blanks")), c(3L, 4L))
})

test_that("a value list matches displayed text, whatever the cell's type", {
  # "==" without a wildcard writes <filters>, which Excel matches against the
  # text a cell displays -- so a text criteria does match a number cell.  The R
  # type of `value` is irrelevant: only the string written to val matters.
  num <- data.frame(q = c(10, 20, 100, 15))
  expect_equal(kept(num, xl_filter("q", "==", "10")), 1L)
  expect_equal(kept(num, xl_filter("q", "==", 10)), 1L)
  expect_equal(kept(num, xl_filter("q", list = c("10", "100"))), c(1L, 3L))
  # logicals display as TRUE/FALSE and match the same way
  expect_equal(kept(data.frame(b = c(TRUE, FALSE, TRUE)),
                    xl_filter("b", list = "TRUE")), c(1L, 3L))
})

test_that("a wildcard turns the same criteria into a typed comparison", {
  # a * or ? forces <customFilters>, which compares by type -- and no number is
  # text, so the wildcard matches nothing at all on a numeric column
  num <- data.frame(q = c(10, 20, 100, 15))
  expect_equal(kept(num, xl_filter("q", "==", "1*")), integer(0))
  # the pair that shows the two forms really do differ: same column, same
  # criteria, and the wildcard is the only change
  expect_equal(kept(num, xl_filter("q", "==", "10")), 1L)
})

test_that("a text comparison leaves number cells alone", {
  # the mirror of the numeric case: a number is not text, so it never equals a
  # text value -- which makes "!=" true for it
  num <- data.frame(q = c(10, 20, 100))
  expect_equal(kept(num, xl_filter("q", "!=", "a*")), c(1L, 2L, 3L))
  txt <- data.frame(f = c("apple", "banana", "cherry"), stringsAsFactors = FALSE)
  expect_equal(kept(txt, xl_filter("f", "!=", "a*")), c(2L, 3L))
})

test_that("a blank survives != but no other comparison", {
  # a blank is not equal to anything, so "!=" keeps it.  This is the one place
  # blanks are not silently excluded, and it is easy to get backwards.
  df <- data.frame(v = c("a", NA, "b"), stringsAsFactors = FALSE)
  expect_equal(kept(df, xl_filter("v", "!=", "a")), c(2L, 3L))
  # every other comparison drops it
  n <- data.frame(q = c(5, NA, 300))
  expect_equal(kept(n, xl_filter("q", ">", 1)), c(1L, 3L))
  expect_equal(kept(n, xl_filter("q", "==", 5)), 1L)
  # "non-blanks" is written as != " " but must still exclude blanks
  expect_equal(kept(df, xl_filter("v", "non-blanks")), c(1L, 3L))
})

test_that("== or == collapses into a two-value list", {
  # libxlsxwriter writes this as <filters> with both values, not a custom
  # filter, so it matches displayed text
  df <- data.frame(f = c("apple", "banana", "cherry"), stringsAsFactors = FALSE)
  f <- xl_filter("f", "==", "apple", "==", "cherry", and_or = "or")
  expect_equal(kept(df, f), c(1L, 3L))
  w <- xlsx_part(write_tmp(list(D = xl_sheet(df, filter = f))),
                 "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_match(w, '<filter val="apple"/><filter val="cherry"/>', fixed = TRUE)
})

test_that("a value list matches a text cell that looks numeric", {
  # the mirror of the numeric column case: val="10" matches the string "10"
  # just as it matches the number 10
  df <- data.frame(t = c("10", "20", "abc"), stringsAsFactors = FALSE)
  expect_equal(kept(df, xl_filter("t", "==", 10)), 1L)
  expect_equal(kept(df, xl_filter("t", "==", "10")), 1L)
})

test_that("a custom filter compares numerically when its value is a number", {
  # val="10" is parsed as a number even when it was given as a string
  num <- data.frame(q = c(10, 20, 100, 15))
  expect_equal(kept(num, xl_filter("q", "!=", "10")), c(2L, 3L, 4L))
  # against a text column the same rule keeps everything: a text cell is not a
  # number, so it is never equal to one
  txt <- data.frame(t = c("10", "20", "abc"), stringsAsFactors = FALSE)
  expect_equal(kept(txt, xl_filter("t", "!=", 10)), c(1L, 2L, 3L))
  # and a magnitude comparison excludes every text cell
  expect_equal(kept(txt, xl_filter("t", ">", 15)), integer(0))
})

test_that("numeric comparisons behave as expected", {
  df <- data.frame(q = c(5, 150, 20, 300, 75))
  expect_equal(kept(df, xl_filter("q", ">", 100)), c(2L, 4L))
  expect_equal(kept(df, xl_filter("q", ">=", 75)), c(2L, 4L, 5L))
  expect_equal(kept(df, xl_filter("q", "<", 75)), c(1L, 3L))
  expect_equal(kept(df, xl_filter("q", "<=", 75)), c(1L, 3L, 5L))
  expect_equal(kept(df, xl_filter("q", "==", 150)), 2L)
  expect_equal(kept(df, xl_filter("q", "!=", 150)), c(1L, 3L, 4L, 5L))
  expect_equal(filter_hidden(df, xl_filter("q", ">", 100)), c(2L, 4L, 6L))
})

test_that("a mixed-type column is matched cell by cell", {
  # rows: 1 = number 10, 2 = text "20", 3 = text "abc", 4 = number 30,
  #       5 = number 20, 6 = empty string, 7 = NA
  mixed <- data.frame(row = 1:7)
  mixed$v <- xl_cell_general(value = list(10, "20", "abc", 30, 20, "", NA))

  # a magnitude comparison reaches the number cells only -- the text "20" in
  # row 2 is not a number, so it does not match even though it looks like one
  expect_equal(kept(mixed, xl_filter("v", ">", 15)), c(4L, 5L))
  # a text comparison reaches the text cells only
  expect_equal(kept(mixed, xl_filter("v", "==", "a*")), 3L)
  # both kinds of empty are blank
  expect_equal(kept(mixed, xl_filter("v", "blanks")), c(6L, 7L))
  expect_equal(kept(mixed, xl_filter("v", "non-blanks")), c(1L, 2L, 3L, 4L, 5L))
  # a value list matches displayed text, so val="20" catches the text "20" in
  # row 2 AND the number 20 in row 5
  expect_equal(kept(mixed, xl_filter("v", list = c("20", "abc"))),
               c(2L, 3L, 5L))
  expect_equal(filter_hidden(mixed, xl_filter("v", ">", 15)),
               c(2L, 3L, 4L, 7L, 8L))
})

test_that("a mixed column's rich strings and hyperlinks filter on their text", {
  df <- data.frame(row = 1:3)
  df$v <- xl_cell_general(value = list(
    xl_rich_string(xl_rich_run("ur", xl_font(bold = TRUE)), xl_rich_run("gent")),
    "calm", NA))
  expect_equal(kept(df, xl_filter("v", "==", "urgent")), 1L)

  h <- data.frame(row = 1:2)
  h$v <- xl_cell_general(hyperlink = c("https://a.example", "https://b.example"))
  expect_equal(kept(h, xl_filter("v", "==", "https://a.example")), 1L)
})

test_that("the R type of `value` does not change the file, so nor the match", {
  # == 10 and == "10" write byte-identical criteria, which is why they must
  # match identically; this pins the equivalence mechanically
  crit <- function(f) {
    w <- xlsx_part(write_tmp(list(D = xl_sheet(data.frame(q = c(10, 20)),
                                               filter = f))),
                   "xl/worksheets/sheet1.xml", raw = TRUE)
    regmatches(w, regexpr("<filterColumn.*</filterColumn>", w))
  }
  expect_identical(crit(xl_filter("q", "==", 10)),
                   crit(xl_filter("q", "==", "10")))
  expect_identical(crit(xl_filter("q", "!=", 10)),
                   crit(xl_filter("q", "!=", "10")))
})

test_that("list membership is case-insensitive too", {
  df <- data.frame(f = c("apple", "Apple", "banana", "BANANA"),
                   stringsAsFactors = FALSE)
  # Excel: all four stay visible for the list apple, BANANA
  expect_equal(kept(df, xl_filter("f", list = c("apple", "BANANA"))), 1:4)
  expect_equal(kept(df, xl_filter("f", list = "apple")), 1:2)
  expect_equal(filter_hidden(df, xl_filter("f", list = "apple")), c(4L, 5L))
})

test_that("two rules combine with and / or", {
  df <- data.frame(q = c(50, 120, 180, 250))
  expect_equal(kept(df, xl_filter("q", ">", 100, "<", 200)), 2:3)
  expect_equal(kept(df, xl_filter("q", "<", 100, ">", 200, and_or = "or")),
               c(1L, 4L))
})

test_that("dates filter on their underlying value", {
  df <- data.frame(when = as.Date("2024-01-01") + c(0, 91, 182, 274))
  expect_equal(kept(df, xl_filter("when", ">=", as.Date("2024-04-01"))), 2:4)
})

# ── The criteria reach the file ───────────────────────────────────────────────

test_that("criteria are written alongside the autofilter", {
  df <- data.frame(q = c(5, 150, 300))
  w <- xlsx_part(write_tmp(list(D = xl_sheet(df, filter = xl_filter("q", ">", 100)))),
                 "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_match(w, "<autoFilter ", fixed = TRUE)
  expect_match(w, 'operator="greaterThan"', fixed = TRUE)
  expect_match(w, 'val="100"', fixed = TRUE)
})

test_that("a filter implies the autofilter", {
  # criteria without an autofilter range is rejected by libxlsxwriter, and is
  # meaningless anyway
  df <- data.frame(q = c(5, 150))
  w <- xlsx_part(write_tmp(list(D = xl_sheet(df, filter = xl_filter("q", ">", 100)))),
                 "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_match(w, '<autoFilter ref="A1:A3">', fixed = TRUE)
})

test_that("a string filter is stored as a filter list", {
  # libxlsxwriter converts a string equality into <filters>, not <customFilter>
  df <- data.frame(f = c("apple", "banana"), stringsAsFactors = FALSE)
  w <- xlsx_part(write_tmp(list(D = xl_sheet(df, filter = xl_filter("f", "==", "apple")))),
                 "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_match(w, '<filter val="apple"/>', fixed = TRUE)
})

test_that("filters on two columns combine with AND", {
  df <- data.frame(q = c(5, 150, 300, 200),
                   f = c("a", "b", "a", "a"), stringsAsFactors = FALSE)
  h <- filter_hidden(df, xl_filter("q", ">", 100), xl_filter("f", "==", "a"))
  # rows 3 (300, "a") and 4 (200, "a") satisfy both; row 1 fails on q and
  # row 2 on f, so sheet rows 2 and 3 go
  expect_equal(h, c(2L, 3L))
  expect_equal(kept(df, xl_filter("q", ">", 100)), c(2L, 3L, 4L))
})

test_that("row hiding survives every special column type", {
  # the filtered column is plain, but the sheet's other columns go through the
  # xl_cell_general machinery; hiding is driven by the row plan, so it must not
  # matter what else the row contains
  base <- data.frame(q = c(5, 150, 300))
  f <- xl_filter("q", ">", 100)
  expect_equal(filter_hidden(base, f), 2L)

  with_col <- function(v) {
    d <- base
    d$extra <- v
    d
  }
  expect_equal(filter_hidden(with_col(xl_cell_general(
    value = c("a", "b", "c"), format = xl_font(bold = TRUE))), f), 2L)
  expect_equal(filter_hidden(with_col(
    xl_formula(c("=A2*2", "=A3*2", "=A4*2"))), f), 2L)
  expect_equal(filter_hidden(with_col(
    xl_cell_general(value = "x", comment = "hi")), f), 2L)
  expect_equal(filter_hidden(with_col(
    xl_hyperlink("https://example.com")), f), 2L)
  expect_equal(filter_hidden(with_col(xl_cell_general(value = list(
    xl_rich_string(xl_rich_run("bo", xl_font(bold = TRUE)), xl_rich_run("ld")),
    xl_rich_string(xl_rich_run("it", xl_font(italic = TRUE)), xl_rich_run("al")),
    xl_rich_string(xl_rich_run("pl"), xl_rich_run("ain"))))), f), 2L)
})

# ── Validation ────────────────────────────────────────────────────────────────

test_that("a filter on a formula column is refused, naming the row", {
  # writexl cannot know what Excel would compute, so it must not guess
  df <- data.frame(x = 1:2)
  df$f <- xl_cell_general(formula = c("=A1", "=A2"))
  expect_error(write_tmp(list(D = xl_sheet(df, filter = xl_filter("f", "==", 1)))),
               "row 1 holds a formula")
  # one formula among plain values is still enough to refuse
  d2 <- data.frame(x = 1:3)
  d2$f <- xl_cell_general(value = list(1, 2, NA),
                          formula = c(NA, NA, "=A2+A3"))
  expect_error(write_tmp(list(D = xl_sheet(d2, filter = xl_filter("f", ">", 1)))),
               "row 3 holds a formula")
})

test_that("a date cannot be filtered by exact value or list", {
  # <filters> matches displayed text, and the displayed text of a date depends
  # on the number format, so writexl refuses rather than guessing
  df <- data.frame(when = as.Date("2024-01-01") + 0:2)
  expect_error(write_tmp(list(D = xl_sheet(df,
    filter = xl_filter("when", list = "2024-01-01")))), "holds dates")
  expect_error(write_tmp(list(D = xl_sheet(df,
    filter = xl_filter("when", "==", as.Date("2024-01-01"))))), "holds dates")
  # a comparison is fine: it is a typed numeric comparison
  expect_equal(kept(df, xl_filter("when", ">=", as.Date("2024-01-02"))), 2:3)
})

test_that("an unorderable text comparison is refused, not guessed", {
  df <- data.frame(f = c("apple", "banana"), stringsAsFactors = FALSE)
  expect_error(kept(df, xl_filter("f", ">", "apple")),
               "cannot predict how Excel orders text")
})

test_that("filter arguments are validated", {
  expect_error(xl_filter(), "must name or index")
  expect_error(xl_filter("q"), "needs a `criteria` or a `list`")
  expect_error(xl_filter("q", ">"), "needs a `value`")
  expect_error(xl_filter("q", "~", 1), "criteria")
  expect_error(xl_filter("q", "==", 1, list = "a"), "not both")
  expect_error(xl_filter("q", list = character(0)), "character vector")
  expect_error(xl_filter("q", list = c("a", NA)), "character vector")
  expect_error(xl_filter("q", ">", 1, and_or = "xor"), "and_or")
})

test_that("an unknown column is reported", {
  df <- data.frame(q = 1:2)
  expect_error(write_tmp(list(D = xl_sheet(df, filter = xl_filter("nope", ">", 1)))),
               "unknown column")
})

test_that("a magnitude comparison needs a numeric column", {
  df <- data.frame(f = c("a", "b"), stringsAsFactors = FALSE)
  # no coercion, so this simply matches nothing rather than erroring
  expect_equal(kept(df, xl_filter("f", ">", 1)), integer(0))
})

test_that("non-filter objects are rejected", {
  expect_error(write_tmp(list(D = xl_sheet(data.frame(q = 1), filter = "q>1"))),
               "must be an xl_filter object")
})

test_that("the print method runs", {
  expect_output(print(xl_filter("q", ">", 100)), "xl_filter")
  expect_output(print(xl_filter("f", list = c("a", "b"))), "2 value")
})

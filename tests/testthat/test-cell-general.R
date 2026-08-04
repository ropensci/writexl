# Tests for xl_cell_general() constructor, S3 methods, and write_xlsx integration

# ── Construction: value types ─────────────────────────────────────────────────

test_that("xl_cell_general: numeric value", {
  x <- xl_cell_general(value = 42.5)
  expect_s3_class(x, "xl_cell_general")
  expect_s3_class(x, "xl_cell")
  expect_equal(length(x), 1L)
  expect_equal(x[[1L]][["value"]], 42.5)
  expect_true(is.na(x[[1L]][["formula"]]))
  expect_identical(x[[1L]][["hyperlink"]], NA)
})

test_that("xl_cell_general: integer value", {
  x <- xl_cell_general(value = 7L)
  expect_equal(x[[1L]][["value"]], 7L)
})

test_that("xl_cell_general: logical value", {
  x <- xl_cell_general(value = TRUE)
  expect_equal(x[[1L]][["value"]], TRUE)
})

test_that("xl_cell_general: character value", {
  x <- xl_cell_general(value = "hello")
  expect_equal(x[[1L]][["value"]], "hello")
})

test_that("xl_cell_general: Date value", {
  d <- as.Date("2024-01-15")
  x <- xl_cell_general(value = d)
  expect_equal(x[[1L]][["value"]], d)
})

test_that("xl_cell_general: POSIXct value", {
  dt <- as.POSIXct("2024-01-15 12:00:00", tz = "UTC")
  x <- xl_cell_general(value = dt)
  expect_equal(x[[1L]][["value"]], dt)
})

test_that("xl_cell_general: mixed-type list value", {
  x <- xl_cell_general(value = list(1.5, "text", TRUE))
  expect_equal(length(x), 3L)
  expect_equal(x[[1L]][["value"]], 1.5)
  expect_equal(x[[2L]][["value"]], "text")
  expect_equal(x[[3L]][["value"]], TRUE)
})

test_that("xl_cell_general: NA value produces blank cell", {
  x <- xl_cell_general(value = NA)
  expect_equal(length(x), 1L)
  expect_true(is.na(x[[1L]][["value"]]))
})

test_that("xl_cell_general: no arguments errors with informative message", {
  expect_error(xl_cell_general(), "At least one")
})

test_that("xl_cell_general: value = NA produces explicit empty cell", {
  x <- xl_cell_general(value = NA)
  expect_equal(length(x), 1L)
  expect_true(is.na(x[[1L]][["value"]]))
})

# ── Construction: formula ──────────────────────────────────────────────────────

test_that("xl_cell_general: formula only", {
  x <- xl_cell_general(formula = "=SUM(A1:A10)")
  expect_equal(x[[1L]][["formula"]], "=SUM(A1:A10)")
  expect_identical(x[[1L]][["hyperlink"]], NA)
})

test_that("xl_cell_general: formula with numeric pre-calculated value", {
  x <- xl_cell_general(value = 55.0, formula = "=SUM(A1:A10)")
  expect_equal(x[[1L]][["formula"]], "=SUM(A1:A10)")
  expect_equal(x[[1L]][["value"]], 55.0)
})

test_that("xl_cell_general: formula with character pre-calculated value", {
  x <- xl_cell_general(value = "result", formula = "=TEXT(A1,\"0\")")
  expect_equal(x[[1L]][["formula"]], "=TEXT(A1,\"0\")")
  expect_equal(x[[1L]][["value"]], "result")
})

test_that("xl_cell_general: formula with NA value -> plain formula", {
  x <- xl_cell_general(value = NA, formula = "=A1+1")
  expect_equal(x[[1L]][["formula"]], "=A1+1")
  expect_true(is.na(x[[1L]][["value"]]))
})

test_that("xl_cell_general: formula vector of length > 1", {
  x <- xl_cell_general(formula = c("=A1", "=A2", "=A3"))
  expect_equal(length(x), 3L)
  expect_equal(x[[2L]][["formula"]], "=A2")
})

test_that("xl_cell_general: NA formula -> no formula written", {
  x <- xl_cell_general(formula = NA_character_)
  expect_true(is.na(x[[1L]][["formula"]]))
})

test_that("xl_cell_general: formula not starting with '=' errors", {
  expect_error(xl_cell_general(formula = "SUM(A1:A10)"),
               "must start with '='")
})

# ── Construction: hyperlink ────────────────────────────────────────────────────

test_that("xl_cell_general: hyperlink as character URL", {
  x <- xl_cell_general(hyperlink = "https://example.com")
  expect_equal(x[[1L]][["hyperlink"]], "https://example.com")
})

test_that("xl_cell_general: hyperlink as named list with url only", {
  x <- xl_cell_general(hyperlink = list(url = "https://example.com"))
  expect_equal(x[[1L]][["hyperlink"]][["url"]], "https://example.com")
})

test_that("xl_cell_general: hyperlink as named list with url and tooltip", {
  x <- xl_cell_general(hyperlink = list(url = "https://example.com",
                                         tooltip = "Go to example.com"))
  h <- x[[1L]][["hyperlink"]]
  expect_equal(h[["url"]],     "https://example.com")
  expect_equal(h[["tooltip"]], "Go to example.com")
})

test_that("xl_cell_general: value provides display text for hyperlink cell", {
  x <- xl_cell_general(value    = "Visit",
                        hyperlink = list(url     = "https://example.com",
                                          tooltip = "Go to example.com"))
  expect_equal(x[[1L]][["value"]], "Visit")
  expect_equal(x[[1L]][["hyperlink"]][["url"]], "https://example.com")
})

test_that("xl_cell_general: NA hyperlink -> no hyperlink written", {
  x <- xl_cell_general(hyperlink = NA)
  expect_identical(x[[1L]][["hyperlink"]], NA)
})

test_that("xl_cell_general: hyperlink vector of length > 1", {
  x <- xl_cell_general(hyperlink = c("https://a.com", "https://b.com"))
  expect_equal(length(x), 2L)
  expect_equal(x[[2L]][["hyperlink"]], "https://b.com")
})

test_that("xl_cell_general: invalid hyperlink format errors", {
  expect_error(
    xl_cell_general(hyperlink = list(list(noturl = "x"))),
    "hyperlink"
  )
})

test_that("xl_cell_general: hyperlink list missing url errors", {
  expect_error(
    xl_cell_general(hyperlink = list(list(tooltip = "hover"))),
    "hyperlink"
  )
})

# ── Vectorization ─────────────────────────────────────────────────────────────

test_that("xl_cell_general: length determined from longest input", {
  x <- xl_cell_general(value = 1:3, formula = "=A1")
  expect_equal(length(x), 3L)
  # formula recycled to length 3
  expect_equal(x[[2L]][["formula"]], "=A1")
  expect_equal(x[[3L]][["formula"]], "=A1")
})

test_that("xl_cell_general: all NULL inputs errors", {
  expect_error(xl_cell_general(), "At least one")
})

test_that("xl_cell_general: length() returns number of cells", {
  x <- xl_cell_general(value = 1:5)
  expect_equal(length(x), 5L)
})

# ── S3 methods ────────────────────────────────────────────────────────────────

test_that("[.xl_cell_general preserves class and picks correct cells", {
  x <- xl_cell_general(value = 1:4)
  sub <- x[2:3]
  expect_s3_class(sub, "xl_cell_general")
  expect_equal(length(sub), 2L)
  expect_equal(sub[[1L]][["value"]], 2L)
  expect_equal(sub[[2L]][["value"]], 3L)
})

test_that("c.xl_cell_general concatenates and preserves class", {
  a <- xl_cell_general(value = 1)
  b <- xl_cell_general(value = 2)
  ab <- c(a, b)
  expect_s3_class(ab, "xl_cell_general")
  expect_equal(length(ab), 2L)
  expect_equal(ab[[1L]][["value"]], 1)
  expect_equal(ab[[2L]][["value"]], 2)
})

test_that("rep.xl_cell_general repeats correctly", {
  x <- xl_cell_general(value = 42)
  r <- rep(x, times = 3L)
  expect_s3_class(r, "xl_cell_general")
  expect_equal(length(r), 3L)
  expect_equal(r[[3L]][["value"]], 42)
})

test_that("as.data.frame.xl_cell_general: data.frame() uses the argument name", {
  x <- xl_cell_general(value = 1:3)
  # names(as.data.frame(x)) is NULL so data.frame() keeps the argument name
  expect_null(names(as.data.frame(x)))
  df <- data.frame(a = 1:3, my_col = x)
  expect_equal(names(df), c("a", "my_col"))
  # not a list, and that is the point: `[<-.data.frame` reads a list value as a
  # list of columns, so a list-backed cell vector could not be assigned with
  # `df[, j] <-` at all
  expect_s3_class(df$my_col, "xl_cell_general")
  expect_false(is.list(df$my_col))
})

test_that("rep.xl_cell_general with length.out", {
  x <- xl_cell_general(value = c(1, 2))
  r <- rep(x, length.out = 5L)
  expect_equal(length(r), 5L)
})


test_that("print.xl_cell_general runs without error for length-1 cell", {
  x <- xl_cell_general(value = 99)
  expect_output(print(x), "xl_cell_general")
})

test_that("print.xl_cell_general shows formula and hyperlink", {
  x <- xl_cell_general(formula = "=A1+1")
  expect_output(print(x), "formula")

  y <- xl_cell_general(hyperlink = "https://x.com")
  expect_output(print(y), "hyperlink")
})

test_that("print.xl_cell_general truncates at max and shows count", {
  x <- xl_cell_general(value = seq_len(20))
  expect_output(print(x, max = 5L), "more")
})

test_that("print.xl_cell_general shows <empty> for explicit NA cell", {
  x <- xl_cell_general(value = NA)
  expect_output(print(x), "empty")
})

# ── Data frame integration ────────────────────────────────────────────────────

test_that("length-1 xl_formula recycles in data frame", {
  df <- data.frame(x = 1:3)
  df$f <- xl_formula("=A1*2")
  expect_true(file.exists(write_xlsx(df)))
})

test_that("length-1 xl_hyperlink recycles in data frame", {
  df <- data.frame(x = 1:3)
  df$h <- xl_hyperlink("https://example.com")
  expect_true(file.exists(write_xlsx(df)))
})

test_that("length-1 xl_cell_general(value) recycles in data frame", {
  df <- data.frame(x = 1:3)
  df$v <- xl_cell_general(value = 99L)
  expect_true(file.exists(write_xlsx(df)))
})

test_that("full-length xl_cell_general column writes without error", {
  df <- data.frame(x = 1:3)
  df$c <- xl_cell_general(value = c(10, 20, 30))
  expect_true(file.exists(write_xlsx(df)))
})

test_that("mixed-type xl_cell_general column (numeric + string + formula) writes", {
  mixed <- c(
    xl_cell_general(value = 1.5),
    xl_cell_general(value = "note"),
    xl_cell_general(formula = "=A1+A2")
  )
  df <- data.frame(x = 1:3)
  df$m <- mixed
  expect_true(file.exists(write_xlsx(df)))
})

test_that("formula with numeric pre-calc value writes successfully", {
  df <- data.frame(x = 1:2)
  df$f <- c(
    xl_cell_general(value = 55.0, formula = "=SUM(A1:A2)"),
    xl_cell_general(value = 10.0, formula = "=A1*2")
  )
  expect_true(file.exists(write_xlsx(df)))
})

test_that("formula with character pre-calc value writes successfully", {
  df <- data.frame(x = 1:2)
  df$f <- c(
    xl_cell_general(value = "total", formula = "=TEXT(A1,\"0\")"),
    xl_cell_general(formula = "=A1+1")
  )
  expect_true(file.exists(write_xlsx(df)))
})

test_that("hyperlink with value display text and tooltip writes successfully", {
  df <- data.frame(x = 1:2)
  df$h <- c(
    xl_cell_general(value    = "Visit",
                    hyperlink = list(url     = "https://example.com",
                                      tooltip = "tooltip text")),
    xl_cell_general(hyperlink = "https://other.com")
  )
  expect_true(file.exists(write_xlsx(df)))
})

test_that("NA hyperlink in xl_cell_general column writes blank cell", {
  df <- data.frame(x = 1:2)
  df$h <- c(
    xl_cell_general(hyperlink = "https://example.com"),
    xl_cell_general(hyperlink = NA)
  )
  expect_true(file.exists(write_xlsx(df)))
})

# ── Backward compatibility ─────────────────────────────────────────────────────

test_that("xl_formula() returns xl_cell_general with correct class", {
  x <- xl_formula("=A1+1")
  expect_s3_class(x, "xl_cell_general")
  expect_s3_class(x, "xl_cell")
})

test_that("xl_formula() validates formula starts with '='", {
  expect_error(xl_formula("A1+1"), "start with")
})

test_that("xl_formula() accepts NA values", {
  x <- xl_formula(c("=A1", NA))
  expect_equal(length(x), 2L)
  expect_true(is.na(x[[2L]][["formula"]]))
})

test_that("xl_formula() writes valid xlsx", {
  df <- data.frame(x = 1:3)
  df$f <- xl_formula("=A1*2")
  expect_true(file.exists(write_xlsx(df)))
})

test_that("xl_hyperlink() returns xl_cell_general with correct class", {
  x <- xl_hyperlink("https://example.com")
  expect_s3_class(x, "xl_cell_general")
})

test_that("xl_hyperlink() with no name stores HYPERLINK formula", {
  x <- xl_hyperlink(c("https://a.com", "https://b.com"))
  expect_equal(length(x), 2L)
  expect_equal(x[[1L]][["formula"]], '=HYPERLINK("https://a.com")')
  expect_equal(x[[2L]][["formula"]], '=HYPERLINK("https://b.com")')
})

test_that("xl_hyperlink() with name stores HYPERLINK formula with display text", {
  x <- xl_hyperlink(c("https://a.com", "https://b.com"), c("Site A", "Site B"))
  expect_equal(length(x), 2L)
  expect_equal(x[[1L]][["formula"]], '=HYPERLINK("https://a.com","Site A")')
  expect_equal(x[[2L]][["formula"]], '=HYPERLINK("https://b.com","Site B")')
})

test_that("xl_hyperlink() with NA url produces NA formula (blank cell)", {
  x <- xl_hyperlink(c("https://a.com", NA))
  expect_equal(length(x), 2L)
  expect_true(is.na(x[[2L]][["formula"]]))
})

test_that("xl_hyperlink() escapes internal double quotes in url and value", {
  x <- xl_hyperlink('http://example.com/q?a="1"', value = 'Say "hello"')
  expect_equal(x[[1L]][["formula"]],
               '=HYPERLINK("http://example.com/q?a=""1""","Say ""hello""")')
})

test_that("xl_hyperlink() writes valid xlsx", {
  df <- data.frame(name = c("UCLA", "Berkeley"))
  df$website <- xl_hyperlink(c("http://www.ucla.edu", "http://www.berkeley.edu"),
                               "homepage")
  expect_true(file.exists(write_xlsx(df)))
})

test_that("xl_hyperlink() with NA url writes valid xlsx", {
  df <- data.frame(x = c("a", "b"))
  df$link <- xl_hyperlink(c("http://a.com", NA))
  expect_true(file.exists(write_xlsx(df)))
})

# ── xl_hyperlink_cell() ───────────────────────────────────────────────────────

test_that("xl_hyperlink_cell() returns xl_cell_general with correct class", {
  x <- xl_hyperlink_cell("https://example.com")
  expect_s3_class(x, "xl_cell_general")
})

test_that("xl_hyperlink_cell() with no value stores URL as hyperlink field", {
  x <- xl_hyperlink_cell(c("https://a.com", "https://b.com"))
  expect_equal(length(x), 2L)
  expect_equal(x[[1L]][["hyperlink"]], "https://a.com")
  expect_equal(x[[2L]][["hyperlink"]], "https://b.com")
})

test_that("xl_hyperlink_cell() with value stores display text and hyperlink", {
  x <- xl_hyperlink_cell(c("https://a.com", "https://b.com"), value = c("A", "B"))
  expect_equal(x[[1L]][["value"]],     "A")
  expect_equal(x[[1L]][["hyperlink"]], "https://a.com")
  expect_equal(x[[2L]][["value"]],     "B")
  expect_equal(x[[2L]][["hyperlink"]], "https://b.com")
})

test_that("xl_hyperlink_cell() recycles value to url length", {
  x <- xl_hyperlink_cell(c("https://a.com", "https://b.com"), value = "Link")
  expect_equal(x[[1L]][["value"]], "Link")
  expect_equal(x[[2L]][["value"]], "Link")
})

test_that("xl_hyperlink_cell() with NA url produces blank cell", {
  x <- xl_hyperlink_cell(c("https://a.com", NA))
  expect_equal(length(x), 2L)
  expect_identical(x[[2L]][["hyperlink"]], NA)
})

test_that("xl_hyperlink_cell() with NA url and value: NA url clears value", {
  x <- xl_hyperlink_cell(c("https://a.com", NA), value = "Link")
  expect_equal(x[[1L]][["value"]], "Link")
  expect_true(is.na(x[[2L]][["value"]]))
})

test_that("xl_hyperlink_cell() writes valid xlsx", {
  df <- data.frame(name = c("UCLA", "Berkeley"))
  df$website <- xl_hyperlink_cell(
    c("http://www.ucla.edu", "http://www.berkeley.edu"),
    value = "homepage"
  )
  expect_true(file.exists(write_xlsx(df)))
})

test_that("xl_hyperlink_cell() with NA url writes valid xlsx", {
  df <- data.frame(x = c("a", "b"))
  df$link <- xl_hyperlink_cell(c("http://a.com", NA))
  expect_true(file.exists(write_xlsx(df)))
})

# ── normalize_df regression ───────────────────────────────────────────────────

test_that("normalize_df passes xl_cell_general columns through unchanged", {
  df <- data.frame(x = 1:3)
  df$c <- xl_cell_general(value = c(1, 2, 3))
  norm <- normalize_df(df)
  expect_s3_class(norm$c, "xl_cell_general")
})

test_that("xl_cell_general rejects value types writexl cannot write", {
  # the same predicate guards columns and cell objects, so an unsupported type
  # cannot slip in by being wrapped in a cell
  expect_error(xl_cell_general(value = complex(real = 1, imaginary = 1)),
               "cannot write these value type")
  expect_error(xl_cell_general(value = as.raw(1)), "raw")
  # inside a mixed-type list the offending element is identified by position
  expect_error(xl_cell_general(value = list(1, complex(real = 1, imaginary = 1))),
               "value[[2]]", fixed = TRUE)

  # supported values, including the blank-cell sentinels, are unaffected
  expect_s3_class(xl_cell_general(value = list(1.5, "text", TRUE)), "xl_cell_general")
  expect_s3_class(xl_cell_general(value = list(as.Date("2020-01-01"),
                                               as.POSIXct("2020-01-01", tz = "UTC"))),
                  "xl_cell_general")
  expect_s3_class(xl_cell_general(value = NA), "xl_cell_general")
  expect_s3_class(xl_cell_general(value = list(1, NULL)), "xl_cell_general")
  expect_s3_class(xl_cell_general(formula = "=1+1"), "xl_cell_general")
})

# ── Formatted blank cells ─────────────────────────────────────────────────────
#
# A cell with a format but no writable value must still exist, or the format is
# silently lost.  Excel ignores blank cells that carry no format, so the
# unformatted case must continue to write no cell at all.

# The style index of one cell, or NA_character_ if the cell is absent.
cell_style <- function(path, ref) {
  w <- xlsx_part(path, "xl/worksheets/sheet1.xml", raw = TRUE)
  m <- regmatches(w, regexec(sprintf('<c r="%s"([^>]*)>?', ref), w))[[1]]
  if (!length(m)) return(NA_character_)
  s <- regmatches(m[2L], regexec('s="([0-9]+)"', m[2L]))[[1]]
  if (length(s) == 2L) s[2L] else ""
}

test_that("an NA value with a format writes a formatted blank cell", {
  df <- data.frame(x = 1:2)
  df$b <- xl_cell_general(value = c(NA, NA),
                          format = xl_fill(background = "yellow"))
  p <- write_tmp(df)

  # the cell exists and carries a real (non-default) style
  expect_false(is.na(cell_style(p, "B2")))
  expect_true(nzchar(cell_style(p, "B2")))
  expect_true(as.integer(cell_style(p, "B2")) > 0L)

  # ... and it is genuinely blank: no value and no inline string
  w <- xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE)
  blank <- regmatches(w, regexec('<c r="B2"[^>]*/>', w))[[1]]
  expect_length(blank, 1L)

  # the fill really reached styles.xml
  expect_match(xlsx_part(p, "xl/styles.xml", raw = TRUE), "FFFFFF00",
               ignore.case = TRUE)
})

test_that("an NA value with no format writes no cell at all", {
  df <- data.frame(x = 1:2)
  df$b <- xl_cell_general(value = c(NA, NA))
  expect_true(is.na(cell_style(write_tmp(df), "B2")))
})

test_that("every NA flavour with a format writes a formatted blank", {
  fmt <- xl_fill(background = "yellow")
  for (v in list(NA, NA_character_, NA_real_, NA_integer_, NaN,
                 as.Date(NA), as.POSIXct(NA_real_, origin = "1970-01-01",
                                         tz = "UTC"))) {
    df <- data.frame(x = 1L)
    df$b <- xl_cell_general(value = list(v), format = fmt)
    lbl <- paste(class(v), collapse = "/")
    expect_true(as.integer(cell_style(write_tmp(df), "B2")) > 0L, label = lbl)
  }
})

test_that("a formatted blank survives alongside a comment on the same cell", {
  df <- data.frame(x = 1L)
  df$b <- xl_cell_general(value = NA, comment = "note",
                          format = xl_fill(background = "yellow"))
  p <- write_tmp(df)
  expect_true(as.integer(cell_style(p, "B2")) > 0L)
  expect_match(xlsx_part(p, "xl/comments1.xml", raw = TRUE), "note")
})

test_that("NA in a plain column is left alone even under a column format", {
  # deliberate: Excel already applies the column format to empty cells, so
  # plain columns keep writing no cell for an NA
  sheet <- xl_sheet(data.frame(x = 1L, b = NA_character_),
                    cols = xl_col_spec("b", format = xl_fill(background = "yellow")))
  expect_true(is.na(cell_style(write_tmp(list(S = sheet)), "B2")))
})

# ── Array and dynamic formulas ────────────────────────────────────────────────

# The <sheetData> section of a written sheet, as a single string.
sheet_data <- function(path) {
  w <- xlsx_part(path, "xl/worksheets/sheet1.xml", raw = TRUE)
  substring(w, regexpr("<sheetData>", w, fixed = TRUE))
}

# A one-formula sheet: column A holds data, column B the formula cell.
array_sheet <- function(..., .cm = NA) {
  df <- data.frame(x = 1:2)
  df$f <- xl_cell_general(formula = c("=SUM(A1:A2)", NA), ...)
  write_tmp(df, constant_memory = .cm)
}

test_that("a single-cell array formula is stored as an array over its own cell", {
  s <- sheet_data(array_sheet(array = TRUE))
  expect_match(s, '<f t="array" ref="B2">SUM(A1:A2)</f>', fixed = TRUE)
})

test_that("a single-cell dynamic formula is marked as one and adds metadata", {
  p <- write_tmp(local({
    df <- data.frame(x = 1:2)
    df$f <- xl_cell_general(formula = c("=UNIQUE(A1:A2)", NA), dynamic = TRUE)
    df
  }))
  # cm="1" on the cell is what marks a dynamic array for Excel ...
  expect_match(sheet_data(p), '<c r="B2" cm="1">', fixed = TRUE)
  # ... and it requires the metadata part, which a plain array formula omits
  expect_false(is.null(xlsx_part(p, "xl/metadata.xml")))
  expect_true(is.null(xlsx_part(array_sheet(array = TRUE), "xl/metadata.xml")))
})

test_that("a pre-calculated result is stored for double and integer alike", {
  # the integer case used to fall through to the no-result writer, silently
  # dropping the value and leaving Excel's placeholder zero behind
  for (v in list(3, 3L)) {
    s <- sheet_data(array_sheet(array = TRUE, value = list(v, NA)))
    expect_match(s, '<f t="array" ref="B2">SUM(A1:A2)</f><v>3</v>',
                 fixed = TRUE, label = typeof(v))
  }
  # ... and the same fix applies to a plain (non-array) formula
  df <- data.frame(x = 1L)
  df$f <- xl_cell_general(formula = "=A1", value = 7L)
  expect_match(sheet_data(write_tmp(df)), "<v>7</v>", fixed = TRUE)
})

test_that("a declared multi-cell range is written and padded", {
  df <- data.frame(x = 1:2)
  df$f <- xl_cell_general(formula = c("=TRANSPOSE(A1:A2)", NA), array = TRUE,
                          array_range = list("B2:D2", NA))
  p <- write_tmp(df)
  s <- sheet_data(p)
  expect_match(s, 'ref="B2:D2"', fixed = TRUE)
  # the cells beyond the anchor exist -- libxlsxwriter pads them
  expect_match(s, '<c r="C2">', fixed = TRUE)
  expect_match(s, '<c r="D2">', fixed = TRUE)
})

test_that("a multi-cell range turns row streaming off, and says why", {
  df <- data.frame(x = 1:2)
  df$f <- xl_cell_general(formula = c("=TRANSPOSE(A1:A2)", NA), array = TRUE,
                          array_range = list("B2:D2", NA))
  d <- .resolve_sheet_formats(df, .new_format_registry(), xl_properties(), 1L)
  cm <- .resolve_constant_memory(list(d), xl_properties())
  expect_equal(cm$on, 0L)
  expect_length(cm$reasons, 1L)
  expect_match(cm$reasons, "multi-cell array formula range")

  # observable end to end: with streaming off, strings move to the shared table
  expect_false(is.null(xlsx_part(write_tmp(df), "xl/sharedStrings.xml")))
  # a single-cell array formula leaves streaming on
  # asked for explicitly: a frame this small would not stream on size alone
  expect_true(is.null(xlsx_part(array_sheet(array = TRUE,
                                            .cm = TRUE),
                                "xl/sharedStrings.xml")))
})

test_that("array flags are inert on a cell with no formula", {
  # recycling one TRUE across a column of mixed formula and value cells is
  # normal, so a flag on a value-only cell must simply do nothing
  df <- data.frame(x = 1:2)
  df$f <- xl_cell_general(formula = c("=SUM(A1:A2)", NA),
                          value = list(NA, 5), array = TRUE)
  s <- sheet_data(write_tmp(df))
  expect_match(s, 'ref="B2"', fixed = TRUE)
  expect_match(s, '<c r="B3"><v>5</v></c>', fixed = TRUE)
})

test_that("array arguments are validated", {
  # a flag that can never apply to anything
  expect_error(xl_cell_general(value = 1, array = TRUE), "no `formula`")
  expect_error(xl_cell_general(value = 1, dynamic = TRUE), "no `formula`")
  # a range with nothing to apply to
  expect_error(xl_cell_general(formula = "=A1", array_range = "A1:B2"),
               "applies only to an `array` or `dynamic`")
  # Excel stores no cached string result for an array formula
  expect_error(xl_cell_general(formula = "=A1", value = "txt", array = TRUE),
               "character `value`")
  expect_error(xl_cell_general(formula = "=A1", value = "txt", dynamic = TRUE),
               "character `value`")
  # a character value is still fine for a plain formula
  expect_s3_class(xl_cell_general(formula = "=A1", value = "txt"),
                  "xl_cell_general")
  # the flags are decisions, so NA is not a valid setting
  expect_error(xl_cell_general(formula = "=A1", array = NA), "TRUE or FALSE")
  expect_error(xl_cell_general(formula = "=A1", dynamic = NA), "TRUE or FALSE")
  # an unusable range spelling
  expect_error(xl_cell_general(formula = "=A1", array = TRUE,
                               array_range = list(1L)),
               "must be NA, an Excel range string")
})

test_that("a declared range must start at the cell holding the formula", {
  df <- data.frame(x = 1:2)
  df$f <- xl_cell_general(formula = c("=A1", NA), array = TRUE,
                          array_range = list("C5:D6", NA))
  expect_error(write_xlsx(df), "must start at the cell that holds the formula")
})

test_that("a declared range may not overlap cells the sheet writes itself", {
  df <- data.frame(x = 1:3)
  # B2:B4 is column B for all three data rows -- cells this sheet writes
  df$f <- xl_cell_general(formula = c("=A1", NA, NA), array = TRUE,
                          array_range = list("B2:B4", NA, NA))
  expect_error(write_xlsx(df), "that the sheet writes itself")
  expect_error(write_xlsx(df), "covers 3 cell", fixed = TRUE)
  expect_error(write_xlsx(df), "spills automatically")
})

test_that("array_range accepts a data-frame-relative spec", {
  df <- data.frame(x = 1:2)
  # cols 2:3 does not exist in a 2-column frame, so name the range in A1 terms
  # and check the list spelling on a frame wide enough to hold it
  df$f <- xl_cell_general(formula = c("=A1", NA), array = TRUE)
  expect_silent(write_xlsx(df))

  wide <- data.frame(a = 1:2, b = 1:2, cc = 1:2)
  wide$f <- xl_cell_general(formula = c("=A1", NA), array = TRUE,
                            array_range = list(list(rows = 1, cols = 4), NA))
  # a single-cell spec resolves to the anchor and needs no memory-mode change
  expect_silent(write_xlsx(wide))
})

# ── xl_hyperlink(): name deprecated in favour of value ────────────────────────

test_that("xl_hyperlink(value =) is the supported spelling", {
  x <- xl_hyperlink("https://a.com", value = "Site A")
  expect_match(x[[1L]][["formula"]], 'HYPERLINK("https://a.com","Site A")',
               fixed = TRUE)
})

test_that("xl_hyperlink(name =) still works but warns", {
  expect_warning(x <- xl_hyperlink("https://a.com", name = "Site A"),
                 "deprecated")
  expect_warning(xl_hyperlink("https://a.com", name = "Site A"), "use `value`")
  # and it produces exactly what value= produces
  expect_equal(unclass(x),
               unclass(xl_hyperlink("https://a.com", value = "Site A")))
})

test_that("giving both value and name is an error", {
  expect_error(xl_hyperlink("https://a.com", value = "A", name = "B"),
               "not both")
})

test_that("positional display text is unaffected by the rename", {
  # `value` took the position `name` used to occupy, so old positional code
  # keeps working and must not warn
  expect_silent(x <- xl_hyperlink("https://a.com", "Site A"))
  expect_match(x[[1L]][["formula"]], 'HYPERLINK("https://a.com","Site A")',
               fixed = TRUE)
  # ... and a positional format argument still lands on format
  y <- xl_hyperlink("https://a.com", "Site A", xl_font(bold = TRUE))
  expect_true(is_xl_format(y[[1L]][["format"]]))
})

# ── The guards nothing else reached ──────────────────────────────────────────

test_that("a single array_range spec is one range, not a list of them", {
  # list(rows = , cols = ) has to be wrapped, or its two elements would be
  # taken as two cells' worth of ranges
  cells <- xl_cell_general(formula = "=A1:B2", array = TRUE,
                           array_range = list(rows = 1:2, cols = 1:2))
  expect_length(cells, 1L)
  expect_equal(cells[[1L]]$array_range, list(rows = 1:2, cols = 1:2))
})

test_that("a NULL hyperlink in a list is a cell with no hyperlink", {
  cells <- xl_cell_general(value = c("a", "b"),
                           hyperlink = list(NULL, "https://example.com"))
  expect_true(is.na(cells[[1L]]$hyperlink) || is.null(cells[[1L]]$hyperlink))
  expect_equal(cells[[2L]]$hyperlink, "https://example.com")
})

test_that("an unset array flag is FALSE for every cell", {
  expect_equal(.cell_flag_vec(NULL, "array", 3L), rep(FALSE, 3L))
  expect_error(.cell_flag_vec(NA, "array", 1L), "must be TRUE or FALSE")
})

test_that("an array range reaching past the sheet's data is accepted", {
  # the overlap check only has something to say where the two rectangles meet;
  # the range starts at its own cell (B2) and spills into empty columns
  df <- data.frame(a = 1:2)
  df$f <- xl_cell_general(formula = c("=TRANSPOSE(A2:A3)", NA), array = TRUE,
                          array_range = list("B2:E2", NA))
  expect_silent(write_tmp(df))
})

test_that("a frame with no rows has no data rectangle to overlap", {
  empty <- data.frame(a = numeric(0))
  empty$f <- xl_cell_general(formula = character(0), array = TRUE)
  expect_silent(write_tmp(empty))
})

test_that("a dynamic formula carries its cached result, spilling or not", {
  read_back <- function(cells) {
    df <- data.frame(a = c(1, 2, 3))
    df$f <- cells
    p <- write_tmp(df)
    list(sheet = xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE),
         got = as.data.frame(readxl::read_xlsx(p)))
  }
  # single cell, with the value Excel would compute
  r <- read_back(xl_cell_general(value = c(6, NA, NA),
                                 formula = c("=SUM(A2:A4)", NA, NA),
                                 dynamic = TRUE))
  expect_equal(r$got$f[1L], 6)

  # a declared range spills into columns the sheet does not write, again with
  # a cached value for the anchor cell
  r <- read_back(xl_cell_general(value = c(1, NA, NA),
                                 formula = c("=TRANSPOSE(A2:A4)", NA, NA),
                                 dynamic = TRUE,
                                 array_range = list("B2:D2", NA, NA)))
  expect_match(r$sheet, 'cm="1"', fixed = TRUE)
  expect_equal(r$got$f[1L], 1)

  # and without a cached value, on both paths
  r <- read_back(xl_cell_general(formula = c("=SUM(A2:A4)", NA, NA),
                                 dynamic = TRUE))
  expect_match(r$sheet, "SUM(A2:A4)", fixed = TRUE)
  r <- read_back(xl_cell_general(formula = c("=TRANSPOSE(A2:A4)", NA, NA),
                                 dynamic = TRUE,
                                 array_range = list("B2:D2", NA, NA)))
  expect_match(r$sheet, "TRANSPOSE", fixed = TRUE)
})

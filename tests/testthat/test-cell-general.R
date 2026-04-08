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
  expect_true(is.list(df$my_col))
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

test_that("xl_hyperlink() with name stores name as value (display text)", {
  x <- xl_hyperlink(c("https://a.com", "https://b.com"), c("Site A", "Site B"))
  expect_equal(length(x), 2L)
  expect_equal(x[[1L]][["value"]],     "Site A")
  expect_equal(x[[1L]][["hyperlink"]], "https://a.com")
  expect_equal(x[[2L]][["value"]],     "Site B")
  expect_equal(x[[2L]][["hyperlink"]], "https://b.com")
})

test_that("xl_hyperlink() with NA url produces NA hyperlink", {
  x <- xl_hyperlink(c("https://a.com", NA))
  expect_equal(length(x), 2L)
  expect_identical(x[[2L]][["hyperlink"]], NA)
})

test_that("xl_hyperlink() writes valid xlsx", {
  df <- data.frame(name = c("UCLA", "Berkeley"))
  df$website <- xl_hyperlink(c("http://www.ucla.edu", "http://www.berkeley.edu"),
                               "homepage")
  expect_true(file.exists(write_xlsx(df)))
})

test_that("xl_hyperlink() with NA name value writes valid xlsx", {
  df <- data.frame(x = c("a", "b"))
  df$link <- xl_hyperlink(c("http://a.com", NA))
  expect_true(file.exists(write_xlsx(df)))
})

# ── normalize_df regression ───────────────────────────────────────────────────

test_that("normalize_df passes xl_cell_general columns through unchanged", {
  df <- data.frame(x = 1:3)
  df$c <- xl_cell_general(value = c(1, 2, 3))
  norm <- writexl:::normalize_df(df)
  expect_s3_class(norm$c, "xl_cell_general")
})

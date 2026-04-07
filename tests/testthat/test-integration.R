# Integration test: single xlsx file exercising every package feature.
#
# When this file is opened in Excel / LibreOffice the "expected" column on
# each sheet describes exactly what should appear in each data column next
# to it, making the file self-documenting for manual spot-checks.

test_that("comprehensive xlsx output exercises all package functionality", {
  skip_if_not_installed("readxl")

  # ── Sheet 1: basic scalar types ───────────────────────────────────────────
  # Three rows: a normal value, a blank (NA), and a second value or edge case.
  basic <- data.frame(
    expected     = c(
      "row 1: normal values in every column",
      "row 2: all columns blank (NA)",
      "row 3: FALSE / 2 / Inf written as text 'Inf' / world / 2024-12-31 / second datetime"
    ),
    logical_col  = c(TRUE,                NA,             FALSE),
    integer_col  = c(1L,                  NA_integer_,    2L),
    double_col   = c(1.5,                 NA_real_,       Inf),
    char_col     = c("hello",             NA_character_,  "world"),
    date_col     = as.Date(c("2024-01-15", NA,            "2024-12-31")),
    datetime_col = as.POSIXct(
      c("2024-06-01 00:00:00", NA, "2024-12-31 23:59:59"), tz = "UTC"
    ),
    stringsAsFactors = FALSE
  )

  # ── Sheet 2: xl_formula ───────────────────────────────────────────────────
  # Formulas are stored in the file; Excel evaluates them on open.
  # NA formula produces a blank cell.
  formulas <- data.frame(
    x        = c(10, 20, 30),
    expected = c(
      "formula =B2*2  ->  Excel shows 20 after recalculation",
      "formula =B3*2  ->  Excel shows 40 after recalculation",
      "blank (NA formula)"
    ),
    result   = xl_formula(c("=B2*2", "=B3*2", NA)),
    stringsAsFactors = FALSE
  )

  # ── Sheet 3: xl_hyperlink ────────────────────────────────────────────────
  # Hyperlinks are clickable; named ones show display text instead of the URL.
  # NA URL produces a blank cell.
  links <- data.frame(
    expected   = c(
      "plain hyperlink  ->  raw URL 'https://www.r-project.org' shown, clickable",
      "named hyperlink  ->  display text 'CRAN' shown, clickable",
      "blank (NA URL, no name)",
      "blank (NA URL, name ignored)"
    ),
    plain_link = xl_hyperlink(
      c("https://www.r-project.org", "https://cran.r-project.org", NA, NA)
    ),
    named_link = xl_hyperlink(
      c("https://www.r-project.org", "https://cran.r-project.org", NA, NA),
      name = c("R Project", "CRAN", "none", "none")
    ),
    stringsAsFactors = FALSE
  )

  # ── Sheet 4: xl_cell_general — one cell per feature ───────────────────────
  # The "expected" column describes exactly what Excel should display.
  general <- data.frame(
    expected = c(
      # value-only cells
      "42  (integer value)",
      "3.14  (double value)",
      "hello  (character value)",
      "TRUE  (logical value)",
      "2024-03-01 as a number (Date value; no column-level date format here)",
      "blank  (value = NA, explicit empty cell)",
      # formula cells
      "formula =A2+1, no cached value  ->  Excel shows 43 after recalculation",
      "formula =SUM(A2:A3) with cached result 45.14  ->  shows 45.14 immediately",
      "formula =TEXT(A2,\"0\") with cached result '42'  ->  shows 42 immediately",
      # hyperlink cells
      "hyperlink https://example.com  ->  raw URL shown, clickable",
      "hyperlink https://example.com  ->  display text 'Visit' shown, clickable",
      "hyperlink https://example.com  ->  display text 'Hover', tooltip on mouseover",
      "hyperlink https://example.com  ->  display text 'Named', tooltip on mouseover"
    ),
    stringsAsFactors = FALSE
  )
  general$result <- c(
    # value-only
    xl_cell_general(value = 42L),
    xl_cell_general(value = 3.14),
    xl_cell_general(value = "hello"),
    xl_cell_general(value = TRUE),
    xl_cell_general(value = as.Date("2024-03-01")),
    xl_cell_general(value = NA),
    # formula
    xl_cell_general(formula = "=A2+1"),
    xl_cell_general(value = 45.14, formula = "=SUM(A2:A3)"),
    xl_cell_general(value = "42",  formula = "=TEXT(A2,\"0\")"),
    # hyperlink
    xl_cell_general(hyperlink = "https://example.com"),
    xl_cell_general(value = "Visit", hyperlink = "https://example.com"),
    xl_cell_general(value = "Hover", hyperlink = list(
      url     = "https://example.com",
      tooltip = "Go to example.com"
    )),
    xl_cell_general(value = "Named", hyperlink = list(
      url     = "https://example.com",
      tooltip = "Tooltip on named hyperlink"
    ))
  )

  # ── Sheet 5: xl_cell_general — mixed-type column built with c() ───────────
  # Different R types in consecutive cells of the same column.
  mixed <- data.frame(
    expected = c(
      "1.5  (double cell)",
      "note  (character cell)",
      "formula =A2+A3  ->  Excel shows 3.5 after recalculation",
      "hyperlink https://www.r-project.org  ->  display text 'R', clickable"
    ),
    stringsAsFactors = FALSE
  )
  mixed$result <- c(
    xl_cell_general(value = 1.5),
    xl_cell_general(value = "note"),
    xl_cell_general(formula = "=A2+A3"),
    xl_cell_general(value = "R", hyperlink = "https://www.r-project.org")
  )

  # ── Write all sheets with bold/centred headers ────────────────────────────
  path <- write_xlsx(
    list(
      basic_types  = basic,
      formulas     = formulas,
      hyperlinks   = links,
      cell_general = general,
      mixed_types  = mixed
    ),
    format_headers = TRUE
  )

  expect_true(file.exists(path))
  expect_gt(file.size(path), 0L)

  # ── col_names = FALSE ─────────────────────────────────────────────────────
  # Data starts in row 1; no header row written.
  path_no_header <- write_xlsx(basic, col_names = FALSE)
  expect_true(file.exists(path_no_header))
  rt_no_header <- readxl::read_xlsx(path_no_header, col_names = FALSE)
  expect_equal(nrow(rt_no_header), nrow(basic))   # same number of data rows

  # ── format_headers = FALSE ────────────────────────────────────────────────
  # Headers written as plain (unformatted) text.
  path_plain <- write_xlsx(basic, format_headers = FALSE)
  expect_true(file.exists(path_plain))

  # ── Round-trip spot-checks on basic_types ─────────────────────────────────
  rt <- readxl::read_xlsx(path, sheet = "basic_types")
  expect_equal(nrow(rt), nrow(basic))

  expect_true(rt$logical_col[1L])              # readxl returns logical TRUE/FALSE
  expect_false(rt$logical_col[3L])
  expect_true(is.na(rt$logical_col[2L]))

  expect_equal(rt$integer_col[1L], 1)          # readxl returns all numbers as double
  expect_true(is.na(rt$integer_col[2L]))

  # double_col: Inf is written as the string "Inf", so readxl coerces the
  # whole column to character — round-trip of numeric values is not expected.
  expect_equal(rt$char_col[1L], "hello")
  expect_true(is.na(rt$char_col[2L]))

  # ── lxw_version ───────────────────────────────────────────────────────────
  # Returns the bundled libxlsxwriter version as a numeric_version.
  ver <- lxw_version()
  expect_s3_class(ver, "numeric_version")
  expect_gte(ver, numeric_version("1.0.0"))
})

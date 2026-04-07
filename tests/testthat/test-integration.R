# Integration test: single xlsx file exercising every package feature.
#
# When this file is opened in Excel / LibreOffice the "expected" column on
# each sheet describes exactly what should appear in each data column next
# to it, making the file self-documenting for manual spot-checks.

test_that("comprehensive xlsx output exercises all package functionality", {
  skip_if_not_installed("readxl")

  # ── Sheet 1: basic_types ─────────────────────────────────────────────────
  # Three rows: a normal value, a blank (NA), and a second value / edge case.
  # Each type gets its own column so readxl can round-trip them correctly.
  #
  # Col A: expected (description)
  # Col B: logical_col    — TRUE / blank / FALSE
  # Col C: integer_col    — 1 / blank / 2
  # Col D: double_col     — 1.5 / blank / "Inf" (written as text)
  # Col E: char_col       — "hello" / blank / "world"
  # Col F: date_col       — 2024-01-15 / blank / 2024-12-31  (yyyy-mm-dd)
  # Col G: datetime_col   — 2024-06-01 00:00:00 / blank / 2024-12-31 23:59:59 UTC
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

  # ── Sheet 2: formulas ────────────────────────────────────────────────────
  # Col A: x        — numeric values used by the formulas in col C
  # Col B: expected — description of what col C should show
  # Col C: result   — formula cells; NA formula produces a blank cell
  #
  # Formulas reference col A (the x values).
  formulas <- data.frame(
    x        = c(10, 20, 30),
    expected = c(
      "formula =A2*2  ->  Excel shows 20 after recalculation",
      "formula =A3*2  ->  Excel shows 40 after recalculation",
      "blank (NA formula)"
    ),
    stringsAsFactors = FALSE
  )
  formulas$result <- xl_formula(c("=A2*2", "=A3*2", NA))

  # ── Sheet 3: hyperlinks ───────────────────────────────────────────────────
  # Col A: expected   — description
  # Col B: plain_link — raw URL shown in cell, clickable; NA = blank
  # Col C: named_link — display text shown instead of URL; NA = blank
  links <- data.frame(
    expected = c(
      "plain hyperlink  ->  raw URL 'https://www.r-project.org' shown, clickable",
      "named hyperlink  ->  display text 'CRAN' shown, clickable",
      "blank (NA URL, no name)",
      "blank (NA URL, name ignored)"
    ),
    stringsAsFactors = FALSE
  )
  links$plain_link <- xl_hyperlink(
    c("https://www.r-project.org", "https://cran.r-project.org", NA, NA)
  )
  links$named_link <- xl_hyperlink(
    c("https://www.r-project.org", "https://cran.r-project.org", NA, NA),
    name = c("R Project", "CRAN", "none", "none")
  )

  # ── Sheet 4: cell_general ─────────────────────────────────────────────────
  # Col A: expected — description of what col B should show
  # Col B: result   — one xl_cell_general per row demonstrating each feature
  #
  # Formulas in col B reference other cells in col B.
  # B2 = 42  (integer value, used as formula input)
  # B3 = 3.14 (double value, used as formula input)
  # B8 = =B2+1     -> 43
  # B9 = =SUM(B2:B3) with cached result 45.14  -> shows 45.14 immediately
  # B10 = =TEXT(B2,"0") with cached result "42" -> shows 42 immediately
  general <- data.frame(
    expected = c(
      # value-only cells (rows 2–7)
      "42  (integer value)",
      "3.14  (double value)",
      "hello  (character value)",
      "TRUE  (logical value)",
      "2024-03-01 as a serial number (Date value; no column-level date format here)",
      "blank  (value = NA, explicit empty cell)",
      # formula cells (rows 8–10)
      "formula =B2+1, no cached value  ->  Excel shows 43 after recalculation",
      "formula =SUM(B2:B3) with cached result 45.14  ->  shows 45.14 immediately",
      "formula =TEXT(B2,\"0\") with cached result '42'  ->  shows 42 immediately",
      # hyperlink cells (rows 11–14)
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
    xl_cell_general(formula = "=B2+1"),
    xl_cell_general(value = 45.14, formula = "=SUM(B2:B3)"),
    xl_cell_general(value = "42",  formula = "=TEXT(B2,\"0\")"),
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

  # ── Sheet 5: mixed_types ─────────────────────────────────────────────────
  # Col A: expected — description
  # Col B: result   — mixed-type column built with c()
  #
  # B2 = 1.5 (double)
  # B3 = "note" (character — cannot be used in arithmetic)
  # B4 = =B2*2  ->  3 (references the numeric value in B2)
  # B5 = hyperlink
  mixed <- data.frame(
    expected = c(
      "1.5  (double cell)",
      "note  (character cell)",
      "formula =B2*2  ->  Excel shows 3 after recalculation",
      "hyperlink https://www.r-project.org  ->  display text 'R', clickable"
    ),
    stringsAsFactors = FALSE
  )
  mixed$result <- c(
    xl_cell_general(value = 1.5),
    xl_cell_general(value = "note"),
    xl_cell_general(formula = "=B2*2"),
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
  expect_equal(nrow(rt_no_header), nrow(basic))

  # ── format_headers = FALSE ────────────────────────────────────────────────
  # Headers written as plain (unformatted) text.
  path_plain <- write_xlsx(basic, format_headers = FALSE)
  expect_true(file.exists(path_plain))

  # ── Round-trip spot-checks on basic_types ─────────────────────────────────
  rt <- readxl::read_xlsx(path, sheet = "basic_types")
  expect_equal(nrow(rt), nrow(basic))
  expect_equal(ncol(rt), ncol(basic))

  # Verify column names are correct
  expect_equal(names(rt), names(basic))

  # Verify column names on formulas and hyperlinks sheets
  rt_f <- readxl::read_xlsx(path, sheet = "formulas")
  expect_equal(names(rt_f), c("x", "expected", "result"))

  rt_l <- readxl::read_xlsx(path, sheet = "hyperlinks")
  expect_equal(names(rt_l), c("expected", "plain_link", "named_link"))

  expect_true(rt$logical_col[1L])             # readxl returns logical TRUE/FALSE
  expect_false(rt$logical_col[3L])
  expect_true(is.na(rt$logical_col[2L]))

  expect_equal(rt$integer_col[1L], 1)         # readxl returns all numbers as double
  expect_true(is.na(rt$integer_col[2L]))

  # double_col: Inf is written as the string "Inf", causing readxl to read
  # the whole column as character — numeric round-trip is not expected.
  expect_equal(rt$char_col[1L], "hello")
  expect_true(is.na(rt$char_col[2L]))

  # ── lxw_version ───────────────────────────────────────────────────────────
  # Returns the bundled libxlsxwriter version as a numeric_version object.
  ver <- lxw_version()
  expect_s3_class(ver, "numeric_version")
  expect_gte(ver, numeric_version("1.0.0"))
})

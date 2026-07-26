roundtrip <- function(df){
  readxl::read_xlsx(writexl::write_xlsx(df))
}

test_that("Types roundtrip properly",{
  kremlin <- "http://\u043F\u0440\u0435\u0437\u0438\u0434\u0435\u043D\u0442.\u0440\u0444"
  num <- c(NA_real_, pi, 1.2345e80)
  int <- c(NA_integer_, 0L, -100L)
  str <- c(NA_character_, "foo", kremlin) #note empty strings don't work yet
  time <- Sys.time() + 1:3
  bigint <- bit64::as.integer64(.Machine$integer.max) ^ c(0,1,1.5)
  input <- data.frame(num = num, int = int, bigint = bigint, str = str, time = time, stringsAsFactors = FALSE)
  expect_warning(output <- roundtrip(input), "int64")
  output$bigint <- bit64::as.integer64(output$bigint)
  # Excel has no time zones. Every datetime here shares one zone, so writexl
  # writes local wall-clock time rather than the UTC instant (see the time-zone
  # tests below); shift back by the UTC offset to recover the instant.
  output$time <- output$time - as.POSIXlt(time)$gmtoff
  attr(output$time, 'tzone') <- attr(time, 'tzone')
  expect_equal(input, as.data.frame(output))
})

# Extract the numeric <v> values of a written worksheet, keyed by cell
# reference (e.g. c(A2 = 1, B2 = 1)).  Reading the raw serials straight out of
# the XML keeps these assertions independent of readxl, so a writexl bug cannot
# be masked by a compensating reader bug.
sheet_serials <- function(path, part = "xl/worksheets/sheet1.xml") {
  xml <- xlsx_part(path, part, raw = TRUE)
  cells <- regmatches(xml, gregexpr("<c r=\"[A-Z]+[0-9]+\"[^>]*><v>[^<]*</v></c>", xml))[[1]]
  refs <- sub("^<c r=\"([A-Z]+[0-9]+)\".*$", "\\1", cells)
  vals <- as.numeric(sub("^.*<v>([^<]*)</v>.*$", "\\1", cells))
  stats::setNames(vals, refs)
}

# Excel's 1900 date system pretends 1900 was a leap year, so serial 60 is the
# nonexistent 1900-02-29 and every date from 1900-03-01 onwards is shifted by
# one.  These serials are Excel's documented ground truth.
date_boundary_cases <- data.frame(
  date = as.Date(c("1900-01-01", "1900-02-28", "1900-03-01", "1970-01-01", "2020-03-01")),
  serial = c(1, 59, 61, 25569, 43891)
)

test_that("Date and POSIXct columns write Excel's documented serial numbers", {
  d <- date_boundary_cases$date
  expected <- date_boundary_cases$serial
  df <- data.frame(dt = d, pt = as.POSIXct(paste(d, "00:00:00"), tz = "UTC"))
  serials <- sheet_serials(write_tmp(df))

  # Row 1 is the header; data starts at row 2.  Column A = Date, B = POSIXct.
  rows <- seq_along(d) + 1L
  expect_equal(unname(serials[paste0("A", rows)]), expected)
  expect_equal(unname(serials[paste0("B", rows)]), expected)

  # The two writers must agree cell for cell, not merely both look plausible.
  expect_equal(unname(serials[paste0("A", rows)]), unname(serials[paste0("B", rows)]))

  # Serial 60 is Excel's phantom 1900-02-29 and must never be written.
  expect_false(any(serials[paste0(c("A", "B"), rep(rows, each = 2))] == 60))
})

test_that("Integer-backed Date columns write the same serials as double-backed", {
  d <- date_boundary_cases$date
  int_date <- structure(as.integer(unclass(d)), class = "Date")
  expect_identical(typeof(int_date), "integer")

  serials <- sheet_serials(write_tmp(data.frame(dt = int_date)))
  expect_equal(unname(serials[paste0("A", seq_along(d) + 1L)]), date_boundary_cases$serial)
})

test_that("NA is preserved in Date columns of either backing type", {
  d <- c(as.Date("1900-02-28"), NA, as.Date("2020-03-01"))
  int_date <- structure(as.integer(unclass(d)), class = "Date")

  for (col in list(d, int_date)) {
    path <- write_tmp(data.frame(dt = col))
    serials <- sheet_serials(path)
    # The NA cell is written as an empty cell, so no <v> exists for A3.
    expect_equal(unname(serials[c("A2", "A4")]), c(59, 43891))
    expect_false("A3" %in% names(serials))
    expect_equal(as.Date(readxl::read_xlsx(path)[[1]]), d)
  }
})

test_that("Dates around the 1900 leap-day boundary roundtrip through readxl", {
  d <- date_boundary_cases$date
  out <- readxl::read_xlsx(write_tmp(data.frame(dt = d)))
  expect_equal(as.Date(out$dt), d)

  # Every day of the 1900 boundary window, not just the hand-picked ones.
  window <- seq(as.Date("1900-01-01"), as.Date("1900-03-05"), by = "day")
  out2 <- readxl::read_xlsx(write_tmp(data.frame(dt = window)))
  expect_equal(as.Date(out2$dt), window)
  expect_false(anyNA(out2$dt))
})

test_that("Writing formulas", {
  df <- data.frame(
    name = c("UCLA", "Berkeley"),
    founded = c(1919, 1868),
    website = xl_hyperlink(c("http://www.ucla.edu", "http://www.berkeley.edu"), "website")
  )

  # repeats a formula for entire column
  df$age <- xl_formula('=(YEAR(TODAY()) - INDIRECT("B" & ROW()))')

  # currently readxl does not support formulas so inspect manually
  expect_true(file.exists(write_xlsx(df)))
})

test_that("unsupported column types raise an error naming the column", {
  df <- data.frame(a = 1)
  df$bad <- complex(real = 1, imaginary = 1)
  expect_error(write_xlsx(df), "cannot write column")
  expect_error(write_xlsx(df), "'bad' (column 2)", fixed = TRUE)
  expect_error(write_xlsx(df), "complex")

  # a bare list column and a raw vector are equally unrepresentable; the
  # underlying type is reported when the class alone would be unhelpful
  df2 <- data.frame(a = 1); df2$bad <- I(list(1))
  expect_error(write_xlsx(df2), "(list)", fixed = TRUE)
  df3 <- data.frame(a = 1); df3$bad <- as.raw(1)
  expect_error(write_xlsx(df3), "raw")

  # every unsupported column is reported, not just the first
  df4 <- data.frame(a = 1)
  df4$b <- complex(real = 1, imaginary = 1)
  df4$c <- as.raw(2)
  expect_error(write_xlsx(df4), "'b'")
  expect_error(write_xlsx(df4), "'c'")
})

test_that("classed numeric columns such as difftime are still written", {
  # difftime is a double underneath, so it is written from its numeric value
  df <- data.frame(a = 1)
  df$elapsed <- as.difftime(c(1.5), units = "days")
  expect_silent(path <- write_xlsx(df))
  expect_equal(readxl::read_xlsx(path)$elapsed, 1.5)
})

test_that("supported column types are still accepted", {
  df <- data.frame(int = 1:2, dbl = c(1.5, 2.5), chr = c("x", "y"),
                   lgl = c(TRUE, FALSE), date = as.Date("2020-01-01") + 0:1,
                   stringsAsFactors = FALSE)
  df$time <- as.POSIXct(c("2020-01-01 10:00", "2020-01-02 10:00"), tz = "UTC")
  df$cell <- xl_cell_general(value = 1:2)
  expect_true(file.exists(write_xlsx(df)))
  # classes that normalize_df coerces must not trip the check
  expect_true(file.exists(write_xlsx(data.frame(f = factor(c("a", "b"))))))
  expect_true(file.exists(write_xlsx(data.frame(p = as.POSIXlt("2020-01-01")))))
})

# ── POSIXct time zones (Excel stores naive datetimes) ─────────────────────────

# The Excel serial for a naive wall clock, independent of writexl's own maths.
expected_serial <- function(wall_clock) {
  25569 + as.numeric(as.POSIXct(wall_clock, tz = "UTC")) / 86400
}
datetime_format_of <- function(path) {
  xml <- xlsx_part(path, "xl/styles.xml", raw = TRUE)
  if (grepl("HH:mm:ss UTC", xml, fixed = TRUE)) "utc-labelled" else "unlabelled"
}

test_that("a single time zone is dropped, writing local wall-clock time", {
  # the reproducer from the issue: Perth 01:00 must not become 17:00 UTC
  df <- data.frame(dt = as.POSIXct("2025-12-01 01:00:00", tz = "Australia/Perth"))
  path <- write_tmp(df)
  expect_equal(sheet_serials(path)[["A2"]],
               expected_serial("2025-12-01 01:00:00"))
  # and the default format must no longer claim the value is UTC
  expect_equal(datetime_format_of(path), "unlabelled")
})

test_that("dropping a time zone is daylight-saving correct", {
  # both instants read 12:00 locally despite different UTC offsets
  df <- data.frame(dt = as.POSIXct(c("2025-01-15 12:00:00", "2025-07-15 12:00:00"),
                                   tz = "America/New_York"))
  s <- sheet_serials(write_tmp(df))
  expect_equal(s[["A2"]], expected_serial("2025-01-15 12:00:00"))
  expect_equal(s[["A3"]], expected_serial("2025-07-15 12:00:00"))
})

test_that("an all-UTC workbook is unchanged in value", {
  df <- data.frame(dt = as.POSIXct("2020-06-01 12:00:00", tz = "UTC"))
  expect_equal(sheet_serials(write_tmp(df))[["A2"]],
               expected_serial("2020-06-01 12:00:00"))
})

test_that("differing time zones warn and are converted to UTC", {
  df <- data.frame(a = as.POSIXct("2025-12-01 01:00:00", tz = "Australia/Perth"))
  df$b <- as.POSIXct("2025-12-01 01:00:00", tz = "UTC")
  expect_warning(path <- write_tmp(df), "different time zones")
  expect_warning(write_tmp(df), "converted to UTC")
  s <- sheet_serials(path)
  expect_equal(s[["A2"]], expected_serial("2025-11-30 17:00:00"))  # Perth -> UTC
  expect_equal(s[["B2"]], expected_serial("2025-12-01 01:00:00"))
  # here the UTC label is accurate, so it is kept
  expect_equal(datetime_format_of(path), "utc-labelled")
})

test_that("the time zone survey includes POSIXct inside cell objects", {
  df <- data.frame(a = as.POSIXct("2025-12-01 01:00:00", tz = "Australia/Perth"))
  df$b <- xl_cell_general(value = as.POSIXct("2025-12-01 01:00:00", tz = "UTC"))
  expect_warning(write_tmp(df), "different time zones")
  # a cell object sharing the column's zone must NOT warn, and is shifted too
  df2 <- data.frame(a = as.POSIXct("2025-12-01 01:00:00", tz = "Australia/Perth"))
  df2$b <- xl_cell_general(value = as.POSIXct("2025-12-01 03:00:00", tz = "Australia/Perth"))
  expect_no_warning(path <- write_tmp(df2))
  s <- sheet_serials(path)
  expect_equal(s[["A2"]], expected_serial("2025-12-01 01:00:00"))
  expect_equal(s[["B2"]], expected_serial("2025-12-01 03:00:00"))
})

test_that("time zones are surveyed across every sheet, not per sheet", {
  s1 <- data.frame(a = as.POSIXct("2025-12-01 01:00:00", tz = "Australia/Perth"))
  s2 <- data.frame(a = as.POSIXct("2025-12-01 01:00:00", tz = "UTC"))
  expect_warning(write_tmp(list(one = s1, two = s2)), "different time zones")
})

test_that("a supplied datetime_format is never overridden", {
  df <- data.frame(dt = as.POSIXct("2025-12-01 01:00:00", tz = "Australia/Perth"))
  wb <- xl_workbook(df, properties = xl_properties(
    datetime_format = xl_num_format("yyyy/mm/dd hh:mm")))
  expect_match(xlsx_part(write_tmp(wb), "xl/styles.xml", raw = TRUE),
               "yyyy/mm/dd hh:mm", fixed = TRUE)
})

test_that("workbooks without POSIXct are untouched", {
  expect_no_warning(write_tmp(data.frame(x = 1:3, d = as.Date("2020-01-01") + 0:2)))
})

test_that("NA datetimes survive the time zone pass", {
  df <- data.frame(dt = as.POSIXct(c("2025-12-01 01:00:00", NA), tz = "Australia/Perth"))
  path <- write_tmp(df)
  s <- sheet_serials(path)
  expect_equal(s[["A2"]], expected_serial("2025-12-01 01:00:00"))
  expect_false("A3" %in% names(s))     # NA stays an empty cell
})

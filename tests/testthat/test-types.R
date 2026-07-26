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

test_that("Performance is OK", {
  skip_if_not_installed("nycflights13")
  flights <- nycflights13::flights
  tmp <- writexl::write_xlsx(flights)
  out <- readxl::read_xlsx(tmp)
  unlink(tmp)
  # Excel has no time zones. Every datetime here shares one zone, so writexl
  # writes local wall-clock time rather than the UTC instant (see the time-zone
  # tests in test-types.R). Shift back to recover the instant; the offset is
  # per value because this column spans both EST and EDT.
  out$time_hour <- out$time_hour - as.POSIXlt(flights$time_hour)$gmtoff
  attr(out$time_hour, "tzone") <- attr(flights$time_hour, "tzone")
  # max_diffs caps the report: testthat uses max_diffs = Inf on CI, so a
  # mismatch in this 336k-row frame would otherwise print one line per row.
  expect_equal(out, flights, max_diffs = 10)
})

context("Performance")

test_that("Performance is OK", {
  tmp <- writexl::write_xlsx(nycflights13::flights)
  out <- readxl::read_xlsx(tmp)
  unlink(tmp)
  attr(out$time_hour, 'tzone') <- attr(nycflights13::flights$time_hour, 'tzone')
  expect_equal(out, nycflights13::flights)
})

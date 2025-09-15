context("Types")

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

test_that("Hyperlinks automatically escape (#89)", {
  skip("Manually test hyperlink escapes-- try opening the files with Excel")
  df <- data.frame(website = xl_hyperlink("http://www.berkeley.edu", 'Berkeley "homepage'))
  file_hyperlink_test_text <- write_xlsx(df)
  message("Try opening this file in Excel to verify the hyperlink text works with double quotes: ", file_hyperlink_test_text)
  df <- data.frame(website = xl_hyperlink("http://www.berkeley.edu/\"", 'Berkeley ""homepage'))
  file_hyperlink_test_url <- write_xlsx(df)
  message("Try opening this file in Excel to verify the hyperlink url works with double quotes: ", file_hyperlink_test_url)
})

test_that("dubquote escapes double quoting (#89)", {
  expect_equal(dubquote('a"a'), '"a""a"')

  # use `CONCATENATE()` for strings longer than 255 characters (#91)
  char_255 <- paste(rep("A", 255), collapse = "")
  char_256 <- paste(rep("A", 256), collapse = "")
  expect_equal(
    dubquote(char_256),
    sprintf('CONCATENATE("%s","%s")', char_255, "A")
  )
  # CONCATENATE handling accounts for double quotes
  char_300qt <- paste(rep('123456789"', 30), collapse = "")
  char_300qt_escape <- paste(rep('123456789""', 30), collapse = "")
  # Note that the split is slightly off the optimum due to a quote near the
  # border of a split. It will be a bit conservative but still work with
  # hundreds of " in a row.
  expect_equal(
    dubquote(char_300qt),
    sprintf('CONCATENATE("%s","%s")', substr(char_300qt_escape, 0, 253), substr(char_300qt_escape, 254, nchar(char_300qt_escape)))
  )

  # Many double-quotes in a row still works
  char_200qt <- paste(rep('"', 200), collapse = "")
  char_200qt_escape <- paste(rep('"', 400), collapse = "")
  expect_equal(
    dubquote(char_200qt),
    sprintf('CONCATENATE("%s","%s")', substr(char_200qt_escape, 0, 254), substr(char_200qt_escape, 255, nchar(char_200qt_escape)))
  )
})

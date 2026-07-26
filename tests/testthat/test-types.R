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

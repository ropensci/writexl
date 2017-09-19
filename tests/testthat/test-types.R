context("Types")

roundtrip <- function(df){
  readxl::read_xlsx(writexl::write_xlsx(df))
}

test_that("Types roundtrip properly",{
  kremlin <- "http://\u043F\u0440\u0435\u0437\u0438\u0434\u0435\u043D\u0442.\u0440\u0444"
  num <- c(NA_real_, pi, 1.2345e80)
  int <- c(NA_integer_, 0L, -100L)
  str <- c(NA_character_, "foo", kremlin) #note emptry string's dont work yet
  time <- Sys.time() + 1:3
  df <- data.frame(num = num, int = int, str = str, time = time, stringsAsFactors = FALSE)
  expect_equal(df, as.data.frame(roundtrip(df)))
})

test_that("Writing formulas", {
  df <- data.frame(
    name = c("UCLA", "Berkeley"),
    website = xl_formula(c(
      '=HYPERLINK("http://www.ucla.edu", "website")',
      '=HYPERLINK("http://www.berkeley.edu", "website")'
    ))
  )
  # currently readxl does not support formulas so inspect manually
  expect_true(file.exists(write_xlsx(df)))
})

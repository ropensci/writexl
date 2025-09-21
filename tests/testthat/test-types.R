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

test_that("xl_hyperlink_cell", {
  df <- data.frame(
    name = c("UCLA", "Berkeley", "other"),
    founded = c(1919, 1868, 1),
    # Test NA in the URL
    website_naurl = xl_hyperlink_cell(url = c("http://www.ucla.edu", "http://www.berkeley.edu", NA), "website"),
    # Test NA in the name
    website_naname = xl_hyperlink_cell(url = c("http://www.ucla.edu", "http://www.berkeley.edu", NA), c("website", NA, NA)),
    # Test no name given
    website_noname = xl_hyperlink_cell(url = c("http://www.ucla.edu", "http://www.berkeley.edu", NA))
  )
  # currently readxl does not support URLs so inspect manually
  file_url_cell <- write_xlsx(df)
  expect_true(file.exists(file_url_cell))

  # Ensure that we cannot crash R with intentionally-malformed objects
  # NULL name
  null_cell <- xl_hyperlink_cell(url = c("http://www.ucla.edu", "http://www.berkeley.edu"), name = c("A", "B"))
  null_cell[[1]]$name <- NULL
  df <- data.frame(null_col = null_cell)
  file_url_cell <- write_xlsx(df)
  expect_true(file.exists(file_url_cell))

  # Numeric URL
  num_url <- xl_hyperlink_cell(url = c("http://www.ucla.edu", "http://www.berkeley.edu"), name = c("A", "B"))
  num_url[[1]]$url <- 5
  df <- data.frame(num_url = num_url)
  expect_error(
    file_url_cell <- write_xlsx(df),
    regexp = "STRING_ELT() can only be applied to a 'character vector', not a 'double'",
    fixed = TRUE
  )
  expect_true(file.exists(file_url_cell))
})

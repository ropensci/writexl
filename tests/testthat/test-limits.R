test_that("C_set_tempdir rejects paths that are too long", {
  long_path <- strrep("a", 2048)
  expect_error(
    writexl:::C_set_tempdir(long_path),
    "tempdir path too long"
  )
})

test_that("write_xlsx errors informatively for too many columns", {
  df_wide <- as.data.frame(matrix(1L, nrow = 1, ncol = 16385))
  expect_error(
    write_xlsx(df_wide),
    "too many columns"
  )
})

test_that("write_xlsx errors informatively for too many rows with header", {
  # 1048576 data rows + 1 header row = 1048577 total, exceeds xlsx limit of 1048576
  df_tall <- data.frame(x = integer(1048576))
  expect_error(
    write_xlsx(df_tall, col_names = TRUE),
    "too many rows"
  )
})

test_that("write_xlsx allows exactly LXW_ROW_MAX rows when col_names=FALSE", {
  skip_on_cran()
  df_max <- data.frame(x = integer(1048576))
  expect_true(file.exists(write_xlsx(df_max, col_names = FALSE)))
})

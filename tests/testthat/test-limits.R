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

test_that("C_set_tempdir rejects paths that are too long", {
  long_path <- strrep("a", 2048)
  expect_error(
    writexl:::C_set_tempdir(long_path),
    "tempdir path too long"
  )
})

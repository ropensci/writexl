# Merged cells.  worksheet_merge_range() writes the top-left text itself and
# blanks the rest, so a merge carries its own content and is applied after the
# sheet's rows.

merged <- function(..., df = data.frame(a = 1:3, b = 4:6)) {
  x <- xlsx_part(write_tmp(list(D = xl_sheet(df, merge = list(...)))),
                 "xl/worksheets/sheet1.xml", raw = TRUE)
  m <- regmatches(x, regexpr("<mergeCells.*?</mergeCells>", x))
  if (!length(m)) NA_character_ else m
}

# ── Construction ──────────────────────────────────────────────────────────────

test_that("xl_merge validates its arguments", {
  expect_error(xl_merge(), "must name the cells")
  expect_error(xl_merge("A1:B1", text = 1), "single non-NA string")
  expect_error(xl_merge("A1:B1", text = c("a", "b")), "single non-NA string")
  expect_error(xl_merge("A1:B1", format = "bold"), "must be an xl_format")
  expect_s3_class(xl_merge("A1:B1"), "xl_merge")
  expect_s3_class(xl_merge("A1:B1", "T", xl_font(bold = TRUE)), "xl_merge")
})

test_that("the print method runs", {
  expect_output(print(xl_merge("A1:C1", "Title")), "xl_merge")
  expect_output(print(xl_merge("A1:C1", "Title")), "A1:C1", fixed = TRUE)
})

# ── Writing ───────────────────────────────────────────────────────────────────

test_that("a merge is recorded in mergeCells", {
  expect_match(merged(xl_merge("A5:B5", "Total")), '<mergeCell ref="A5:B5"/>',
               fixed = TRUE)
})

test_that("several merges are all recorded", {
  x <- merged(xl_merge("A5:B5", "Total"), xl_merge("A6:B6", "Note"))
  expect_match(x, 'count="2"', fixed = TRUE)
  expect_match(x, 'ref="A5:B5"', fixed = TRUE)
  expect_match(x, 'ref="A6:B6"', fixed = TRUE)
})

test_that("a single xl_merge need not be wrapped in a list", {
  x <- xlsx_part(write_tmp(list(D = xl_sheet(data.frame(a = 1:3),
                                             merge = xl_merge("A5:B5", "T")))),
                 "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_match(x, 'ref="A5:B5"', fixed = TRUE)
})

test_that("the merged text is written and readable", {
  skip_if_not_installed("readxl")
  p <- write_tmp(list(D = xl_sheet(data.frame(a = 1:3, b = 4:6),
                                   merge = xl_merge("A5:B5", "Total"))))
  rd <- as.data.frame(readxl::read_xlsx(p))
  expect_equal(rd$a[4], "Total")
})

test_that("a merge over written cells keeps only the merged value", {
  # this is Excel's own behaviour: merging discards everything but the
  # top-left value, and it is why merges are applied after the rows
  skip_if_not_installed("readxl")
  p <- write_tmp(list(D = xl_sheet(data.frame(a = 1:3, b = 4:6),
                                   merge = xl_merge("A2:B2", "Merged"))))
  rd <- as.data.frame(readxl::read_xlsx(p))
  expect_equal(rd$a[1], "Merged")   # replaced the 1 that was there
  expect_equal(rd$a[2], "2")        # later rows untouched
})

test_that("a merge format reaches styles.xml", {
  p <- write_tmp(list(D = xl_sheet(data.frame(a = 1:3),
                                   merge = xl_merge("A5:B5", "Total",
                                     format = xl_fill(background = "yellow")))))
  expect_match(xlsx_part(p, "xl/styles.xml", raw = TRUE), "FFFFFF00",
               ignore.case = TRUE)
})

test_that("a merge with no text writes an empty merged cell", {
  expect_match(merged(xl_merge("A5:B5")), 'ref="A5:B5"', fixed = TRUE)
})

test_that("a merge accepts a data-frame-relative spec", {
  df <- data.frame(a = 1:3, b = 4:6)
  x <- merged(xl_merge(list(rows = 1, cols = c("a", "b")), "Merged"), df = df)
  # data row 1 is sheet row 2 once the header is counted
  expect_match(x, 'ref="A2:B2"', fixed = TRUE)
})

# ── Validation ────────────────────────────────────────────────────────────────

test_that("a single-cell merge is refused", {
  # libxlsxwriter returns a bare parameter-validation error for this, so the
  # message is raised in R where it can name the merge
  expect_error(write_tmp(list(D = xl_sheet(data.frame(a = 1:3),
                                           merge = xl_merge("A5:A5", "x")))),
               "single cell")
  # a range string that is not a rectangle is rejected by the range parser
  expect_error(write_tmp(list(D = xl_sheet(data.frame(a = 1:3),
                                           merge = xl_merge("A5", "x")))),
               "merge")
})

test_that("non-merge objects in the list are rejected", {
  expect_error(write_tmp(list(D = xl_sheet(data.frame(a = 1),
                                           merge = list("A1:B1")))),
               "must be an xl_merge object")
  expect_error(write_tmp(list(D = xl_sheet(data.frame(a = 1), merge = "A1:B1"))),
               "must be an xl_merge object")
})

test_that("an inverted merge range is rejected by the shared parser", {
  expect_error(write_tmp(list(D = xl_sheet(data.frame(a = 1:3),
                                           merge = xl_merge("B5:A5", "x")))),
               "inverted")
})

# ── Memory mode ───────────────────────────────────────────────────────────────

test_that("a merge turns row streaming off, and says why", {
  df <- data.frame(a = 1:3, b = 4:6)
  sheets <- list(D = xl_sheet(df, merge = xl_merge("A5:B5", "Total")))
  plan <- Map(function(el, d) .resolve_sheet_plan(el, d, .new_format_registry(),
                                                  1L, xl_properties()),
              sheets, list(df))
  cm <- .resolve_constant_memory(list(df), xl_properties(), plan)
  expect_equal(cm$on, 0L)
  expect_match(cm$reasons, "merged range")

  # observable end to end: with streaming off, strings go to the shared table
  p <- write_tmp(sheets)
  expect_false(is.null(xlsx_part(p, "xl/sharedStrings.xml")))
  # ... and a sheet with no merge keeps streaming on
  expect_true(is.null(xlsx_part(write_tmp(list(D = xl_sheet(df)),
                                          constant_memory = TRUE),
                                "xl/sharedStrings.xml")))
})

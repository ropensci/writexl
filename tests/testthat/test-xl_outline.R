# Outline symbol display, and the cell error indicators Excel can be told to
# stop drawing.  Both are worksheet-level settings with no cell data of their
# own, so the tests read the worksheet part directly.

ws <- function(...) {
  tmp <- tempfile(fileext = ".xlsx")
  write_xlsx(list(D = xl_sheet(data.frame(a = 1:3, b = c("x", "y", "z"),
                                          stringsAsFactors = FALSE), ...)),
             tmp)
  xlsx_part(tmp, "xl/worksheets/sheet1.xml", raw = TRUE)
}

# ── Outline ───────────────────────────────────────────────────────────────────

test_that("outline settings reach sheetPr", {
  w <- ws(cols = xl_col_spec("a", level = 1),
          outline = xl_outline(symbols_below = FALSE, symbols_right = FALSE))
  expect_match(w, 'summaryBelow="0"', fixed = TRUE)
  expect_match(w, 'summaryRight="0"', fixed = TRUE)
})

test_that("the defaults match Excel's, so nothing is written for them", {
  # Excel's own defaults are below/right, so writing them would be noise
  w <- ws(cols = xl_col_spec("a", level = 1), outline = xl_outline())
  expect_no_match(w, 'summaryBelow="0"', fixed = TRUE)
  expect_no_match(w, 'summaryRight="0"', fixed = TRUE)
})

test_that("outline display is independent of the grouping itself", {
  # hiding the symbols must not disturb the outlineLevel that creates the group
  w <- ws(cols = xl_col_spec("a", level = 2), outline = xl_outline(visible = FALSE))
  expect_match(w, 'outlineLevel="2"', fixed = TRUE)
})

test_that("outline arguments are validated", {
  expect_error(xl_outline(visible = NA), "`visible` must be TRUE or FALSE")
  expect_error(xl_outline(symbols_below = "yes"),
               "`symbols_below` must be TRUE or FALSE")
  expect_error(xl_outline(auto_style = c(TRUE, FALSE)),
               "`auto_style` must be TRUE or FALSE")
  expect_error(ws(outline = "yes"), "must be an xl_outline object")
})

test_that("the outline print method runs", {
  expect_output(print(xl_outline()), "xl_outline")
  expect_output(print(xl_outline(visible = FALSE)), "visible=FALSE")
})

# ── Ignoring cell errors ──────────────────────────────────────────────────────

test_that("an ignored error reaches ignoredErrors with its range", {
  w <- ws(ignore_errors = list(number_stored_as_text = "B2:B4"))
  expect_match(w, '<ignoredError sqref="B2:B4" numberStoredAsText="1"/>',
               fixed = TRUE)
})

test_that("several types are written together", {
  w <- ws(ignore_errors = list(number_stored_as_text = "B2:B4",
                               eval_error = "A2:A4"))
  expect_match(w, 'numberStoredAsText="1"', fixed = TRUE)
  expect_match(w, 'evalError="1"', fixed = TRUE)
})

test_that("the range takes any spelling the shared resolver accepts", {
  # a data-frame-relative spec resolves against the sheet's columns, with the
  # header row accounted for -- column b, data rows 1:3, is B2:B4
  w <- ws(ignore_errors = list(number_stored_as_text = list(cols = "b")))
  expect_match(w, 'sqref="B2:B4"', fixed = TRUE)
  expect_match(ws(ignore_errors = list(eval_error = list(cols = "a", rows = 1:2))),
               'sqref="A2:A3"', fixed = TRUE)
})

test_that("every documented error type is accepted", {
  # the docs quantify over the whole list, so enumerate it rather than
  # spot-checking: a newly added type fails here until it is covered
  for (ty in names(.LXW_IGNORE_TYPE)) {
    args <- list(ignore_errors = stats::setNames(list("A2:A4"), ty))
    expect_match(do.call(ws, args), "<ignoredError ", fixed = TRUE, label = ty)
  }
  expect_equal(length(.LXW_IGNORE_TYPE), 9L)
})

test_that("ignore_errors arguments are validated", {
  expect_error(ws(ignore_errors = list("A1:A2")), "must be a named list")
  expect_error(ws(ignore_errors = list(nope = "A1:A2")),
               "unknown error type\\(s\\): nope")
  # the message lists what would have worked
  expect_error(ws(ignore_errors = list(nope = "A1:A2")), "number_stored_as_text")
  expect_error(ws(ignore_errors = list(eval_error = "A1", eval_error = "B1")),
               "only once")
  expect_error(ws(ignore_errors = list(eval_error = "not a range")),
               "range must look like")
})

test_that("no outline or ignored errors means neither element is written", {
  w <- ws()
  expect_no_match(w, "<ignoredErrors>", fixed = TRUE)
  expect_no_match(w, "outlinePr", fixed = TRUE)
})

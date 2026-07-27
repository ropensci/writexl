# Sheet tab state.  These are the only worksheet settings whose rules span the
# whole workbook, so most of these fixtures need more than one sheet.

one <- function() data.frame(x = 1L)

wb_xml <- function(sheets) {
  xlsx_part(write_tmp(sheets), "xl/workbook.xml", raw = TRUE)
}

# ── What reaches the file ─────────────────────────────────────────────────────

test_that("a hidden sheet is marked hidden in workbook.xml", {
  x <- wb_xml(list(A = xl_sheet(one()), B = xl_sheet(one(), visible = FALSE)))
  expect_match(x, '<sheet name="B" sheetId="2" state="hidden"', fixed = TRUE)
  # the visible one is untouched
  expect_match(x, '<sheet name="A" sheetId="1" r:id="rId1"/>', fixed = TRUE)
})

test_that("active and first_tab reach workbookView", {
  x <- wb_xml(list(A = xl_sheet(one()),
                   B = xl_sheet(one(), active = TRUE),
                   C = xl_sheet(one(), first_tab = TRUE)))
  expect_match(x, 'activeTab="1"', fixed = TRUE)   # 0-based: sheet B
  expect_match(x, 'firstSheet="2"', fixed = TRUE)  # 0-based: sheet C
})

test_that("selected marks the tab without making it active", {
  x <- xlsx_part(write_tmp(list(A = xl_sheet(one()),
                                B = xl_sheet(one(), selected = TRUE))),
                 "xl/worksheets/sheet2.xml", raw = TRUE)
  expect_match(x, 'tabSelected="1"', fixed = TRUE)
})

test_that("visible = TRUE is not the same as unset, and neither hides", {
  x <- wb_xml(list(A = xl_sheet(one(), visible = TRUE), B = xl_sheet(one())))
  expect_false(grepl("state=", x, fixed = TRUE))
})

test_that("sheet view settings leave the cells alone", {
  skip_if_not_installed("readxl")
  df <- data.frame(a = 1:3, stringsAsFactors = FALSE)
  p <- write_tmp(list(A = xl_sheet(df),
                      B = xl_sheet(df, active = TRUE, first_tab = TRUE)))
  expect_equal(as.data.frame(readxl::read_xlsx(p, sheet = "B")), df)
})

# ── The cross-sheet rules ─────────────────────────────────────────────────────

test_that("hiding every sheet is refused", {
  expect_error(write_tmp(list(A = xl_sheet(one(), visible = FALSE),
                              B = xl_sheet(one(), visible = FALSE))),
               "at least one visible sheet")
})

test_that("a hidden sheet cannot also be active or selected", {
  expect_error(write_tmp(list(A = xl_sheet(one()),
                              B = xl_sheet(one(), visible = FALSE,
                                           active = TRUE))),
               "hidden but also marked active/selected")
  expect_error(write_tmp(list(A = xl_sheet(one()),
                              B = xl_sheet(one(), visible = FALSE,
                                           selected = TRUE))),
               "hidden but also marked active/selected")
})

test_that("only one sheet may be active", {
  expect_error(write_tmp(list(A = xl_sheet(one(), active = TRUE),
                              B = xl_sheet(one(), active = TRUE))),
               "only one sheet may be active")
})

test_that("the first sheet cannot be hidden unless another is activated", {
  # Excel opens on the first sheet, so hiding it with nothing else active
  # produces a workbook with no sheet to show
  expect_error(write_tmp(list(A = xl_sheet(one(), visible = FALSE),
                              B = xl_sheet(one()))),
               "cannot be hidden unless another sheet")
  # ... and it is allowed once another sheet is active
  x <- wb_xml(list(A = xl_sheet(one(), visible = FALSE),
                   B = xl_sheet(one(), active = TRUE)))
  expect_match(x, '<sheet name="A" sheetId="1" state="hidden"', fixed = TRUE)
  expect_match(x, 'activeTab="1"', fixed = TRUE)
})

test_that("error messages name the offending sheet", {
  expect_error(write_tmp(list(Data = xl_sheet(one()),
                              Notes = xl_sheet(one(), visible = FALSE,
                                               active = TRUE))),
               '"Notes"', fixed = TRUE)
  expect_error(write_tmp(list(First = xl_sheet(one(), active = TRUE),
                              Second = xl_sheet(one(), active = TRUE))),
               '"First" and "Second"', fixed = TRUE)
})

test_that("unnamed sheets are identified by position in errors", {
  expect_error(write_tmp(list(xl_sheet(one(), active = TRUE),
                              xl_sheet(one(), active = TRUE))),
               "sheets 1 and 2")
})

# ── Interaction with plain data frames ────────────────────────────────────────

test_that("plain data frames are treated as unset and never trip the rules", {
  # a bare data frame has no tab settings, so it must not be read as hidden
  expect_silent(write_tmp(list(A = one(), B = one())))
  # mixed with an xl_sheet
  expect_silent(write_tmp(list(A = one(), B = xl_sheet(one(), active = TRUE))))
  # a single plain data frame is not "every sheet hidden"
  expect_silent(write_tmp(one()))
})

test_that("the flags are validated as logicals", {
  expect_error(xl_sheet(one(), active = "yes"), "active")
  expect_error(xl_sheet(one(), visible = 1), "visible")
  expect_error(xl_sheet(one(), first_tab = c(TRUE, TRUE)), "first_tab")
})

# ── Zero display, direction, selection and scroll position ───────────────────

view_xml <- function(...) {
  xlsx_part(write_tmp(list(D = xl_sheet(data.frame(x = 1:5), ...))),
            "xl/worksheets/sheet1.xml", raw = TRUE)
}

test_that("hide_zero and right_to_left reach sheetView", {
  expect_match(view_xml(hide_zero = TRUE), 'showZeros="0"', fixed = TRUE)
  expect_match(view_xml(right_to_left = TRUE), 'rightToLeft="1"', fixed = TRUE)
  # FALSE is not TRUE: nothing is written
  expect_false(grepl("showZeros", view_xml(hide_zero = FALSE), fixed = TRUE))
})

test_that("selection accepts a single cell and a range", {
  expect_match(view_xml(selection = "B2"), 'sqref="B2"', fixed = TRUE)
  x <- view_xml(selection = "B2:C4")
  expect_match(x, 'sqref="B2:C4"', fixed = TRUE)
  # the top-left corner becomes the active cell
  expect_match(x, 'activeCell="B2"', fixed = TRUE)
})

test_that("selection accepts a data-frame-relative spec", {
  df <- data.frame(a = 1:5, b = 1:5)
  x <- xlsx_part(write_tmp(list(D = xl_sheet(df, selection = list(rows = 1:2,
                                                                 cols = "a")))),
                 "xl/worksheets/sheet1.xml", raw = TRUE)
  # data rows 1:2 are sheet rows 2:3 once the header is counted
  expect_match(x, 'sqref="A2:A3"', fixed = TRUE)
})

test_that("an inverted selection is rejected by the shared range parser", {
  # Excel would read the corner order as naming the active cell; writexl gives
  # that up in exchange for one strict range parser, and says so rather than
  # silently normalising
  expect_error(write_tmp(list(D = xl_sheet(data.frame(x = 1:5),
                                           selection = "C4:B2"))),
               "inverted")
})

test_that("top_left sets the scrolled-to cell", {
  expect_match(view_xml(top_left = "A3"), 'topLeftCell="A3"', fixed = TRUE)
  expect_match(view_xml(top_left = "C10"), 'topLeftCell="C10"', fixed = TRUE)
})

test_that("selection and top_left are independent", {
  x <- view_xml(selection = "B2", top_left = "A3")
  expect_match(x, 'topLeftCell="A3"', fixed = TRUE)
  expect_match(x, 'sqref="B2"', fixed = TRUE)
})

test_that("the view scalars are validated", {
  expect_error(xl_sheet(data.frame(x = 1), hide_zero = "yes"), "hide_zero")
  expect_error(xl_sheet(data.frame(x = 1), right_to_left = 1), "right_to_left")
  expect_error(write_tmp(list(D = xl_sheet(data.frame(x = 1),
                                           selection = "not a range"))),
               "selection")
})

test_that("the view scalars leave the cells alone", {
  skip_if_not_installed("readxl")
  df <- data.frame(a = 1:5, stringsAsFactors = FALSE)
  p <- write_tmp(list(D = xl_sheet(df, hide_zero = TRUE, right_to_left = TRUE,
                                   selection = "A2", top_left = "A2")))
  expect_equal(as.data.frame(readxl::read_xlsx(p)), df)
})

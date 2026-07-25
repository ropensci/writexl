# Tests for xl_sheet() autofilter and worksheet protection.

sheet1_xml <- function(x) {
  tmp <- tempfile(fileext = ".xlsx")
  write_xlsx(x, tmp)
  xlsx_part(tmp, "xl/worksheets/sheet1.xml", raw = TRUE)
}

test_that("autofilter = TRUE covers the whole used range", {
  df <- data.frame(name = c("a", "b", "c"), value = 1:3)
  w <- sheet1_xml(list(D = xl_sheet(df, autofilter = TRUE)))
  expect_match(w, '<autoFilter ref="A1:B4"')   # header + 3 data rows, 2 cols
})

test_that("autofilter accepts an explicit range", {
  df <- data.frame(a = 1:3, b = 4:6, c = 7:9)
  w <- sheet1_xml(list(D = xl_sheet(df, autofilter = "A1:C4")))
  expect_match(w, '<autoFilter ref="A1:C4"')
})

test_that("autofilter is absent by default and validated", {
  w <- sheet1_xml(data.frame(a = 1:2))
  expect_false(grepl("autoFilter", w))
  expect_error(xl_sheet(data.frame(a = 1), autofilter = 1), "TRUE/FALSE or an Excel range")
  expect_error(xl_sheet(data.frame(a = 1), autofilter = c(TRUE, FALSE)), "TRUE/FALSE")
})

test_that(".parse_range parses and rejects", {
  expect_equal(writexl:::.parse_range("A1:D51"), c(0L, 0L, 50L, 3L))
  expect_equal(writexl:::.parse_range("B2:B2"), c(1L, 1L, 1L, 1L))
  expect_error(writexl:::.parse_range("A1"), "A1:D51")
  expect_error(writexl:::.parse_range("A1:9"), "A1:D51")
})

test_that("protect = TRUE adds default sheet protection", {
  w <- sheet1_xml(list(D = xl_sheet(data.frame(a = 1:2), protect = TRUE)))
  expect_match(w, "<sheetProtection")
  expect_match(w, 'sheet="1"')
})

test_that("protect = password sets a hashed password", {
  w <- sheet1_xml(list(D = xl_sheet(data.frame(a = 1:2), protect = "secret")))
  expect_match(w, "<sheetProtection")
  expect_match(w, "password=")
})

test_that("protect = list gives fine-grained control", {
  w <- sheet1_xml(list(D = xl_sheet(data.frame(a = 1:2),
    protect = list(password = "pw", format_cells = TRUE,
                   insert_rows = TRUE, sort = TRUE))))
  # a libxlsxwriter option set TRUE -> the action is allowed (attr = "0")
  expect_match(w, 'formatCells="0"')
  expect_match(w, 'insertRows="0"')
  expect_match(w, 'sort="0"')
  expect_match(w, "password=")
})

test_that("protect list without options uses defaults", {
  w <- sheet1_xml(list(D = xl_sheet(data.frame(a = 1:2),
                                    protect = list(password = "pw"))))
  expect_match(w, "<sheetProtection")
  expect_match(w, "password=")
})

test_that("protection is absent by default and validated", {
  w <- sheet1_xml(data.frame(a = 1:2))
  expect_false(grepl("sheetProtection", w))
  expect_error(xl_sheet(data.frame(a = 1), protect = list(1, 2)), "fully named")
  expect_error(xl_sheet(data.frame(a = 1), protect = list(bogus = TRUE)), "unknown")
  expect_error(xl_sheet(data.frame(a = 1), protect = 42), "TRUE/FALSE")
})

test_that("protect = NULL means no protection", {
  expect_silent(xl_sheet(data.frame(a = 1), protect = NULL))
  w <- sheet1_xml(list(D = xl_sheet(data.frame(a = 1:2), protect = NULL)))
  expect_false(grepl("sheetProtection", w))
})

test_that("cell locking interacts with sheet protection", {
  # a cell explicitly unlocked, on a protected sheet
  df <- data.frame(a = 1)
  df$b <- xl_cell_general(value = 2, format = xl_protection(locked = FALSE))
  w <- sheet1_xml(list(D = xl_sheet(df, protect = TRUE)))
  expect_match(w, "<sheetProtection")     # sheet protected
  s <- styles_string(list(D = xl_sheet(df, protect = TRUE)))
  expect_match(s, 'locked="0"')           # the cell is unlocked in the style
})

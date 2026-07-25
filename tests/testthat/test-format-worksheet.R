# Tests for the worksheet layer: xl_sheet(), xl_col_spec(), xl_row_spec().

sheet_xml <- function(x, ...) {
  tmp <- tempfile(fileext = ".xlsx")
  write_xlsx(x, tmp, ...)
  xlsx_part(tmp, "xl/worksheets/sheet1.xml", raw = TRUE)
}

test_that("xl_col_spec / xl_row_spec are xl_format subclasses", {
  cs <- xl_col_spec("a", width = 12, format = xl_num_format("0.0"))
  expect_s3_class(cs, "xl_col_spec")
  expect_s3_class(cs, "xl_colrow_spec")
  expect_s3_class(cs, "xl_format")
  expect_equal(attr(cs, "xl_target")$index, "a")
  expect_equal(attr(cs, "xl_geometry")$width, 12)

  rs <- xl_row_spec(1, height = 20)
  expect_s3_class(rs, "xl_row_spec")
  expect_equal(attr(rs, "xl_geometry")$height, 20)
})

test_that("specs combine with + preserving target, geometry, and class", {
  cs <- xl_col_spec("revenue", width = 14) + xl_font(bold = TRUE)
  expect_s3_class(cs, "xl_col_spec")
  expect_equal(attr(cs, "xl_target")$index, "revenue")
  expect_equal(attr(cs, "xl_geometry")$width, 14)
  expect_true(unclass(cs)$font$bold)
  # a combined format (via +) can be passed as `format`
  cs2 <- xl_col_spec("x", width = 10,
                     format = xl_font(italic = TRUE) + xl_fill(background = "yellow"))
  expect_true(unclass(cs2)$font$italic)
  expect_equal(unclass(cs2)$fill$background, xl_color("yellow"))
})

test_that("print methods run", {
  expect_output(print(xl_col_spec("a", width = 12)), "target")
  expect_output(print(xl_col_spec("a", width = 12)), "geometry")
  expect_output(print(xl_sheet(data.frame(a = 1), cols = xl_col_spec("a", width = 9),
                               rows = xl_row_spec(1, height = 9), freeze = "A2")),
                "xl_sheet")
})

test_that("spec constructors validate their inputs", {
  expect_error(xl_col_spec(), "at least one column")
  expect_error(xl_col_spec(TRUE), "character .* or numeric")
  expect_error(xl_row_spec(), "at least one row")
  expect_error(xl_row_spec("a"), "numeric")
  expect_error(xl_row_spec(0), "positive")
  expect_error(xl_col_spec("a", width = -1), "between 0")
  expect_error(xl_col_spec("a", level = 9), "between 0 and 7")
  expect_error(xl_col_spec("a", format = "x"), "must be an xl_format")
})

test_that("xl_sheet validates data and spec lists", {
  expect_error(xl_sheet(1:3), "must be a data frame")
  expect_error(xl_sheet(data.frame(a = 1), cols = "x"),
               "must be an xl_col_spec")
  expect_error(xl_sheet(data.frame(a = 1), rows = list(xl_col_spec("a"))),
               "must be an xl_row_spec")
  s <- xl_sheet(data.frame(a = 1), cols = list(xl_col_spec("a", width = 5)))
  expect_length(s$cols, 1L)
})

test_that("freeze parsing handles cell refs, lists, and errors", {
  expect_equal(writexl:::.parse_freeze(NULL), c(-1L, -1L))
  expect_equal(writexl:::.parse_freeze(NA), c(-1L, -1L))
  expect_equal(writexl:::.parse_freeze("A2"), c(1L, 0L))
  expect_equal(writexl:::.parse_freeze("B3"), c(2L, 1L))
  expect_equal(writexl:::.parse_freeze("AA1"), c(0L, 26L))
  expect_equal(writexl:::.parse_freeze(list(row = 2, col = 1)), c(2L, 1L))
  expect_error(writexl:::.parse_freeze("nonsense"), "cell reference")
  expect_error(writexl:::.parse_freeze(42), "cell reference")
})

test_that("column targeting resolves names and positions", {
  expect_equal(writexl:::.resolve_col_index(c("b", "a"), c("a", "b")), c(2L, 1L))
  expect_equal(writexl:::.resolve_col_index(c(1, 3), c("a", "b", "c")), c(1L, 3L))
  expect_error(writexl:::.resolve_col_index("z", c("a", "b")), "unknown column")
  expect_error(writexl:::.resolve_col_index(9, c("a", "b")), "out of range")
})

test_that("column width and format reach the worksheet/styles xml", {
  s <- xl_sheet(data.frame(name = c("a", "b"), revenue = c(1.5, 2.5)),
                cols = xl_col_spec("revenue", width = 16,
                                   format = xl_num_format("#,##0.00")))
  w <- sheet_xml(list(D = s))
  expect_match(w, 'customWidth="1"')
  expect_match(w, 'width="16')            # Excel pads the exact value
  expect_match(w, '<col min="2" max="2"')  # targeted the 2nd column
})

test_that("column targeting by position works", {
  s <- xl_sheet(data.frame(a = 1:2, b = 3:4),
                cols = xl_col_spec(1, width = 25))
  w <- sheet_xml(list(D = s))
  expect_match(w, '<col min="1" max="1"')
  expect_match(w, 'width="25')
})

test_that("row height, freeze, gridlines, tab color, and zoom are written", {
  s <- xl_sheet(data.frame(a = 1:3),
                rows = xl_row_spec(1, height = 30),
                freeze = "A2", gridlines = FALSE,
                tab_color = "red", zoom = 120)
  w <- sheet_xml(list(D = s))
  expect_match(w, 'ht="30"')
  expect_match(w, "<pane")
  expect_match(w, 'showGridLines="0"')
  expect_match(w, "tabColor")
  expect_match(w, 'zoomScale="120"')
})

test_that("hidden columns and outline levels are written", {
  s <- xl_sheet(data.frame(a = 1:2, b = 3:4),
                cols = xl_col_spec("b", hidden = TRUE, level = 1))
  w <- sheet_xml(list(D = s))
  expect_match(w, 'hidden="1"')
  expect_match(w, 'outlineLevel="1"')
})

test_that("plain data frames still work and Date columns stay formatted", {
  # regression: date formatting moved from C to the R sheet plan
  df <- data.frame(d = as.Date("2020-01-01") + 0:2, x = 1:3)
  tmp <- tempfile(fileext = ".xlsx")
  write_xlsx(df, tmp)
  s <- xlsx_part(tmp, "xl/styles.xml", raw = TRUE)
  expect_match(s, "yyyy", fixed = TRUE)
  expect_equal(as.Date(readxl::read_xlsx(tmp)$d), df$d)
})

test_that("a sheet's default row height is applied", {
  s <- xl_sheet(data.frame(a = 1:3), default_row_height = 22)
  w <- sheet_xml(list(D = s))
  expect_match(w, 'defaultRowHeight="22"')
})

test_that("xl_sheet mixes with plain data frames in one workbook", {
  wb <- list(
    Styled = xl_sheet(data.frame(a = 1.5), cols = xl_col_spec("a", width = 12)),
    Plain  = data.frame(b = 1:2)
  )
  tmp <- tempfile(fileext = ".xlsx")
  expect_silent(write_xlsx(wb, tmp))
  expect_equal(readxl::read_xlsx(tmp, sheet = "Plain")$b, 1:2)
})

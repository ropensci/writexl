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
  expect_equal(.parse_freeze(NULL), c(-1L, -1L))
  expect_equal(.parse_freeze(NA), c(-1L, -1L))
  expect_equal(.parse_freeze("A2"), c(1L, 0L))
  expect_equal(.parse_freeze("B3"), c(2L, 1L))
  expect_equal(.parse_freeze("AA1"), c(0L, 26L))
  expect_equal(.parse_freeze(list(row = 2, col = 1)), c(2L, 1L))
  expect_error(.parse_freeze("nonsense"), "cell reference")
  expect_error(.parse_freeze(42), "cell reference")
})

test_that("column targeting resolves names and positions", {
  expect_equal(.resolve_col_index(c("b", "a"), c("a", "b")), c(2L, 1L))
  expect_equal(.resolve_col_index(c(1, 3), c("a", "b", "c")), c(1L, 3L))
  expect_error(.resolve_col_index("z", c("a", "b")), "unknown column")
  expect_error(.resolve_col_index(9, c("a", "b")), "out of range")
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

test_that("geometry can be given in pixels instead", {
  # converted with libxlsxwriter's own formulas, so the width Excel stores
  # round-trips back to the pixel count asked for: 100 px -> 13.571 chars,
  # stored as 13.571 + 5/7 = 14.2857, which renders at 100 px again
  expect_equal(.pixels_to_width(100), (100 - 5) / 7)
  expect_equal(.pixels_to_height(40), 30)
  # the two documented special cases: the pixel defaults map to the unit
  # defaults exactly, rather than through the general formula
  expect_equal(.pixels_to_width(64), 8.43)
  expect_equal(.pixels_to_height(20), 15)
  expect_equal(.pixels_to_width(12), 1)      # the <= 12 branch

  s <- xl_sheet(data.frame(a = 1:2, b = 3:4),
                cols = xl_col_spec("a", width_pixels = 100),
                rows = xl_row_spec(1, height_pixels = 40))
  w <- sheet_xml(list(D = s))
  expect_match(w, 'width="14.28515625"', fixed = TRUE)
  expect_match(w, 'ht="30"', fixed = TRUE)
})

test_that("a geometry may not be given in both units", {
  expect_error(xl_col_spec("a", width = 10, width_pixels = 100),
               "either `width` or `width_pixels`, not both")
  expect_error(xl_row_spec(1, height = 10, height_pixels = 40),
               "either `height` or `height_pixels`, not both")
  # each on its own is fine
  expect_s3_class(xl_col_spec("a", width_pixels = 100), "xl_col_spec")
  expect_s3_class(xl_row_spec(1, height_pixels = 40), "xl_row_spec")
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

test_that("auto_colwidth sizes columns to their content", {
  df <- data.frame(x = c("a", "bb"), header_is_long = c(1L, 2L))
  w <- sheet_xml(list(D = xl_sheet(df, auto_colwidth = TRUE)))
  expect_match(w, '<col min="1" max="1" width="3')    # header "x", max value "bb"
  expect_match(w, '<col min="2" max="2" width="15')   # 14-char header
})

test_that("explicit width overrides auto_colwidth", {
  df <- data.frame(text = c("a very long value here indeed", "x"))
  w <- sheet_xml(list(D = xl_sheet(df, auto_colwidth = TRUE,
                                   cols = xl_col_spec("text", width = 5))))
  expect_match(w, '<col min="1" max="1" width="5')
})

test_that("auto_colwidth handles Date and general columns", {
  d <- data.frame(when = as.Date("2020-01-01") + 0:1)     # "2020-01-01" = 10
  expect_match(sheet_xml(list(D = xl_sheet(d, auto_colwidth = TRUE))),
               '<col min="1" max="1" width="11')
  g <- data.frame(a = 1:2)
  g$b <- xl_cell_general(value = c(1234.5, 6.7))           # "1234.5" = 6
  expect_match(sheet_xml(list(D = xl_sheet(g, auto_colwidth = TRUE))),
               '<col min="2" max="2" width="7')
})

test_that("auto_colwidth is validated", {
  expect_error(xl_sheet(data.frame(a = 1), auto_colwidth = "yes"), "TRUE or FALSE")
  expect_error(xl_sheet(data.frame(a = 1), auto_colwidth = NA), "TRUE or FALSE")
})

test_that("cell display width covers value/formula/hyperlink/blank", {
  gen <- c(
    xl_cell_general(value = 12345),
    xl_cell_general(formula = "=SUM(A1:A9)"),
    xl_cell_general(hyperlink = "http://example.com"),
    xl_cell_general(hyperlink = list(url = "http://foo.com")),
    xl_cell_general(value = NA)
  )
  w <- .content_nchar(gen)
  expect_equal(w[1], 5L)
  expect_equal(w[2], nchar("=SUM(A1:A9)"))
  expect_equal(w[3], nchar("http://example.com"))
  expect_equal(w[4], nchar("http://foo.com"))
  expect_equal(w[5], 0L)
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

# --- sheet overlays ---------------------------------------------------------
#
# Range-scoped features are applied through one payload list dispatched on
# `kind` in C.  The autofilter is the first kind (its own behavior is covered
# in test-format-protect.R); these tests cover the mechanism itself.

test_that("the autofilter reaches C as an overlay payload", {
  plan <- .resolve_sheet_plan(xl_sheet(data.frame(a = 1:3), autofilter = TRUE),
                              data.frame(a = 1:3), .new_format_registry(), 1L,
                              xl_properties())
  expect_equal(plan$overlay,
               list(list(kind = "autofilter", range = c(0L, 0L, 3L, 0L))))
  expect_null(plan$autofilter)   # the old scalar slot is gone
})

test_that("a sheet with no range features carries no overlays", {
  plan <- .resolve_sheet_plan(xl_sheet(data.frame(a = 1:3)), data.frame(a = 1:3),
                              .new_format_registry(), 1L, xl_properties())
  expect_equal(plan$overlay, list())
})

test_that("an overlay set on the sheet is applied", {
  s <- xl_sheet(data.frame(a = 1:3))
  s$overlay <- list(kind = "autofilter", range = c(0L, 0L, 3L, 0L))
  expect_match(sheet_xml(list(D = s)), '<autoFilter ref="A1:A4"')
})

test_that("an unknown overlay kind is an error, not a silent no-op", {
  s <- xl_sheet(data.frame(a = 1:3))
  s$overlay <- list(kind = "bogus")
  expect_error(write_tmp(list(D = s)), "unknown sheet overlay kind 'bogus'")
})

test_that("an overlay payload without a kind is an error", {
  s <- xl_sheet(data.frame(a = 1:3))
  s$overlay <- list(list(range = c(0L, 0L, 1L, 0L)))
  expect_error(write_tmp(list(D = s)), "missing a 'kind'")
})

test_that("an autofilter overlay without a usable range is an error", {
  s <- xl_sheet(data.frame(a = 1:3))
  s$overlay <- list(kind = "autofilter", range = 1:2)
  expect_error(write_tmp(list(D = s)), "length-4 range")
})

test_that("overlay payload lists are normalised", {
  one <- list(kind = "autofilter", range = 1:4)
  expect_equal(.as_overlay_list(NULL), list())
  expect_equal(.as_overlay_list(one), list(one))
  expect_equal(.as_overlay_list(list(one, one)), list(one, one))
  expect_error(.as_overlay_list("x"), "overlay payload")
})

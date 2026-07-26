# Tests for the xl_comment object and its format verification.

test_that("xl_comment builds a payload from text and options", {
  cm <- xl_comment("hello", author = "QA", width = 200, height = 90,
                   x_scale = 1.5, y_scale = 2, start_row = 3, start_col = 4,
                   x_offset = 5, y_offset = 6)
  p <- unclass(cm)
  expect_true(is_xl_comment(cm))
  expect_equal(p$text, "hello")
  expect_equal(p$author, "QA")
  expect_equal(p$width, 200L)
  expect_equal(p$height, 90L)
  expect_equal(p$x_scale, 1.5)
  expect_equal(p$y_scale, 2)
  expect_equal(p$start_row, 3L)
  expect_equal(p$start_col, 4L)
  expect_equal(p$x_offset, 5L)
  expect_equal(p$y_offset, 6L)
})

test_that("visible maps NA/TRUE/FALSE to the libxlsxwriter enum", {
  expect_false("visible" %in% names(unclass(xl_comment("x"))))       # NA -> unset
  expect_equal(unclass(xl_comment("x", visible = TRUE))$visible, 2L)  # VISIBLE
  expect_equal(unclass(xl_comment("x", visible = FALSE))$visible, 1L) # HIDDEN
})

test_that("styling is extracted from an xl_format (color + font only)", {
  cm <- xl_comment("note",
                   format = xl_font(name = "Arial", size = 10, family = 2) +
                            xl_fill(background = "lightyellow"))
  p <- unclass(cm)
  expect_equal(p$color, xl_color("lightyellow"))
  expect_equal(p$font_name, "Arial")
  expect_equal(p$font_size, 10)
  expect_equal(p$font_family, 2L)
})

test_that(".check_comment_format warns on unsupported properties", {
  expect_warning(xl_comment("x", format = xl_font(bold = TRUE)), "font\\$bold")
  expect_warning(xl_comment("x", format = xl_font(color = "red")), "font\\$color")
  expect_warning(xl_comment("x", format = xl_fill(foreground = "red")), "fill\\$foreground")
  expect_warning(xl_comment("x", format = xl_border(all = "thin")), "border")
  expect_warning(xl_comment("x", format = xl_align(horizontal = "center")), "align")
  expect_warning(xl_comment("x", format = xl_num_format("0.0")), "num_format")
  expect_warning(xl_comment("x", format = xl_protection(locked = FALSE)), "protection")
  expect_warning(xl_comment("x", format = xl_format(quote_prefix = TRUE)), "quote_prefix")
  expect_warning(xl_comment("x", format = xl_format(hyperlink = TRUE)), "hyperlink")
})

test_that("solid pattern is silent, non-solid pattern warns", {
  expect_no_warning(xl_comment("x", format = xl_fill(background = "yellow")))  # auto solid
  expect_warning(xl_comment("x", format = xl_fill(background = "yellow", pattern = "light-gray")),
                 "fill\\$pattern")
})

test_that("supported-only format does not warn", {
  expect_no_warning(
    xl_comment("x", format = xl_font(name = "Calibri", size = 9, family = 2) +
                            xl_fill(background = "white")))
})

test_that(".comment_payload normalizes strings, xl_comment, and NA", {
  expect_equal(writexl:::.comment_payload("hi")$text, "hi")
  expect_equal(writexl:::.comment_payload(xl_comment("hey"))$text, "hey")
  expect_null(writexl:::.comment_payload(NA))
  expect_null(writexl:::.comment_payload(NA_character_))
  expect_null(writexl:::.comment_payload(NULL))
  expect_error(writexl:::.comment_payload(42), "NA, a single string, or an xl_comment")
})

test_that("xl_comment validates its inputs", {
  expect_error(xl_comment(c("a", "b")), "single non-NA string")
  expect_error(xl_comment(NA), "single non-NA string")
  expect_error(xl_comment(1), "single non-NA string")
  expect_error(xl_comment("x", format = "nope"), "must be an xl_format")
  expect_error(xl_comment("x", width = -1), "between 0")
  expect_error(xl_comment("x", visible = "yes"), "logical")
})

test_that("print.xl_comment runs", {
  expect_output(print(xl_comment("hi", author = "me")), "xl_comment")
  expect_output(print(xl_comment("hi", author = "me")), "author")
})

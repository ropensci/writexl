# Tests for cell comments: the xl_comment object, its format verification, and
# the write path (comments live in xl/comments1.xml and the VML drawing).

comments_xml <- function(x) {
  tmp <- tempfile(fileext = ".xlsx")
  write_xlsx(x, tmp)
  xlsx_part(tmp, "xl/comments1.xml", raw = TRUE)
}
vml_xml <- function(x) {
  tmp <- tempfile(fileext = ".xlsx")
  write_xlsx(x, tmp)
  xlsx_part(tmp, "xl/drawings/vmlDrawing1.vml", raw = TRUE)
}

# --- the xl_comment object ---------------------------------------------------

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
  expect_output(print(xl_comment("hi", format = xl_font(name = "Arial"))), "Arial")
})

# --- styling: the xl_format is kept, then flattened for C --------------------

test_that("the comment keeps its xl_format and stays editable", {
  cm <- xl_comment("note", format = xl_font(name = "Arial", size = 10))
  expect_true(is_xl_format(unclass(cm)$format))
  # the whole point of keeping it: a format can be adjusted after construction
  cm$format <- cm$format + xl_fill(background = "lightyellow")
  p <- writexl:::.comment_c_payload(unclass(cm))
  expect_equal(p$font_name, "Arial")
  expect_equal(p$color, xl_color("lightyellow"))
})

test_that(".comment_c_payload flattens the format to the supported C fields", {
  cm <- xl_comment("note",
                   format = xl_font(name = "Arial", size = 10, family = 2) +
                            xl_fill(background = "lightyellow"))
  p <- writexl:::.comment_c_payload(unclass(cm))
  expect_null(p$format)                     # replaced by flat fields
  expect_equal(p$text, "note")
  expect_equal(p$color, xl_color("lightyellow"))
  expect_equal(p$font_name, "Arial")
  expect_equal(p$font_size, 10)
  expect_equal(p$font_family, 2L)
  # no format -> no styling fields, and NULL passes through
  expect_null(writexl:::.comment_c_payload(unclass(xl_comment("x")))$color)
  expect_null(writexl:::.comment_c_payload(NULL))
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

# --- attaching comments to cells and writing them ---------------------------

test_that("a character vector attaches per-cell comment text", {
  df <- data.frame(x = c("a", "b", "c"))
  df$x <- xl_cell_general(value = c("a", "b", "c"),
                          comment = c("first note", NA, "third note"))
  cm <- comments_xml(df)
  expect_match(cm, "first note")
  expect_match(cm, "third note")
  expect_false(grepl("second", cm))     # NA -> no comment
})

test_that("a single xl_comment recycles to every cell", {
  df <- data.frame(x = 1:2)
  df$x <- xl_cell_general(value = 1:2, comment = xl_comment("same", author = "QA"))
  cm <- comments_xml(df)
  expect_equal(lengths(regmatches(cm, gregexpr("same", cm)))[1], 2L)
  expect_match(cm, "QA")
})

test_that("comment styling reaches comments1.xml and the VML", {
  df <- data.frame(x = 1L)
  df$x <- xl_cell_general(value = 1,
    comment = xl_comment("styled", author = "Rev", visible = TRUE, width = 200,
                         format = xl_font(name = "Arial", size = 12) +
                                  xl_fill(background = "red")))
  cm <- comments_xml(df)
  vml <- vml_xml(df)
  expect_match(cm, "Rev")
  expect_match(cm, "Arial")            # rFont
  expect_match(cm, 'val="12"')         # font size
  expect_match(vml, "FF0000", ignore.case = TRUE)  # box fill color
  expect_match(vml, "visible")         # forced visible
})

test_that("a per-cell list mixes strings, xl_comment, and NA", {
  df <- data.frame(x = 1:3)
  df$x <- xl_cell_general(
    value = 1:3,
    comment = list("plain", xl_comment("rich", author = "me"), NA))
  cm <- comments_xml(df)
  expect_match(cm, "plain")
  expect_match(cm, "rich")
  expect_match(cm, "me")
})

test_that("a comment can stand alone on an otherwise-blank cell", {
  df <- data.frame(a = 1L)
  df$note <- xl_cell_general(comment = "standalone")
  expect_match(comments_xml(df), "standalone")
})

test_that("print shows a comment marker on the cell", {
  expect_output(print(xl_cell_general(value = 1, comment = "hi")), "comment=<set>")
})

test_that("values round-trip when comments are attached", {
  df <- data.frame(x = 1:3)
  df$x <- xl_cell_general(value = c(10, 20, 30), comment = c("a", "b", "c"))
  tmp <- tempfile(fileext = ".xlsx")
  write_xlsx(df, tmp)
  expect_equal(readxl::read_xlsx(tmp)$x, c(10, 20, 30))
})

test_that("xl_cell_general validates the comment argument", {
  expect_error(xl_cell_general(value = 1, comment = 42),
               "character vector, an xl_comment, or a list")
  expect_error(xl_cell_general(value = 1, comment = list(1, 2)),
               "NA, a single string, or an xl_comment")
})

# --- sheet / workbook comment defaults --------------------------------------

test_that("sheet comment_author and show_comments apply", {
  df <- data.frame(x = 1L); df$x <- xl_cell_general(value = 1, comment = "a note")
  s <- xl_sheet(df, comment_author = "SheetAuthor", show_comments = TRUE)
  tmp <- tempfile(fileext = ".xlsx"); write_xlsx(list(D = s), tmp)
  expect_match(xlsx_part(tmp, "xl/comments1.xml", raw = TRUE), "SheetAuthor")
  expect_match(xlsx_part(tmp, "xl/drawings/vmlDrawing1.vml", raw = TRUE), "visible")
})

test_that("a per-comment author overrides the sheet default", {
  df <- data.frame(x = 1L)
  df$x <- xl_cell_general(value = 1, comment = xl_comment("n", author = "CellWins"))
  tmp <- tempfile(fileext = ".xlsx")
  write_xlsx(list(D = xl_sheet(df, comment_author = "SheetAuthor")), tmp)
  expect_match(xlsx_part(tmp, "xl/comments1.xml", raw = TRUE), "CellWins")
})

test_that("comments work on a sheet with no comment defaults set", {
  df <- data.frame(x = 1L); df$x <- xl_cell_general(value = 1, comment = "n")
  cm <- comments_xml(list(D = xl_sheet(df)))
  expect_match(cm, "n")
  # comments are hidden by default
  expect_false(grepl("visible", vml_xml(list(D = xl_sheet(df)))))
})

test_that("show_comments is validated", {
  expect_error(xl_sheet(data.frame(a = 1), show_comments = "x"), "TRUE or FALSE")
  expect_error(xl_sheet(data.frame(a = 1), show_comments = NA), "TRUE or FALSE")
})

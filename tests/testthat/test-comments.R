# Integration tests for cell comments via xl_cell_general(comment=).
# Comments live in xl/comments1.xml (text/author/font) and the VML drawing
# (visibility, box color).

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

test_that("print shows a comment marker", {
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

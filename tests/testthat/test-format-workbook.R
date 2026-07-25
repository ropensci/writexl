# Tests for the workbook layer: xl_properties(), xl_workbook(), and the
# de-hardcoded default formats.

test_that("xl_properties validates its arguments", {
  expect_error(xl_properties(default_format = "x"), "must be an xl_format")
  expect_error(xl_properties(header_format = 1), "must be an xl_format")
  expect_error(xl_properties(custom = list(1, 2)), "must be a named list")
  expect_error(xl_properties(window_size = 100), "length-2")
  expect_error(xl_properties(names = list("=1")), "named list")
  p <- xl_properties(title = "T", author = "A")
  expect_s3_class(p, "xl_properties")
})

test_that("xl_workbook validates and stores its parts", {
  expect_error(xl_workbook(1:3), "must be a data frame")
  expect_error(xl_workbook(data.frame(a = 1), properties = list()),
               "must be an xl_properties")
  wb <- xl_workbook(data.frame(a = 1), col_names = FALSE)
  expect_s3_class(wb, "xl_workbook")
  expect_false(wb$col_names)
  expect_length(wb$sheets, 1L)
})

test_that("print methods run", {
  expect_output(print(xl_properties(title = "T", author = "A", read_only = TRUE)),
                "xl_properties")
  expect_output(print(xl_workbook(data.frame(a = 1))), "xl_workbook")
})

test_that("document metadata is written to docProps/core.xml", {
  wb <- xl_workbook(data.frame(a = 1),
                    properties = xl_properties(title = "My Report",
                                               author = "Bill",
                                               company = "HP",
                                               subject = "Q1"))
  tmp <- write_tmp(wb)
  core <- xlsx_part(tmp, "docProps/core.xml", raw = TRUE)
  expect_match(core, "My Report")
  expect_match(core, "Bill")
  expect_match(core, "Q1")
})

test_that("custom properties of every scalar type are written", {
  wb <- xl_workbook(data.frame(a = 1),
                    properties = xl_properties(custom = list(
                      Project  = "Alpha",     # string
                      Revision = 3L,          # integer
                      Score    = 9.5,         # number
                      Final    = TRUE         # boolean
                    )))
  tmp <- write_tmp(wb)
  cust <- xlsx_part(tmp, "docProps/custom.xml", raw = TRUE)
  expect_match(cust, "Project");  expect_match(cust, "Alpha")
  expect_match(cust, "Revision"); expect_match(cust, "<vt:i4>3")
  expect_match(cust, "Score");    expect_match(cust, "9.5")
  expect_match(cust, "Final");    expect_match(cust, "<vt:bool>true")
})

test_that("read_only, window size, and defined names reach workbook.xml", {
  wb <- xl_workbook(data.frame(a = 1),
                    properties = xl_properties(read_only = TRUE,
                                               window_size = c(1000L, 600L),
                                               names = list(tax = "=0.2")))
  tmp <- write_tmp(wb)
  wx <- xlsx_part(tmp, "xl/workbook.xml", raw = TRUE)
  expect_match(wx, "readOnlyRecommended")
  expect_match(wx, 'windowWidth="15000"')   # set_size converts to twips (x15)
  expect_match(wx, 'windowHeight="9000"')
  expect_match(wx, 'name="tax"')
})

test_that("default_format cascades under every cell and header", {
  wb <- xl_workbook(data.frame(a = 1),
                    properties = xl_properties(default_format = xl_font(name = "Arial")))
  s <- styles_string(wb)
  expect_match(s, "Arial")
})

test_that("header_format and hyperlink_format can be overridden", {
  wb <- xl_workbook(
    data.frame(a = 1),
    properties = xl_properties(
      header_format = xl_font(bold = TRUE, color = "white") + xl_fill(background = "navy")
    )
  )
  s <- styles_string(wb)
  expect_match(s, "FF000080", ignore.case = TRUE)  # navy header fill

  df <- data.frame(a = 1L)
  df$h <- xl_hyperlink_cell("http://x.com")
  wb2 <- xl_workbook(list(S = df),
                     properties = xl_properties(
                       hyperlink_format = xl_font(color = "red", underline = "double")))
  s2 <- styles_string(wb2)
  expect_match(s2, "FFFF0000", ignore.case = TRUE)  # red hyperlink
})

test_that("date_format can be overridden at the workbook level", {
  wb <- xl_workbook(data.frame(d = as.Date("2020-01-01")),
                    properties = xl_properties(date_format = xl_num_format("dd/mm/yyyy")))
  expect_match(styles_string(wb), "dd/mm/yyyy", fixed = TRUE)
})

test_that("workbook col_names / format_headers override write_xlsx args", {
  wb <- xl_workbook(data.frame(a = 1:2), col_names = FALSE)
  tmp <- write_tmp(wb, col_names = TRUE)   # workbook wins -> no header
  rd <- readxl::read_xlsx(tmp, col_names = FALSE)
  expect_equal(as.integer(rd[[1]]), 1:2)
})

test_that("long and duplicate sheet names are still handled", {
  long <- list(data.frame(a = 1))
  names(long) <- paste(rep("x", 40), collapse = "")
  expect_warning(write_tmp(long), "Truncating")
  dup <- list(data.frame(a = 1), data.frame(b = 2))
  names(dup) <- c("S", "S")
  expect_warning(write_tmp(dup), "Deduplicating")
})

test_that("de-hardcoded defaults match the previous behavior (regression)", {
  # plain data frame, no workbook customization
  df <- data.frame(name = c("a", "b"), when = as.Date("2020-01-01") + 0:1)
  df$link <- xl_hyperlink_cell(c("http://x.com", "http://y.com"))
  s <- styles_string(df)
  expect_match(s, "<b/>")                          # bold header
  expect_match(s, 'horizontal="center"')           # centered header
  expect_match(s, "yyyy", fixed = TRUE)            # date number format
  expect_match(s, "FF0000FF", ignore.case = TRUE)  # blue hyperlink font
})

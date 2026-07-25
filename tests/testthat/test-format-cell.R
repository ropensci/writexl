# Integration tests for cell-level formatting via xl_cell_general(format=).
# Formatting is verified by inspecting the written xl/styles.xml.

test_that("xl_cell_general accepts a single format, a list, or NULL", {
  x <- xl_cell_general(value = 1:3, format = xl_font(bold = TRUE))
  expect_true(is_xl_format(unclass(x)[[1]]$format))
  expect_true(is_xl_format(unclass(x)[[3]]$format))

  # per-cell list, recycled
  y <- xl_cell_general(value = 1:2,
                       format = list(xl_font(bold = TRUE), xl_font(italic = TRUE)))
  expect_true(unclass(unclass(y)[[1]]$format)$font$bold)
  expect_true(unclass(unclass(y)[[2]]$format)$font$italic)

  # NULL -> no format
  z <- xl_cell_general(value = 1:2)
  expect_null(unclass(z)[[1]]$format)

  expect_error(xl_cell_general(value = 1, format = list("bad")),
               "must be an xl_format")
  expect_error(xl_cell_general(value = 1, format = 42), "must be an xl_format")
})

test_that("print shows a format marker", {
  expect_output(print(xl_cell_general(value = 1, format = xl_font(bold = TRUE))),
                "format=<set>")
})

test_that("font formatting reaches styles.xml", {
  df <- data.frame(a = 1L)
  df$b <- xl_cell_general(value = "hi",
                          format = xl_font(bold = TRUE, italic = TRUE,
                                           color = "red", size = 16))
  s <- styles_string(df)
  expect_match(s, "<b/>")
  expect_match(s, "<i/>")
  expect_match(s, "FFFF0000", ignore.case = TRUE)  # red font color, ARGB
  expect_match(s, 'val="16"')
})

test_that("fill formatting reaches styles.xml", {
  df <- data.frame(a = 1L)
  df$b <- xl_cell_general(value = 1, format = xl_fill(background = "yellow"))
  s <- styles_string(df)
  expect_match(s, "<patternFill")
  expect_match(s, "FFFFFF00", ignore.case = TRUE)   # yellow
})

test_that("border formatting reaches styles.xml", {
  df <- data.frame(a = 1L)
  df$b <- xl_cell_general(value = 1, format = xl_border(all = "thin", color = "black"))
  s <- styles_string(df)
  expect_match(s, "<border")
  expect_match(s, "thin")
})

test_that("alignment formatting reaches styles.xml", {
  df <- data.frame(a = 1L)
  df$b <- xl_cell_general(value = "x",
                          format = xl_align(horizontal = "center", vertical = "top",
                                            wrap = TRUE))
  s <- styles_string(df)
  expect_match(s, 'horizontal="center"')
  expect_match(s, 'vertical="top"')
  expect_match(s, 'wrapText="1"')
})

test_that("number format reaches styles.xml", {
  df <- data.frame(a = 1L)
  df$b <- xl_cell_general(value = 1234.5, format = xl_num_format("#,##0.00"))
  s <- styles_string(df)
  expect_match(s, "#,##0.00", fixed = TRUE)
})

test_that("protection formatting reaches styles.xml", {
  df <- data.frame(a = 1L)
  df$b <- xl_cell_general(value = 1, format = xl_protection(locked = FALSE, hidden = TRUE))
  s <- styles_string(df)
  expect_match(s, 'locked="0"')
  expect_match(s, 'hidden="1"')
})

test_that("date/time general cells get the default number format merged", {
  df <- data.frame(a = 1L)
  df$d <- xl_cell_general(value = as.Date("2020-03-01"), format = xl_font(bold = TRUE))
  df$t <- xl_cell_general(value = as.POSIXct("2020-03-01 12:00", tz = "UTC"))
  s <- styles_string(df)
  expect_match(s, "yyyy", fixed = TRUE)          # date number format present
  expect_match(s, "<b/>")                        # bold preserved alongside it
  # explicit num_format on a date cell wins over the default
  df2 <- data.frame(a = 1L)
  df2$d <- xl_cell_general(value = as.Date("2020-03-01"),
                           format = xl_num_format("dd/mm/yyyy"))
  expect_match(styles_string(df2), "dd/mm/yyyy", fixed = TRUE)
})

test_that("values still round-trip with formatting applied", {
  df <- data.frame(x = 1:3)
  df$n <- xl_cell_general(value = c(1.5, 2.5, 3.5),
                          format = xl_num_format("0.0"))
  tmp <- write_tmp(df)
  rd <- readxl::read_xlsx(tmp)
  expect_equal(rd$n, c(1.5, 2.5, 3.5))
})

test_that("format= flows through xl_formula and xl_hyperlink_cell", {
  df <- data.frame(a = 1L)
  df$f <- xl_formula("=1+2", format = xl_font(bold = TRUE))
  expect_match(styles_string(df), "<b/>")

  df2 <- data.frame(a = 1L)
  df2$h <- xl_hyperlink_cell("https://example.com", value = "link",
                             format = xl_font(color = "green"))
  s2 <- styles_string(df2)
  expect_match(s2, "FF00FF00", ignore.case = TRUE)  # R "green" == #00FF00

  df3 <- data.frame(a = 1L)
  df3$h <- xl_hyperlink("https://example.com", "link", format = xl_font(italic = TRUE))
  expect_match(styles_string(df3), "<i/>")
})

test_that("wrappers coerce factor input while applying a format", {
  df <- data.frame(a = 1L)
  df$f <- xl_formula(factor("=1+1"), format = xl_font(bold = TRUE))
  expect_match(styles_string(df), "<b/>")

  df2 <- data.frame(a = 1L)
  df2$h <- xl_hyperlink(factor("http://example.com"), format = xl_font(italic = TRUE))
  expect_match(styles_string(df2), "<i/>")

  df3 <- data.frame(a = 1L)
  df3$h <- xl_hyperlink_cell(factor("http://example.com"),
                             format = xl_font(strikeout = TRUE))
  expect_match(styles_string(df3), "strike")
})

test_that("identical formats deduplicate to one style definition", {
  df <- data.frame(a = 1:50)
  df$b <- xl_cell_general(value = 1:50, format = xl_num_format("0.000"))
  s <- styles_string(df)
  hits <- gregexpr("0.000", s, fixed = TRUE)[[1]]
  expect_equal(sum(hits > 0), 1L)   # the custom number format is defined once
})

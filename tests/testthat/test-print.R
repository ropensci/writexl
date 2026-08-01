# Every object writexl hands back prints something.  A print method is the one
# piece of a package nothing else exercises -- no other code path calls it, so
# a typo in one sits there until a user types the object's name at the prompt.
#
# Each object is built twice where it matters: once bare and once with the
# optional parts filled in, since most of these methods only describe what is
# set.

expect_prints <- function(x, ...) {
  out <- utils::capture.output(print(x))
  expect_true(length(out) >= 1L)
  for (pattern in c(...)) expect_true(any(grepl(pattern, out, fixed = TRUE)),
                                      info = pattern)
  # print() returns its argument invisibly, as R expects of a print method
  vis <- NULL
  utils::capture.output(vis <- withVisible(print(x))$visible)
  expect_identical(vis, FALSE)
  invisible(out)
}

df <- data.frame(quarter = c("Q1", "Q2"), revenue = c(10, 25),
                 stringsAsFactors = FALSE)

test_that("the cell content objects print", {
  expect_prints(xl_cell_general(value = 1:2), "xl_cell_general")
  expect_prints(xl_comment("note", author = "QA"), "value:", "note")
  expect_prints(xl_merge("A1:B1", "Total"), "xl_merge", "Total")
  expect_prints(xl_merge(list(rows = 1, cols = 1:2)), "<spec>")
  expect_prints(xl_rich_run("bold", xl_font(bold = TRUE)), "formatted")
  expect_prints(xl_rich_run("plain"), "xl_rich_run")
  expect_prints(xl_rich_string("a ", xl_rich_run("b", xl_font(bold = TRUE))),
                "2 runs", "formatted")
})

test_that("the format objects print", {
  expect_prints(xl_format(), "xl_format")
  expect_prints(xl_font(bold = TRUE) + xl_fill(background = "yellow"), "font")
  expect_prints(xl_col_spec("revenue", width = 12), "<xl_col_spec>", "width=12")
  expect_prints(xl_row_spec(1, height = 20), "<xl_row_spec>", "height=20")
})

test_that("the worksheet and workbook objects print", {
  expect_prints(xl_sheet(df), "2 rows x 2 cols")
  expect_prints(xl_sheet(df, cols = xl_col_spec("revenue", width = 12),
                         rows = xl_row_spec(1, height = 20)),
                "column specs:", "row specs:")
  expect_prints(xl_workbook(list(D = df)), "xl_workbook")
  expect_prints(xl_properties(), "xl_properties")
  expect_prints(xl_properties(title = "T", author = "A", company = "C",
                              subject = "S", read_only = TRUE),
                "title:", "read_only: TRUE")
  expect_prints(xl_page_setup(orientation = "landscape"), "xl_page_setup")
  expect_prints(xl_sheet_view(), "xl_sheet_view")
  expect_prints(xl_sheet_view(active = TRUE, hide_zero = TRUE),
                "active", "hide_zero")
  expect_prints(xl_outline(symbols_below = FALSE), "xl_outline")
})

test_that("the worksheet feature objects print", {
  expect_prints(xl_validation("A1:A5", type = "integer", criteria = ">",
                              value = 0), "xl_validation")
  expect_prints(xl_cond_cell("A1:A5", type = "cell", criteria = ">",
                             value = 5, format = xl_font(bold = TRUE)),
                "xl_conditional")
  expect_prints(xl_cond_scale("A1:A5"), "xl_conditional")
  expect_prints(xl_filter(col = "quarter", criteria = "==", value = "Q1"),
                "xl_filter")
  expect_prints(xl_table(name = "Sales"), "xl_table")
  expect_prints(xl_table_column("revenue", total = "sum"), "revenue", "total=")
  expect_prints(xl_table_column("revenue", formula = "=1"), "formula")
  expect_prints(xl_image(png_file(), "A1"), "xl_image")
})

test_that("the chart objects print", {
  s <- xl_chart_series(values = list(cols = "revenue"))
  expect_prints(s, "xl_chart_series")
  expect_prints(xl_chart("column", s), "xl_chart", "column")
  expect_prints(xl_chart_axis(title = "Quarter", min = 0), "xl_chart_axis")
  expect_prints(xl_chart_axis(), "xl_chart_axis")
  expect_prints(xl_chart_marker(type = "circle", size = 8), "xl_chart_marker")
  expect_prints(xl_chart_labels(show_value = TRUE), "xl_chart_labels")
  expect_prints(xl_chart_label(value = "peak"), "xl_chart_label")
  expect_prints(xl_chart_trendline(type = "linear"), "xl_chart_trendline")
  expect_prints(xl_chart_error_bars(type = "percentage", value = 5),
                "xl_chart_error_bars")
  expect_prints(xl_chart_legend(position = "bottom"), "xl_chart_legend")
  expect_prints(xl_chart_table(show_keys = TRUE), "xl_chart_table")
  expect_prints(xl_chartsheet(xl_chart("pie", s)), "xl_chartsheet", "pie")
  expect_prints(xl_chartsheet(xl_chart("pie", s), zoom = 150), "set:", "zoom")
})

# Chartsheets: a sheet that is one chart and no cells.  A chartsheet supports a
# small subset of what a worksheet does, and the tests below are mostly about
# refusing the rest rather than letting Excel drop it.

cs_sales <- data.frame(quarter = c("Q1", "Q2", "Q3", "Q4"),
                       revenue = c(10, 25, 18, 32),
                       stringsAsFactors = FALSE)

cs_series <- function()
  xl_chart_series(values = list(sheet = "Data", cols = "revenue"),
                  categories = list(sheet = "Data", cols = "quarter"))

# Write a workbook holding one chartsheet and return its parts.
cs_parts <- function(..., name = "Overview") {
  t <- tempfile(fileext = ".xlsx")
  wb <- list(Data = cs_sales)
  wb[[name]] <- xl_chartsheet(...)
  write_xlsx(wb, t)
  d <- tempfile(); dir.create(d); utils::unzip(t, exdir = d)
  f <- list.files(file.path(d, "xl/chartsheets"), pattern = "^sheet",
                  full.names = TRUE)
  list(files = list.files(d, recursive = TRUE),
       sheet = if (length(f)) paste(readLines(f[1L], warn = FALSE),
                                    collapse = "") else "",
       workbook = paste(readLines(file.path(d, "xl/workbook.xml"),
                                  warn = FALSE), collapse = ""))
}

# ── What reaches the file ─────────────────────────────────────────────────────

test_that("a chartsheet is its own part, with no worksheet beside it", {
  r <- cs_parts(xl_chart("column", cs_series(), title = "Revenue"))
  expect_true(any(grepl("^xl/chartsheets/sheet1.xml$", r$files)))
  expect_true(any(grepl("^xl/charts/chart1.xml$", r$files)))
  expect_true(any(grepl("^xl/drawings/drawing1.xml$", r$files)))
  # the data sheet is the only worksheet
  expect_length(grep("^xl/worksheets/sheet", r$files), 1L)
  expect_match(r$sheet, "<drawing r:id=", fixed = TRUE)
  expect_match(r$workbook, "Overview", fixed = TRUE)
})

test_that("the chart on a chartsheet carries everything a chart can", {
  r <- cs_parts(xl_chart("line",
                         xl_chart_series(values = list(sheet = "Data",
                                                       cols = "revenue"),
                                         categories = list(sheet = "Data",
                                                           cols = "quarter"),
                                         marker = xl_chart_marker("circle", 7)),
                         title = "T",
                         legend = xl_chart_legend("bottom"),
                         y_axis = xl_chart_axis(title = "Revenue", min = 0)))
  # the chart itself is covered by the chart tests; what matters here is that
  # a fully-featured chart still reaches a chartsheet
  expect_true(any(grepl("^xl/charts/chart1.xml$", r$files)))
  expect_true(any(grepl("^xl/chartsheets/sheet1.xml$", r$files)))
})

test_that("the settings a chartsheet does support reach the file", {
  r <- cs_parts(xl_chart("pie", cs_series()),
                tab_color = "red", zoom = 150,
                view = xl_sheet_view(active = TRUE),
                page = xl_page_setup(orientation = "landscape", paper = "A4",
                                     header = "&LTop"),
                protect = "secret")
  expect_match(r$sheet, "FF0000", fixed = TRUE)
  expect_match(r$sheet, 'zoomScale="150"', fixed = TRUE)
  expect_match(r$sheet, 'orientation="landscape"', fixed = TRUE)
  expect_match(r$sheet, 'paperSize="9"', fixed = TRUE)
  expect_match(r$sheet, "&amp;LTop", fixed = TRUE)
  expect_match(r$sheet, "sheetProtection", fixed = TRUE)
})

test_that("a chartsheet takes part in the workbook-wide tab rules", {
  # the visibility checks span every sheet, and a chartsheet is a sheet
  expect_error(
    write_xlsx(list(Data = cs_sales,
                    Hidden = xl_chartsheet(xl_chart("pie", cs_series()),
                                           view = xl_sheet_view(visible = FALSE,
                                                                active = TRUE))),
               tempfile(fileext = ".xlsx")),
    "hidden")
  # and one can be the active tab
  expect_silent(write_xlsx(list(
    Data = cs_sales,
    Front = xl_chartsheet(xl_chart("pie", cs_series()),
                          view = xl_sheet_view(active = TRUE))),
    tempfile(fileext = ".xlsx")))
})

# ── What a chartsheet refuses ─────────────────────────────────────────────────

test_that("a chartsheet holds exactly one chart", {
  expect_error(xl_chartsheet(list(xl_chart("column",
                                           xl_chart_series(values = "A1:A5")),
                                  xl_chart("pie",
                                           xl_chart_series(values = "A1:A5")))),
               "holds exactly one chart, not 2")
  expect_error(xl_chartsheet(), "must be the xl_chart")
})

test_that("every range must name its sheet, since a chartsheet has none", {
  bad <- function(...) write_xlsx(list(Data = cs_sales,
                                       Bad = xl_chartsheet(xl_chart(...))),
                                  tempfile(fileext = ".xlsx"))
  expect_error(bad("column", xl_chart_series(values = list(cols = "revenue"))),
               "does not name a sheet")
  expect_error(bad("column", xl_chart_series(values = list(cols = "revenue"))),
               'chartsheet "Bad" has no cells of its own')
  # an unqualified A1 range is refused for the same reason
  expect_error(bad("column", xl_chart_series(values = "B2:B5")),
               "does not name a sheet")
  # the categories, the series name and the title are checked too
  expect_error(bad("column",
                   xl_chart_series(values = list(sheet = "Data",
                                                 cols = "revenue"),
                                   categories = list(cols = "quarter"))),
               "series\\[\\[1\\]\\]\\$categories` does not name a sheet")
  expect_error(bad("column", cs_series(), title = list(header = "revenue")),
               "`title` does not name a sheet")
  # a literal title or name carries no range, so neither is checked
  expect_silent(bad("column", xl_chart_series(values = list(sheet = "Data",
                                                            cols = "revenue"),
                                              name = "Revenue"),
                    title = "Revenue"))
})

test_that("the page options a chartsheet has no setter for are refused", {
  ch <- xl_chart("column", cs_series())
  expect_error(xl_chartsheet(ch, page = xl_page_setup(scale = 90)),
               "`page` sets `scale`, which a chartsheet has no setter for",
               fixed = TRUE)
  expect_error(xl_chartsheet(ch, page = xl_page_setup(print_area = "A1:B2")),
               "A chartsheet takes: orientation, paper, margins", fixed = TRUE)
  # and the five it does have are accepted
  expect_s3_class(xl_chartsheet(ch, page = xl_page_setup(
                    orientation = "landscape", paper = "A4",
                    margins = c(1, 1, 1, 1), header = "&Lx", footer = "&Ry")),
                  "xl_chartsheet")
})

test_that("the view options a chartsheet has no setter for are refused", {
  ch <- xl_chart("column", cs_series())
  for (opt in list(list(hide_zero = TRUE), list(right_to_left = TRUE),
                   list(selection = "A1"), list(top_left = "A1")))
    expect_error(xl_chartsheet(ch, view = do.call(xl_sheet_view, opt)),
                 "which a chartsheet has no setter for", label = names(opt))
  for (opt in list(list(active = TRUE), list(selected = TRUE),
                   list(visible = FALSE), list(first_tab = TRUE)))
    expect_true(inherits(xl_chartsheet(ch, view = do.call(xl_sheet_view, opt)),
                         "xl_chartsheet"), info = names(opt))
})

test_that("chartsheet arguments are validated", {
  ch <- xl_chart("column", cs_series())
  expect_error(xl_chartsheet(ch, view = list(active = TRUE)),
               "must be an xl_sheet_view")
  expect_error(xl_chartsheet(ch, page = list(paper = "A4")),
               "must be an xl_page_setup")
  expect_error(xl_chartsheet(ch, zoom = 5), "must be between 10 and 400")
  expect_error(xl_chartsheet(ch, tab_color = "not a colour"), "tab_color")
  expect_error(xl_chartsheet(ch, protect = list(nope = TRUE)), "protect")
})

# ── Alongside the rest of the workbook ────────────────────────────────────────

test_that("a chartsheet counts as a drawing in the ordering check", {
  # a chartsheet's chart is a drawing, so it is a victim of the same
  # libxlsxwriter numbering bug a floating image is
  png <- tempfile(fileext = ".png")
  writeBin(as.raw(c(
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
    0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00,
    0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
    0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82)), png)
  expect_error(
    write_xlsx(list(
      Data = xl_sheet(cs_sales,
                      page = xl_page_setup(header = "&L&G",
                                           header_image = list(left = png))),
      Chart = xl_chartsheet(xl_chart("column", cs_series()))),
      tempfile(fileext = ".xlsx")),
    "earlier in the workbook")
})

test_that("a chartsheet contributes no cells to the streaming estimate", {
  # it has no data frame, so it cannot tip the constant-memory decision
  res <- .resolve_constant_memory(
    list(Data = cs_sales, Overview = .chartsheet_placeholder()),
    xl_properties(), NULL, NA, 1)
  expect_equal(res$on, 1L)
})

test_that("a chartsheet may be given to xl_workbook() like any other sheet", {
  wb <- xl_workbook(list(Data = cs_sales,
                         Overview = xl_chartsheet(xl_chart("pie",
                                                           cs_series()))),
                    properties = xl_properties(title = "T"))
  expect_s3_class(wb$sheets$Overview, "xl_chartsheet")
  expect_silent(write_xlsx(wb, tempfile(fileext = ".xlsx")))
})

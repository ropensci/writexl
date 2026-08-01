# The chart itself: the legend, the data table, the plot and chart areas, and
# the options that belong to one family of chart.

ex_sales <- data.frame(quarter = c("Q1", "Q2", "Q3", "Q4"),
                       revenue = c(10, 25, 18, 32),
                       cost = c(8, 12, 25, 20),
                       stringsAsFactors = FALSE)

ex_two <- function() list(
  xl_chart_series(values = list(cols = "revenue"),
                  categories = list(cols = "quarter")),
  xl_chart_series(values = list(cols = "cost"),
                  categories = list(cols = "quarter")))

ex_xml <- function(..., type = "column", series = ex_two()) {
  t <- tempfile(fileext = ".xlsx")
  write_xlsx(list(Data = xl_sheet(ex_sales,
                                  chart = xl_chart(type, series, ...))), t)
  d <- tempfile(); dir.create(d); utils::unzip(t, exdir = d)
  paste(readLines(list.files(file.path(d, "xl/charts"), pattern = "^chart",
                             full.names = TRUE)[1L], warn = FALSE),
        collapse = "")
}

# ── The legend ────────────────────────────────────────────────────────────────

test_that("the legend positions are libxlsxwriter's, with none at zero", {
  # NONE is 0 and the rest run 1..8.  Anything above the range
  # chart_legend_set_position() accepts is refused there with a warning, and
  # the legend stays where it is.
  expect_equal(.LXW_CHART_LEGEND[["none"]], 0L)
  expect_equal(unname(.LXW_CHART_LEGEND), 0:8)
  expect_equal(names(.LXW_CHART_LEGEND)[[1L]], "none")
})

test_that("a legend can be moved, styled, placed and removed", {
  expect_match(ex_xml(legend = xl_chart_legend("bottom")),
               '<c:legendPos val="b"/>', fixed = TRUE)
  expect_match(ex_xml(legend = xl_chart_legend("top_right")),
               '<c:legendPos val="tr"/>', fixed = TRUE)
  expect_match(ex_xml(legend = xl_chart_legend(format = xl_font(size = 8,
                                                                italic = TRUE))),
               'i="1"', fixed = TRUE)
  expect_match(ex_xml(legend = xl_chart_legend(layout = c(0.1, 0.1, 0.2, 0.2))),
               "<c:manualLayout>", fixed = TRUE)
  # and "none" removes it outright, with no warning from the C layer
  none <- expect_silent(ex_xml(legend = xl_chart_legend("none")))
  expect_false(grepl("<c:legend>", none, fixed = TRUE))
})

test_that("a series can be left out of the legend while still being plotted", {
  x <- ex_xml(legend = xl_chart_legend(delete_series = 2))
  expect_match(x, "<c:legendEntry>", fixed = TRUE)
  expect_match(x, '<c:idx val="1"/>', fixed = TRUE)   # 0-based on the way out
  # both series are still drawn
  expect_length(regmatches(x, gregexpr("<c:ser>", x))[[1L]], 2L)
})

test_that("legend arguments are validated", {
  expect_error(xl_chart_legend("middle"), "`position` must be one of")
  expect_error(xl_chart_legend(delete_series = 0), "counting from 1")
  expect_error(xl_chart_legend(delete_series = "first"), "counting from 1")
  expect_error(xl_chart_legend(format = xl_fill(background = "red")),
               "takes no fill")
  expect_error(xl_chart_legend(layout = c(0.5, 0.5, 0.5)),
               "c\\(x, y\\) or c\\(x, y, width, height\\)")
  expect_error(xl_chart("column", ex_two(), legend = "bottom"),
               "must be an xl_chart_legend")
})

# ── The data table ────────────────────────────────────────────────────────────

test_that("a data table reaches the file with its grid options", {
  expect_match(ex_xml(data_table = xl_chart_table()), "<c:dTable>",
               fixed = TRUE)
  x <- ex_xml(data_table = xl_chart_table(show_keys = TRUE, vertical_border = FALSE))
  expect_match(x, '<c:showKeys val="1"/>', fixed = TRUE)
  expect_match(x, '<c:showHorzBorder val="1"/>', fixed = TRUE)
  # a border switched off is written as an absence, not as val="0"
  expect_false(grepl("showVertBorder", x, fixed = TRUE))
})

test_that("a data table's font is the only thing it can be styled with", {
  expect_match(ex_xml(data_table = xl_chart_table(format = xl_font(size = 7))),
               'sz="700"', fixed = TRUE)
  expect_error(xl_chart_table(format = xl_border(all = "thin")),
               "takes no line")
  expect_error(xl_chart("column", ex_two(), data_table = TRUE),
               "must be an xl_chart_table")
})

# ── The plot and chart areas ──────────────────────────────────────────────────

test_that("the plot area and the chart area take a line, fill and pattern", {
  x <- ex_xml(plot_area_format = xl_fill(background = "#F2F2F2") +
                                 xl_border(all = "thin", color = "gray"),
              chart_area_format = xl_border(all = "dashed", color = "navy"))
  expect_match(x, "F2F2F2", fixed = TRUE)
  expect_match(x, '<a:prstDash val="dash"/>', fixed = TRUE)
  expect_match(ex_xml(chart_area_format = xl_fill(pattern = "light-up",
                                                  foreground = "#2E75B6",
                                                  background = "white")),
               "<a:pattFill", fixed = TRUE)
  # an area is a shape, so no font
  expect_error(xl_chart("column", ex_two(),
                        plot_area_format = xl_font(bold = TRUE)),
               "a shape and has no text")
})

test_that("the plot area and the title can be placed by hand", {
  expect_match(ex_xml(plot_area_layout = c(0.1, 0.1, 0.8, 0.7)),
               "<c:manualLayout>", fixed = TRUE)
  x <- ex_xml(title = "T", title_layout = c(0.3, 0.02),
              title_overlay = TRUE)
  expect_match(x, "<c:manualLayout>", fixed = TRUE)
  expect_match(x, '<c:overlay val="1"/>', fixed = TRUE)
})

test_that("placing a title needs a title to place", {
  expect_error(xl_chart("column", ex_two(), title_layout = c(0.3, 0.02)),
               "give a `title` too", fixed = TRUE)
  expect_error(xl_chart("column", ex_two(), title = FALSE,
                        title_overlay = TRUE),
               "`title = FALSE` removes it", fixed = TRUE)
  expect_error(xl_chart("column", ex_two(), title = "T",
                        title_layout = c(0, 0.5)),
               "each above 0 and at most 1")
})

# ── The options that belong to one family ─────────────────────────────────────

test_that("each family-only option is refused on the other families", {
  cases <- list(
    hole_size = list(val = 40, ok = "doughnut", no = "pie"),
    rotation = list(val = 90, ok = "pie", no = "column"),
    drop_lines = list(val = TRUE, ok = "line", no = "column"),
    high_low_lines = list(val = TRUE, ok = "line", no = "column"),
    up_down_bars = list(val = TRUE, ok = "line", no = "column"),
    series_gap = list(val = 50, ok = "column", no = "line"),
    series_overlap = list(val = -20, ok = "column", no = "line"))
  for (nm in names(cases)) {
    k <- cases[[nm]]
    args <- list(k$no, xl_chart_series(values = "A1:A5"))
    args[[nm]] <- k$val
    expect_error(do.call(xl_chart, args),
                 sprintf('`%s` does not apply to a "%s" chart', nm, k$no),
                 fixed = TRUE, label = nm)
    args <- list(k$ok, xl_chart_series(values = "A1:A5"))
    args[[nm]] <- k$val
    expect_true(inherits(do.call(xl_chart, args), "xl_chart"), info = nm)
  }
  # a hole belongs to a doughnut and not to a pie, which shares its family
  expect_error(xl_chart("pie", xl_chart_series(values = "A1"), hole_size = 40),
               "It applies to: doughnut", fixed = TRUE)
})

test_that("the family-only options reach the file", {
  d <- ex_xml(type = "doughnut", series = ex_two()[[1L]],
              hole_size = 40, rotation = 90)
  expect_match(d, '<c:holeSize val="40"/>', fixed = TRUE)
  expect_match(d, '<c:firstSliceAng val="90"/>', fixed = TRUE)
  b <- ex_xml(series_gap = 50, series_overlap = -20)
  expect_match(b, '<c:gapWidth val="50"/>', fixed = TRUE)
  expect_match(b, '<c:overlap val="-20"/>', fixed = TRUE)
})

test_that("drop, high-low and up-down lines take a format or a plain TRUE", {
  x <- ex_xml(type = "line",
              drop_lines = xl_border(all = "dotted", color = "gray"),
              high_low_lines = TRUE,
              up_down_bars = list(up = xl_fill(background = "green"),
                                  down = xl_fill(background = "red")))
  expect_match(x, "<c:dropLines>", fixed = TRUE)
  expect_match(x, "<c:hiLowLines", fixed = TRUE)
  expect_match(x, "<c:upDownBars>", fixed = TRUE)
  expect_match(x, "00FF00", fixed = TRUE)      # the up bar
  expect_match(x, "FF0000", fixed = TRUE)      # the down bar
  expect_match(ex_xml(type = "line", up_down_bars = TRUE), "<c:upDownBars>",
               fixed = TRUE)
})

test_that("the switch arguments say what they accept", {
  expect_error(xl_chart("line", ex_two(), drop_lines = "yes"),
               "must be TRUE, or an xl_format")
  expect_error(xl_chart("line", ex_two(), up_down_bars = list(sideways = 1)),
               "must be TRUE, or list\\(up = , down = \\)")
  expect_error(xl_chart("line", ex_two(),
                        drop_lines = xl_fill(background = "red")),
               "takes no fill")
  # FALSE is the same as leaving them out
  expect_null(unclass(xl_chart("line", ex_two(),
                               drop_lines = FALSE))$drop_lines)
})

test_that("the numeric options are held to Excel's ranges", {
  expect_error(xl_chart("doughnut", ex_two()[[1L]], hole_size = 5),
               "must be between 10 and 90")
  expect_error(xl_chart("pie", ex_two()[[1L]], rotation = 400),
               "must be between 0 and 360")
  expect_error(xl_chart("column", ex_two(), series_gap = 600),
               "must be between 0 and 500")
  expect_error(xl_chart("column", ex_two(), series_overlap = -200),
               "must be between -100 and 100")
})

# ── Blanks and hidden data ────────────────────────────────────────────────────

test_that("what an empty cell does to the plot can be chosen", {
  expect_equal(unname(.LXW_CHART_BLANKS), 0:2)
  expect_match(ex_xml(show_blanks = "zero"), '<c:dispBlanksAs val="zero"/>',
               fixed = TRUE)
  expect_match(ex_xml(show_blanks = "connected"),
               '<c:dispBlanksAs val="span"/>', fixed = TRUE)
  expect_error(xl_chart("column", ex_two(), show_blanks = "skip"),
               "`show_blanks` must be one of")
})

test_that("hidden data is plotted by leaving plotVisOnly out", {
  # libxlsxwriter expresses this as an absence rather than as val="0", so the
  # test pins the absence
  expect_match(ex_xml(), "<c:plotVisOnly", fixed = TRUE)
  expect_false(grepl("plotVisOnly", ex_xml(show_hidden_data = TRUE),
                     fixed = TRUE))
})

# ── Coverage of the C API ─────────────────────────────────────────────────────

test_that("every chart-level chart_*() function libxlsxwriter offers is called", {
  hdr <- NULL
  for (p in c("../../src/libxlsxwriter/include/xlsxwriter/chart.h",
              "../../../src/libxlsxwriter/include/xlsxwriter/chart.h"))
    if (file.exists(p)) hdr <- p
  skip_if(is.null(hdr), "libxlsxwriter sources not alongside the tests")
  src <- NULL
  for (p in c("../../src/write_xlsx.c", "../../../src/write_xlsx.c"))
    if (file.exists(p)) src <- p
  skip_if(is.null(src), "write_xlsx.c not alongside the tests")

  decl <- grep("^(void|lxw_error|uint8_t) chart_", readLines(hdr), value = TRUE)
  fns <- unique(regmatches(decl, regexpr("chart_[a-z_0-9]+", decl)))
  # the axes and the series have gates of their own
  fns <- fns[!grepl("^chart_(series|axis)_", fns)]
  # the title and the built-in style are in here too
  expect_equal(length(fns), 31L)

  called <- paste(readLines(src), collapse = "\n")
  for (f in fns)
    expect_true(grepl(paste0(f, "("), called, fixed = TRUE), info = f)
})

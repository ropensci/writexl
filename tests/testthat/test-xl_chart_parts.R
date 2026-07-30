# Markers, data labels, trendlines, error bars and individual points.  As with
# the axes, the point of most of this is that libxlsxwriter documents a
# restriction and then lets Excel drop the setting in silence.

part_sales <- data.frame(quarter = c("Q1", "Q2", "Q3", "Q4"),
                         revenue = c(10, 20, 30, 40),
                         stringsAsFactors = FALSE)

part_xml <- function(..., type = "line", df = part_sales) {
  t <- tempfile(fileext = ".xlsx")
  se <- xl_chart_series(values = list(cols = "revenue"),
                        categories = list(cols = "quarter"), ...)
  write_xlsx(list(Data = xl_sheet(df, chart = xl_chart(type, se))), t)
  d <- tempfile(); dir.create(d); utils::unzip(t, exdir = d)
  paste(readLines(list.files(file.path(d, "xl/charts"), pattern = "^chart",
                             full.names = TRUE)[1L], warn = FALSE),
        collapse = "")
}

# ── Markers ───────────────────────────────────────────────────────────────────

test_that("the marker types are libxlsxwriter's, in its order", {
  expect_equal(unname(.LXW_MARKER_TYPE), 0:10)
  expect_equal(names(.LXW_MARKER_TYPE)[1:2], c("automatic", "none"))
})

test_that("a marker reaches the file with its size and styling", {
  x <- part_xml(marker = xl_chart_marker(type = "diamond", size = 9,
                                         format = xl_fill(background = "red")))
  expect_match(x, '<c:symbol val="diamond"/>', fixed = TRUE)
  expect_match(x, '<c:size val="9"/>', fixed = TRUE)
  expect_match(x, "FF0000", fixed = TRUE)
  expect_match(part_xml(marker = xl_chart_marker(type = "none")),
               '<c:symbol val="none"/>', fixed = TRUE)
})

test_that("an automatic marker cannot also be sized or styled", {
  # documented in chart.h: "If automatic is on then other marker properties
  # such as size, line or fill cannot be set"
  expect_error(xl_chart_marker(type = "automatic", size = 8),
               "cannot then be given a size or a format")
  expect_error(xl_chart_marker(type = "automatic",
                               format = xl_fill(background = "red")),
               "cannot then be given a size or a format")
  expect_s3_class(xl_chart_marker(type = "automatic"), "xl_chart_marker")
  # and the restriction is on "automatic" alone
  expect_s3_class(xl_chart_marker(type = "circle", size = 8),
                  "xl_chart_marker")
})

test_that("marker arguments are validated", {
  expect_error(xl_chart_marker(type = "blob"), "`type` must be one of")
  expect_error(xl_chart_marker(size = 1), "must be between 2 and 72")
  expect_error(xl_chart_marker(format = xl_font(bold = TRUE)),
               "a shape and has no text")
})

# ── Data labels ───────────────────────────────────────────────────────────────

test_that("labels default to showing the value, as Excel does", {
  x <- part_xml(labels = xl_chart_labels())
  expect_match(x, "<c:dLbls>", fixed = TRUE)
  expect_match(x, '<c:showVal val="1"/>', fixed = TRUE)
})

test_that("every label option reaches the file", {
  x <- part_xml(labels = xl_chart_labels(
    show_value = TRUE, show_category = TRUE, show_name = FALSE,
    show_legend_key = TRUE, separator = "newline", position = "above",
    num_format = "0.0", leader_lines = TRUE,
    format = xl_font(size = 9, bold = TRUE) +
             xl_fill(background = "#FFF2CC") +
             xl_border(all = "thin", color = "gray")))
  for (w in c('<c:showVal val="1"/>', '<c:showCatName val="1"/>',
              '<c:showLegendKey val="1"/>',
              "<c:separator>", '<c:dLblPos val="t"/>', 'formatCode="0.0"',
              '<c:showLeaderLines val="1"/>', 'sz="900"', 'b="1"',
              "FFF2CC"))
    expect_true(grepl(w, x, fixed = TRUE), info = w)
  # what is switched off is left out rather than written as 0, which is how
  # Excel writes it too -- these all default to off
  expect_false(grepl("showSerName", x, fixed = TRUE))
})

test_that("a label holds exactly the parts named, and the value otherwise", {
  # Excel puts the value in every label unless told otherwise, so asking for
  # the percentage alone used to produce "10, 11.8%" -- found in Excel, and
  # the reason naming any part now names them all
  only_pc <- part_xml(labels = xl_chart_labels(show_percentage = TRUE),
                      type = "pie")
  expect_match(only_pc, '<c:showPercent val="1"/>', fixed = TRUE)
  expect_false(grepl("showVal", only_pc, fixed = TRUE))

  # naming nothing leaves Excel's default
  expect_match(part_xml(labels = xl_chart_labels(), type = "pie"),
               '<c:showVal val="1"/>', fixed = TRUE)

  # and both together, when both are asked for
  both <- part_xml(labels = xl_chart_labels(show_value = TRUE,
                                            show_percentage = TRUE),
                   type = "pie")
  expect_match(both, '<c:showVal val="1"/>', fixed = TRUE)
  expect_match(both, '<c:showPercent val="1"/>', fixed = TRUE)

  # naming the category alone drops the value too
  cat_only <- part_xml(labels = xl_chart_labels(show_category = TRUE),
                       type = "pie")
  expect_match(cat_only, '<c:showCatName val="1"/>', fixed = TRUE)
  expect_false(grepl("showVal", cat_only, fixed = TRUE))
})

test_that("a label position is refused where the chart type has no such place", {
  # the table lives in chart.h and Excel drops what does not apply
  expect_error(xl_chart("column", xl_chart_series(values = "A1:A5",
                          labels = xl_chart_labels(position = "above"))),
               'does not apply to a "column" chart')
  expect_error(xl_chart("column", xl_chart_series(values = "A1:A5",
                          labels = xl_chart_labels(position = "above"))),
               "the positions are: default, center, inside_base", fixed = TRUE)
  expect_error(xl_chart("line", xl_chart_series(values = "A1:A5",
                          labels = xl_chart_labels(position = "best_fit"))),
               "applies to: doughnut, pie", fixed = TRUE)
  # and each is accepted where the table says it belongs
  expect_s3_class(xl_chart("line", xl_chart_series(values = "A1:A5",
                            labels = xl_chart_labels(position = "above"))),
                  "xl_chart")
  expect_s3_class(xl_chart("column", xl_chart_series(values = "A1:A5",
                            labels = xl_chart_labels(position = "inside_end"))),
                  "xl_chart")
  expect_s3_class(xl_chart("pie", xl_chart_series(values = "A1:A5",
                            labels = xl_chart_labels(position = "best_fit"))),
                  "xl_chart")
  # "center" is the one position every type allows
  for (ty in names(.LXW_CHART_TYPE)) {
    se <- if (identical(.CHART_FAMILY(ty), "scatter"))
      xl_chart_series(values = "B1:B5", categories = "A1:A5",
                      labels = xl_chart_labels(position = "center"))
    else xl_chart_series(values = "A1:A5",
                         labels = xl_chart_labels(position = "center"))
    expect_true(inherits(xl_chart(ty, se), "xl_chart"), info = ty)
  }
})

test_that("custom labels give one point its own text, styling or nothing", {
  x <- part_xml(type = "pie", labels = xl_chart_labels(
    custom = list(xl_chart_label(value = "biggest"), NULL,
                  xl_chart_label(hide = TRUE),
                  xl_chart_label(format = xl_font(bold = TRUE)))))
  expect_match(x, "biggest", fixed = TRUE)
  expect_match(x, '<c:delete val="1"/>', fixed = TRUE)
})

test_that("a hidden custom label carries nothing else", {
  expect_error(xl_chart_label(hide = TRUE, value = "x"),
               "cannot also carry a value or a format")
  expect_error(xl_chart_label(hide = TRUE, format = xl_font(bold = TRUE)),
               "cannot also carry a value or a format")
  expect_error(xl_chart_labels(custom = list("text")),
               "must be an xl_chart_label\\(\\) or NULL")
  # one label need not be wrapped in a list
  expect_length(unclass(xl_chart_labels(custom = xl_chart_label(value = "a")))$custom,
                1L)
})

test_that("label arguments are validated", {
  expect_error(xl_chart_labels(position = "middle"), "`position` must be one of")
  expect_error(xl_chart_labels(separator = "tab"), "`separator` must be one of")
  expect_error(xl_chart_labels(num_format = 42), "must be an Excel format string")
  expect_error(xl_chart_labels(num_format = xl_font(bold = TRUE)),
               "no number format in it")
  expect_equal(unclass(xl_chart_labels(num_format = xl_num_format("0%")))$num_format,
               "0%")
})

# ── Trendlines ────────────────────────────────────────────────────────────────

test_that("a trendline reaches the file with all of its options", {
  x <- part_xml(trendline = xl_chart_trendline(
    "linear", forward = 1, backward = 0.5, equation = TRUE, r_squared = TRUE,
    intercept = 0, name = "fit",
    format = xl_border(all = "dashed", color = "green")))
  for (w in c("<c:trendline>", '<c:trendlineType val="linear"/>',
              '<c:forward val="1"/>', '<c:backward val="0.5"/>',
              '<c:dispEq val="1"/>', '<c:dispRSqr val="1"/>',
              '<c:intercept val="0"/>', "fit", '<a:prstDash val="dash"/>'))
    expect_true(grepl(w, x, fixed = TRUE), info = w)
})

test_that("a polynomial needs an order and a moving average a period", {
  expect_error(xl_chart_trendline("poly"), "needs an `order`")
  expect_error(xl_chart_trendline("average"), "needs a `period`")
  expect_error(xl_chart_trendline("linear", order = 2),
               "not of a \"linear\" one")
  expect_error(xl_chart_trendline("linear", period = 2),
               "a \"linear\" fit has none")
  x <- part_xml(trendline = xl_chart_trendline("poly", order = 3))
  expect_match(x, '<c:order val="3"/>', fixed = TRUE)
  y <- part_xml(trendline = xl_chart_trendline("average", period = 2))
  expect_match(y, '<c:trendlineType val="movingAvg"/>', fixed = TRUE)
  expect_match(y, '<c:period val="2"/>', fixed = TRUE)
})

test_that("a moving average has no forecast, equation or R-squared", {
  # chart.h: "This feature isn't available for Moving Average in Excel"
  for (a in list(list(forward = 1), list(backward = 1), list(equation = TRUE),
                 list(r_squared = TRUE)))
    expect_error(do.call(xl_chart_trendline,
                         c(list("average", period = 2), a)),
                 "does not apply to a moving average", label = names(a))
  # the same options are fine on a fitted type
  expect_s3_class(xl_chart_trendline("exp", forward = 1, equation = TRUE),
                  "xl_chart_trendline")
})

test_that("an intercept applies to three of the six fits", {
  # chart.h: "only available in Excel for Exponential, Linear and Polynomial"
  for (ty in c("exp", "linear"))
    expect_s3_class(xl_chart_trendline(ty, intercept = 0),
                    "xl_chart_trendline")
  expect_s3_class(xl_chart_trendline("poly", order = 2, intercept = 0),
                  "xl_chart_trendline")
  for (ty in c("log", "power"))
    expect_error(xl_chart_trendline(ty, intercept = 0),
                 "exponential, linear or polynomial", label = ty)
  expect_error(xl_chart_trendline("average", period = 2, intercept = 0),
               "exponential, linear or polynomial")
})

test_that("trendline arguments are validated", {
  expect_error(xl_chart_trendline(), "must name the fit")
  expect_error(xl_chart_trendline("cubic"), "`type` must be one of")
  expect_error(xl_chart_trendline("poly", order = 1), "must be between 2 and NA")
  expect_error(xl_chart_trendline("linear", format = xl_fill(background = "red")),
               "takes no fill")
})

# ── Error bars ────────────────────────────────────────────────────────────────

test_that("error bars reach both axes of a series", {
  t <- tempfile(fileext = ".xlsx")
  write_xlsx(list(Data = xl_sheet(part_sales, chart = xl_chart("scatter",
    xl_chart_series(values = list(cols = "revenue"),
                    categories = list(cols = "revenue"),
                    x_error_bars = xl_chart_error_bars("fixed", 2,
                                                       direction = "plus"),
                    y_error_bars = xl_chart_error_bars("std_dev", 1))))), t)
  d <- tempfile(); dir.create(d); utils::unzip(t, exdir = d)
  x <- paste(readLines(list.files(file.path(d, "xl/charts"), pattern = "^chart",
                                  full.names = TRUE)[1L], warn = FALSE),
             collapse = "")
  expect_length(regmatches(x, gregexpr("<c:errBars>", x))[[1L]], 2L)
  expect_match(x, '<c:errBarType val="plus"/>', fixed = TRUE)
  expect_match(x, '<c:errValType val="stdDev"/>', fixed = TRUE)
})

test_that("an error bar's endcap and line reach the file", {
  x <- part_xml(y_error_bars = xl_chart_error_bars(
    "percentage", 10, endcap = FALSE,
    format = xl_border(all = "dotted", color = "purple")))
  expect_match(x, '<c:noEndCap val="1"/>', fixed = TRUE)
  expect_match(x, '<c:errValType val="percentage"/>', fixed = TRUE)
  expect_match(x, '<c:val val="10"/>', fixed = TRUE)
  expect_match(x, '<a:prstDash val="dot"/>', fixed = TRUE)
})

test_that("every error-bar type but the standard error needs a value", {
  for (ty in c("fixed", "percentage", "std_dev"))
    expect_error(xl_chart_error_bars(ty), "needs a `value`", label = ty)
  expect_s3_class(xl_chart_error_bars("std_error"), "xl_chart_error_bars")
  expect_error(xl_chart_error_bars(), "must say how the bars are sized")
  expect_error(xl_chart_error_bars("guess", 1), "`type` must be one of")
  expect_error(xl_chart_error_bars("fixed", 1, direction = "up"),
               "`direction` must be one of")
})

# ── Individual points ─────────────────────────────────────────────────────────

test_that("points are styled one at a time, and NULL leaves one alone", {
  x <- part_xml(type = "pie",
                points = list(xl_fill(background = "red"),
                              xl_fill(background = "green"), NULL,
                              xl_fill(background = "blue")))
  # three styled, one left as it was
  expect_length(regmatches(x, gregexpr("<c:dPt>", x))[[1L]], 3L)
  expect_match(x, "00FF00", fixed = TRUE)
})

test_that("a point takes no font, and one format need not be a list", {
  expect_error(xl_chart_series(values = "A1:A5",
                               points = list(xl_font(bold = TRUE))),
               "a shape and has no text")
  expect_length(unclass(xl_chart_series(values = "A1:A5",
                                        points = xl_fill(background = "red")))$points,
                1L)
})

# ── The series arguments themselves ───────────────────────────────────────────

test_that("each part argument insists on its own constructor", {
  for (nm in names(.PART_CLASS)) {
    args <- list(values = "A1:A5")
    args[[nm]] <- "yes please"
    expect_error(do.call(xl_chart_series, args),
                 sprintf("`%s` must be an %s", nm, .PART_CLASS[[nm]]),
                 fixed = TRUE, label = nm)
  }
})

# ── Coverage of the C API ─────────────────────────────────────────────────────

test_that("every chart_series_*() function libxlsxwriter offers is called", {
  hdr <- NULL
  for (p in c("../../src/libxlsxwriter/include/xlsxwriter/chart.h",
              "../../../src/libxlsxwriter/include/xlsxwriter/chart.h"))
    if (file.exists(p)) hdr <- p
  skip_if(is.null(hdr), "libxlsxwriter sources not alongside the tests")
  src <- NULL
  for (p in c("../../src/write_xlsx.c", "../../../src/write_xlsx.c"))
    if (file.exists(p)) src <- p
  skip_if(is.null(src), "write_xlsx.c not alongside the tests")

  # declarations only: the header's prose mentions chart_series_set_marker_xxx()
  decl <- grep("^(void|lxw_error|lxw_series_error_bars \\*|lxw_chart_series \\*)",
               readLines(hdr), value = TRUE)
  fns <- unique(regmatches(decl, regexpr("chart_series_[a-z_]+", decl)))
  fns <- fns[nzchar(fns)]
  expect_equal(length(fns), 40L)

  called <- paste(readLines(src), collapse = "\n")
  for (f in fns)
    expect_true(grepl(paste0(f, "("), called, fixed = TRUE), info = f)
})

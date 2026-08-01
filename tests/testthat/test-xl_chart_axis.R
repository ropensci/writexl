# Chart axes.  Two things are being pinned here: that each of libxlsxwriter's
# 32 chart_axis_*() functions is reachable and reaches the file, and that the
# options only one *kind* of axis can use are refused on the other kind rather
# than written and ignored by Excel.

axis_sales <- data.frame(quarter = c("Q1", "Q2", "Q3", "Q4"),
                         revenue = c(10, 20, 30, 40),
                         stringsAsFactors = FALSE)

axis_ser <- function(...)
  xl_chart_series(values = list(cols = "revenue"),
                  categories = list(cols = "quarter"), ...)

# The chart XML for one chart over axis_sales.
axis_xml <- function(...) {
  t <- tempfile(fileext = ".xlsx")
  write_xlsx(list(Data = xl_sheet(axis_sales,
                                  chart = xl_chart("column", axis_ser(), ...))),
             t)
  d <- tempfile(); dir.create(d); utils::unzip(t, exdir = d)
  f <- list.files(file.path(d, "xl/charts"), pattern = "^chart",
                  full.names = TRUE)
  paste(readLines(f[1L], warn = FALSE), collapse = "")
}

# ── Which kind of axis is which ───────────────────────────────────────────────

test_that("a scatter chart has two value axes and everything else one", {
  # this is the whole basis of the option check, so it is pinned directly
  for (ty in names(.LXW_CHART_TYPE)) {
    if (identical(.CHART_FAMILY(ty), "pie")) next    # no axes at all
    expect_equal(.axis_kind(ty, "y"), "value", info = ty)
    expect_equal(.axis_kind(ty, "x"),
                 if (identical(.CHART_FAMILY(ty), "scatter")) "value"
                 else "category", info = ty)
  }
  # a bar chart draws its categories up the side, but they are still the x axis
  expect_equal(.axis_kind("bar", "x"), "category")
  expect_equal(.axis_kind("bar_stacked", "x"), "category")
})

test_that("a value-only option is refused on a category axis, and vice versa", {
  for (opt in names(.AXIS_OPTION_KIND)) {
    kind <- .AXIS_OPTION_KIND[[opt]]
    # a column chart has a category x and a value y, so one of the two is wrong
    wrong <- if (identical(kind, "value")) "x_axis" else "y_axis"
    right <- if (identical(kind, "value")) "y_axis" else "x_axis"
    val <- switch(opt, position = "on_tick", label_align = "left",
                  display_units = "thousands",
                  display_units_visible = TRUE, log_base = 10, 2)
    ax <- do.call(xl_chart_axis, stats::setNames(list(val), opt))
    args <- list("column", xl_chart_series(values = "B1:B5"))
    args[[wrong]] <- ax
    expect_error(do.call(xl_chart, args),
                 sprintf("$%s` applies to a %s axis", opt, kind),
                 fixed = TRUE, label = opt)
    args <- list("column", xl_chart_series(values = "B1:B5"))
    args[[right]] <- ax
    expect_true(inherits(do.call(xl_chart, args), "xl_chart"), info = opt)
  }
})

test_that("both axes of a scatter take the value-axis options", {
  se <- xl_chart_series(values = "B1:B5", categories = "A1:A5")
  expect_s3_class(xl_chart("scatter", se, x_axis = xl_chart_axis(min = 0),
                           y_axis = xl_chart_axis(log_base = 10)), "xl_chart")
  # and neither takes the category-axis ones, with a message that says why
  expect_error(xl_chart("scatter", se, x_axis = xl_chart_axis(interval_unit = 2)),
               "Both axes of a scatter chart are value axes", fixed = TRUE)
})

test_that("pie and doughnut charts have no axes to describe", {
  for (ty in c("pie", "doughnut"))
    expect_error(xl_chart(ty, xl_chart_series(values = "A1:A5"),
                          y_axis = xl_chart_axis(title = "n")),
                 "does not apply to a", label = ty)
})

test_that("the option-kind table matches the documented axis types", {
  # the table was read out of chart.h's "Axis types" lines; if a name here
  # stops being an xl_chart_axis() argument the table has drifted
  expect_true(all(names(.AXIS_OPTION_KIND) %in% names(formals(xl_chart_axis))))
  expect_setequal(unique(unname(.AXIS_OPTION_KIND)), c("value", "category"))
})

# ── The constructor ───────────────────────────────────────────────────────────

test_that("the enums cover libxlsxwriter's, in its order", {
  expect_equal(.LXW_AXIS_POSITION, c(on_tick = 1L, between = 2L))
  expect_equal(unname(.LXW_AXIS_LABEL_POSITION), 0:3)
  expect_equal(unname(.LXW_AXIS_LABEL_ALIGN), 0:2)
  expect_equal(unname(.LXW_AXIS_TICK_MARK), 0:4)
  expect_equal(unname(.LXW_AXIS_UNITS), 0:9)
  expect_true("crossing" %in% names(.LXW_AXIS_TICK_MARK))
})

test_that("axis arguments are validated", {
  expect_error(xl_chart_axis(position = "middle"), "`position` must be one of")
  expect_error(xl_chart_axis(major_tick = "big"), "`major_tick` must be one of")
  expect_error(xl_chart_axis(display_units = "squillions"),
               "`display_units` must be one of")
  expect_error(xl_chart_axis(log_base = 1), "must be between 2 and 1000")
  expect_error(xl_chart_axis(interval_unit = 0), "must be between 1 and NA")
  expect_error(xl_chart_axis(crossing = "middle"),
               "must be a number, \"min\" or \"max\"")
  expect_error(xl_chart_axis(crossing = c(1, 2)),
               "must be a number, \"min\" or \"max\"")
  expect_error(xl_chart_axis(num_format = 42), "must be an Excel format string")
  expect_error(xl_chart_axis(title_layout = c(0, 0.5)),
               "each above 0 and at most 1")
  expect_error(xl_chart_axis(title_layout = 0.5), "must be c\\(x, y\\)")
  expect_error(xl_chart("column", xl_chart_series(values = "A1"),
                        x_axis = "bottom"),
               "must be an xl_chart_axis")
})

test_that("a number format may be given either way", {
  expect_equal(unclass(xl_chart_axis(num_format = "#,##0"))$num_format, "#,##0")
  expect_equal(unclass(xl_chart_axis(num_format = xl_num_format("0.0%")))$num_format,
               "0.0%")
  # an xl_format that carries no number format is a mistake, not an empty one
  expect_error(xl_chart_axis(num_format = xl_font(bold = TRUE)),
               "no number format in it")
})

test_that("each part of an axis takes only the format groups it can use", {
  expect_error(xl_chart_axis(title = "n", title_format = xl_fill(background = "red")),
               "takes no fill")
  expect_error(xl_chart_axis(label_format = xl_border(all = "thin")),
               "takes no line")
  expect_error(xl_chart_axis(major_gridlines_format = xl_fill(background = "red")),
               "takes no fill")
  # the axis line itself is a shape, so it takes no font
  expect_error(xl_chart_axis(line_format = xl_font(bold = TRUE)),
               "a shape and has no text")
})

test_that("a title format needs a title to style", {
  expect_error(xl_chart_axis(title_format = xl_font(bold = TRUE)),
               "give a `title` too", fixed = TRUE)
  expect_s3_class(xl_chart_axis(title = "n", title_format = xl_font(bold = TRUE)),
                  "xl_chart_axis")
})

test_that("styling a gridline turns it on", {
  # a styled gridline that stays hidden is not something anyone means
  p <- .axis_payload(xl_chart_axis(major_gridlines_format = xl_border(all = "thin")),
                     "y_axis", list(Data = axis_sales), "Data", 1L)
  expect_equal(p$major_gridlines, 1L)
  expect_false(is.null(p$major_gridlines_line))
  # unless it is switched off explicitly, which stays off
  q <- .axis_payload(xl_chart_axis(major_gridlines = FALSE,
                                   major_gridlines_format = xl_border(all = "thin")),
                     "y_axis", list(Data = axis_sales), "Data", 1L)
  expect_equal(q$major_gridlines, 0L)
})

# ── What reaches the file ─────────────────────────────────────────────────────

test_that("every axis option reaches the chart XML", {
  x <- axis_xml(
    x_axis = xl_chart_axis(title = "Quarter", position = "on_tick",
                           label_align = "left", interval_unit = 2,
                           interval_tick = 3, major_tick = "inside",
                           minor_tick = "crossing", label_position = "low",
                           title_format = xl_font(size = 12, bold = TRUE),
                           label_format = xl_font(italic = TRUE),
                           line_format = xl_border(all = "dashed",
                                                   color = "red")),
    y_axis = xl_chart_axis(title = list(header = "revenue"), min = 5, max = 45,
                           major_unit = 10, minor_unit = 2.5,
                           num_format = "$#,##0.0", reverse = TRUE,
                           crossing = "max", display_units = "thousands",
                           display_units_visible = TRUE,
                           major_gridlines = TRUE, minor_gridlines = TRUE,
                           minor_gridlines_format = xl_border(all = "dotted",
                                                              color = "blue"),
                           title_layout = c(0.02, 0.4)))
  want <- c(
    "Quarter",                              # name
    "<c:f>Data!$B$1</c:f>",                 # name from a header cell
    '<c:tickLblPos val="low"/>',            # label_position
    '<c:lblAlgn val="l"/>',                 # label_align
    '<c:tickLblSkip val="2"/>',             # interval_unit
    '<c:tickMarkSkip val="3"/>',            # interval_tick
    '<c:majorTickMark val="in"/>',          # major_tick
    '<c:minorTickMark val="cross"/>',       # minor_tick
    '<a:prstDash val="dash"/>',             # the axis line
    '<c:min val="5"/>', '<c:max val="45"/>',
    '<c:majorUnit val="10"/>', '<c:minorUnit val="2.5"/>',
    'formatCode="$#,##0.0"',                # num_format
    '<c:orientation val="maxMin"/>',        # reverse
    '<c:crosses val="max"/>',               # crossing
    '<c:builtInUnit val="thousands"/>',     # display_units
    "<c:dispUnitsLbl>",                     # display_units_visible
    "<c:majorGridlines/>", "<c:minorGridlines>",
    '<a:prstDash val="dot"/>',              # the minor gridline's own line
    "0000FF",                               # and its colour
    "<c:manualLayout>",                     # title_layout
    'sz="1200"', 'i="1"'                    # the two fonts
  )
  for (w in want) expect_true(grepl(w, x, fixed = TRUE), info = w)
  # position = "on_tick" is written as the absence of crossBetween="between"
  expect_match(x, '<c:crossBetween val="midCat"/>', fixed = TRUE)
})

test_that("an axis can be hidden, and its gridlines and labels turned off", {
  x <- axis_xml(x_axis = xl_chart_axis(visible = FALSE),
                y_axis = xl_chart_axis(major_gridlines = FALSE,
                                       label_position = "none"))
  expect_match(x, '<c:delete val="1"/>', fixed = TRUE)
  expect_false(grepl("<c:majorGridlines", x, fixed = TRUE))
  expect_match(x, '<c:tickLblPos val="none"/>', fixed = TRUE)
})

test_that("a log scale reaches both axes of a scatter chart", {
  t <- tempfile(fileext = ".xlsx")
  write_xlsx(list(Data = xl_sheet(axis_sales, chart = xl_chart("scatter",
    xl_chart_series(values = list(cols = "revenue"),
                    categories = list(cols = "revenue")),
    x_axis = xl_chart_axis(log_base = 10),
    y_axis = xl_chart_axis(log_base = 2)))), t)
  d <- tempfile(); dir.create(d); utils::unzip(t, exdir = d)
  x <- paste(readLines(list.files(file.path(d, "xl/charts"), pattern = "^chart",
                                  full.names = TRUE)[1L], warn = FALSE),
             collapse = "")
  expect_match(x, '<c:logBase val="10"/>', fixed = TRUE)
  expect_match(x, '<c:logBase val="2"/>', fixed = TRUE)
})

test_that("an axis name may come from another sheet", {
  t <- tempfile(fileext = ".xlsx")
  write_xlsx(list(
    Chart = xl_sheet(data.frame(z = 1),
                     chart = xl_chart("column",
                       xl_chart_series(values = list(sheet = "Data",
                                                     cols = "revenue")),
                       y_axis = xl_chart_axis(title = list(sheet = "Data",
                                                          header = "revenue")))),
    Data = axis_sales), t)
  d <- tempfile(); dir.create(d); utils::unzip(t, exdir = d)
  x <- paste(readLines(list.files(file.path(d, "xl/charts"), pattern = "^chart",
                                  full.names = TRUE)[1L], warn = FALSE),
             collapse = "")
  expect_match(x, "<c:f>Data!$B$1</c:f>", fixed = TRUE)
})

# ── Coverage of the C API ─────────────────────────────────────────────────────

test_that("every chart_axis_*() function libxlsxwriter offers is called", {
  # A mechanical gate rather than a promise in prose: a function added
  # upstream, or one dropped from apply_axis(), fails here.  Skipped where the
  # sources are not alongside the tests, as in an installed package.
  hdr <- NULL
  for (p in c("../../src/libxlsxwriter/include/xlsxwriter/chart.h",
              "../../../src/libxlsxwriter/include/xlsxwriter/chart.h"))
    if (file.exists(p)) hdr <- p
  skip_if(is.null(hdr), "libxlsxwriter sources not alongside the tests")
  src <- NULL
  for (p in c("../../src/write_xlsx.c", "../../../src/write_xlsx.c"))
    if (file.exists(p)) src <- p
  skip_if(is.null(src), "write_xlsx.c not alongside the tests")

  decl <- grep("^(void|lxw_chart_axis \\*|uint8_t) ?chart_axis",
               readLines(hdr), value = TRUE)
  fns <- unique(regmatches(decl, regexpr("chart_axis_[a-z_]+", decl)))
  # chart_axis_get() fetches an axis by enum; writexl reaches the axes through
  # the chart struct's own x_axis/y_axis members instead.
  fns <- setdiff(fns, "chart_axis_get")
  expect_equal(length(fns), 32L)

  called <- paste(readLines(src), collapse = "\n")
  for (f in fns)
    expect_true(grepl(paste0(f, "("), called, fixed = TRUE), info = f)
})

test_that("the display-units caption is on unless it is turned off", {
  # chart_axis_set_display_units() sets display_units_visible itself, so the
  # caption comes with the rescaling and `TRUE` adds nothing.  Pinned because
  # the argument reads as opt-in, and because a change upstream would silently
  # alter every such chart.
  on <- axis_xml(y_axis = xl_chart_axis(display_units = "thousands"))
  expect_match(on, "<c:dispUnitsLbl>", fixed = TRUE)
  same <- axis_xml(y_axis = xl_chart_axis(display_units = "thousands",
                                          display_units_visible = TRUE))
  expect_match(same, "<c:dispUnitsLbl>", fixed = TRUE)
  off <- axis_xml(y_axis = xl_chart_axis(display_units = "thousands",
                                         display_units_visible = FALSE))
  expect_match(off, '<c:builtInUnit val="thousands"/>', fixed = TRUE)
  expect_false(grepl("<c:dispUnitsLbl>", off, fixed = TRUE))
})

test_that("only the options in the kind table are checked against the axis", {
  # .check_axis() intersects with .AXIS_OPTION_KIND before checking, so an
  # option outside the table -- `title`, `visible`, the gridlines -- reaches
  # every axis of every chart type untouched
  untyped <- setdiff(names(formals(xl_chart_axis)), names(.AXIS_OPTION_KIND))
  expect_true(length(untyped) > 0L)
  for (ty in c("column", "scatter", "line"))
    expect_s3_class(xl_chart(ty, xl_chart_series(values = "A1:A5",
                                                 categories = "B1:B5"),
                             x_axis = xl_chart_axis(title = "t",
                                                    visible = FALSE),
                             y_axis = xl_chart_axis(title = "t",
                                                    visible = FALSE)),
                    "xl_chart")
  # and every option that IS in the table belongs to exactly one kind
  expect_true(all(.AXIS_OPTION_KIND %in% c("category", "value")))
})

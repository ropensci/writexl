# Charts.  This file covers the constructors, the chart-type feature matrix and
# range specs; what reaches the file is tested once apply_charts() lands.

# ── Chart types ───────────────────────────────────────────────────────────────

test_that("every chart type libxlsxwriter offers is reachable", {
  # the enum runs 1..22 with no gaps, so a type added upstream shows up here as
  # a count mismatch rather than being quietly unreachable
  expect_equal(length(.LXW_CHART_TYPE), 22L)
  expect_equal(sort(unname(.LXW_CHART_TYPE)), 1:22)
  for (ty in names(.LXW_CHART_TYPE)) {
    # scatter needs categories; see the crash test below
    se <- if (identical(.CHART_FAMILY(ty), "scatter"))
      xl_chart_series(values = "B1:B5", categories = "A1:A5")
    else xl_chart_series(values = "A1:A5")
    expect_true(inherits(xl_chart(ty, se), "xl_chart"), info = ty)
  }
})

test_that("an unknown chart type is refused, listing the real ones", {
  expect_error(xl_chart("bubble", xl_chart_series(values = "A1:A5")),
               "`type` must be one of")
  expect_error(xl_chart("bubble", xl_chart_series(values = "A1:A5")),
               "doughnut")
})

test_that("every type maps to a family", {
  fams <- vapply(names(.LXW_CHART_TYPE), .CHART_FAMILY, character(1))
  expect_false(any(fams == "other"))
  expect_setequal(unique(fams),
                  c("area", "bar", "pie", "line", "scatter", "radar"))
})

# ── The feature matrix ────────────────────────────────────────────────────────

test_that("type-restricted features match what libxlsxwriter documents", {
  # these are documented in prose and not enforced: Excel drops them silently,
  # so the refusal has to be ours
  expect_true(.chart_supports("doughnut", "hole_size"))
  expect_false(.chart_supports("pie", "hole_size"))   # doughnut only, not both

  expect_true(.chart_supports("pie", "rotation"))
  expect_true(.chart_supports("doughnut", "rotation"))
  expect_false(.chart_supports("column", "rotation"))

  for (f in c("up_down_bars", "high_low_lines")) {
    expect_true(.chart_supports("line", f), label = f)
    expect_false(.chart_supports("column", f), label = f)
    expect_false(.chart_supports("pie", f), label = f)
  }

  for (f in c("series_gap", "series_overlap")) {
    expect_true(.chart_supports("bar", f), label = f)
    expect_true(.chart_supports("column_stacked", f), label = f)
    expect_false(.chart_supports("line", f), label = f)
  }

  expect_true(.chart_supports("line", "smooth"))
  expect_true(.chart_supports("scatter", "smooth"))
  expect_false(.chart_supports("column", "smooth"))

  # pie and doughnut have no axes at all
  expect_false(.chart_supports("pie", "axes"))
  expect_false(.chart_supports("doughnut", "axes"))
  expect_true(.chart_supports("column", "axes"))
  expect_true(.chart_supports("scatter", "axes"))
})

test_that("every feature is supported by at least one type, and denied by one", {
  # a feature nothing supports would be dead code; one everything supports
  # would not need to be in the matrix
  for (f in names(.CHART_FEATURE_FAMILIES)) {
    ok <- vapply(names(.LXW_CHART_TYPE), .chart_supports, logical(1), feature = f)
    expect_true(any(ok), label = f)
    expect_true(any(!ok), label = f)
  }
})

test_that("an unsupported feature is refused, naming the types that work", {
  expect_error(xl_chart("column", xl_chart_series(values = "A1", smooth = TRUE)),
               'does not apply to a "column" chart')
  expect_error(xl_chart("column", xl_chart_series(values = "A1", smooth = TRUE)),
               "It applies to: line")
  # and is accepted where it belongs
  expect_s3_class(xl_chart("line", xl_chart_series(values = "A1", smooth = TRUE)),
                  "xl_chart")
  expect_s3_class(xl_chart("scatter_smooth",
                           xl_chart_series(values = "B1:B5", smooth = TRUE,
                                           categories = "A1:A5")),
                  "xl_chart")
})

# ── Series ────────────────────────────────────────────────────────────────────

test_that("a series carries its ranges", {
  s <- unclass(xl_chart_series(values = "Data!B2:B10",
                               categories = "Data!A2:A10", name = "2024"))
  expect_equal(s$values$spec, "Data!B2:B10")
  expect_equal(s$categories$spec, "Data!A2:A10")
  expect_equal(s$name$text, "2024")
})

test_that("a range may name another sheet", {
  s <- unclass(xl_chart_series(values = list(sheet = "Data", cols = "revenue")))
  expect_equal(s$values$sheet, "Data")
  expect_equal(s$values$spec, list(cols = "revenue"))
  # without a sheet it resolves against the chart's own
  s2 <- unclass(xl_chart_series(values = list(cols = "revenue")))
  expect_null(s2$values$sheet)
})

test_that("a series name is literal text, and a range needs a spec", {
  # a string is never guessed at: a series may legitimately be called "Q1!",
  # so taking the name from a cell has to be asked for explicitly
  expect_equal(unclass(xl_chart_series(values = "A1", name = "2024"))$name$text,
               "2024")
  expect_equal(unclass(xl_chart_series(values = "A1",
                                       name = "Data!A1"))$name$text,
               "Data!A1")
  from_cell <- unclass(xl_chart_series(
    values = "A1", name = list(sheet = "Data", rows = 1, cols = 1)))
  expect_equal(from_cell$name$sheet, "Data")
  expect_null(from_cell$name$text)
})

test_that("series arguments are validated", {
  expect_error(xl_chart_series(), "must name the range")
  expect_error(xl_chart_series(values = 42), "must be a range string")
  expect_error(xl_chart_series(values = list(nope = "a")),
               "unknown `values` element\\(s\\): nope")
  expect_error(xl_chart_series(values = list(sheet = 1, cols = "a")),
               "must be a single sheet name")
  expect_error(xl_chart_series(values = list("a")), "list must be named")
  expect_error(xl_chart_series(values = "A1", format = "bold"),
               "must be an xl_format")
  expect_error(xl_chart_series(values = "A1", smooth = "yes"),
               "`smooth` must be")
})

# ── Charts ────────────────────────────────────────────────────────────────────

test_that("a chart carries its type, series and title", {
  ch <- unclass(xl_chart("column",
                         list(xl_chart_series(values = "A1:A5"),
                              xl_chart_series(values = "B1:B5")),
                         title = "Revenue"))
  expect_equal(ch$type, "column")
  expect_length(ch$series, 2L)
  expect_equal(ch$title$text, "Revenue")
})

test_that("title = FALSE removes the title Excel would generate", {
  # a single-series chart gets an automatic title, so switching it off has to
  # be expressible -- NULL means "leave Excel to it", which is different
  expect_true(isTRUE(unclass(xl_chart("pie", xl_chart_series(values = "A1"),
                                      title = FALSE))$title$off))
  expect_null(unclass(xl_chart("pie", xl_chart_series(values = "A1")))$title)
})

test_that("placement uses the same vocabulary as an image", {
  # lxw_chart_options is lxw_image_options minus url/tip/cell_format, so the
  # R names must not drift apart
  shared <- c("at", "scale", "offset", "position", "description", "decorative")
  expect_true(all(shared %in% names(formals(xl_chart))))
  expect_true(all(shared %in% names(formals(xl_image))))
  # and mean the same things
  ch <- unclass(xl_chart("column", xl_chart_series(values = "A1"),
                         at = "E2", scale = c(2, 3), offset = c(4, 5),
                         position = "dont_move_dont_size"))
  expect_equal(ch$at, "E2")
  expect_equal(ch$scale, c(2, 3))
  expect_equal(ch$offset, c(4, 5))
  expect_equal(ch$position, "dont_move_dont_size")
})

test_that("chart arguments are validated", {
  expect_error(xl_chart("column"), "`series` must give at least one")
  expect_error(xl_chart("column", list()), "needs at least one series")
  expect_error(xl_chart("column", "series"), "must be an xl_chart_series")
  expect_error(xl_chart("column", list(xl_chart_series(values = "A1"), "x")),
               "`series\\[\\[2\\]\\]` must be an xl_chart_series")
  expect_error(xl_chart("column", xl_chart_series(values = "A1"), style = 0),
               "`style` must be between 1 and 48")
  expect_error(xl_chart("column", xl_chart_series(values = "A1"), style = 49),
               "`style` must be between 1 and 48")
  expect_error(xl_chart("column", xl_chart_series(values = "A1"), scale = -1),
               "`scale` must be positive")
  expect_error(xl_chart("column", xl_chart_series(values = "A1"),
                        description = 1), "single non-NA string")
})

test_that("charts normalise from one or a list", {
  expect_length(.chart_list(xl_chart("pie", xl_chart_series(values = "A1"))), 1L)
  expect_length(.chart_list(list(xl_chart("pie", xl_chart_series(values = "A1")),
                                 xl_chart("bar", xl_chart_series(values = "A1")))),
                2L)
  expect_length(.chart_list(NULL), 0L)
  expect_error(.chart_list("chart"), "must be an xl_chart object")
})

test_that("the print methods run", {
  expect_output(print(xl_chart("column", xl_chart_series(values = "A1"))),
                "xl_chart")
  expect_output(print(xl_chart("column", xl_chart_series(values = "A1"),
                               title = "T")), "\"T\"")
  expect_output(print(xl_chart_series(values = "A1", name = "S")), "S")
})

# ── Range resolution across the workbook ──────────────────────────────────────

crange_sales <- data.frame(fruit = c("a", "b", "c"), qty = c(5, 150, 300),
                           stringsAsFactors = FALSE)
crange_other <- data.frame(x = 1:4, y = c(2, 4, 6, 8))

# Resolve a workbook's charts the way write_xlsx() does.
chart_plan <- function(wb, sheet = "Data") {
  dfs <- lapply(wb, function(e) if (inherits(e, "xl_sheet")) e$data else e)
  dfs <- lapply(dfs, writexl:::normalize_df)
  .resolve_charts(wb[[sheet]], dfs[[sheet]], .new_format_registry(), 1L,
                  xl_properties(), dfs, sheet)
}

test_that("a series resolves against the chart's own sheet by default", {
  p <- chart_plan(list(Data = xl_sheet(crange_sales, chart = xl_chart("column",
    xl_chart_series(values = list(cols = "qty"),
                    categories = list(cols = "fruit"))))))
  s <- p[[1L]]$series[[1L]]
  expect_equal(s$values_sheet, "Data")
  # column qty is 0-based col 1, data rows 1..3 are sheet rows 1..3
  expect_equal(s$values_range, c(1L, 1L, 3L, 1L))
  expect_equal(s$categories_range, c(1L, 0L, 3L, 0L))
})

test_that("a series may plot another sheet, in either spelling", {
  wb <- function(v) list(Data = xl_sheet(crange_sales,
                                         chart = xl_chart("line",
                                           xl_chart_series(values = v))),
                         Other = crange_other)
  a <- chart_plan(wb("Other!B2:B5"))[[1L]]$series[[1L]]
  b <- chart_plan(wb(list(sheet = "Other", cols = "y")))[[1L]]$series[[1L]]
  expect_equal(a$values_sheet, "Other")
  expect_equal(b$values_sheet, "Other")
  # the two spellings must resolve to the same block
  expect_equal(a$values_range, b$values_range)
})

test_that("an unknown sheet is refused, listing the real ones", {
  expect_error(chart_plan(list(Data = xl_sheet(crange_sales,
    chart = xl_chart("column",
      xl_chart_series(values = list(sheet = "Nope", cols = "qty")))))),
    'names sheet "Nope", which is not in the workbook')
  expect_error(chart_plan(list(Data = xl_sheet(crange_sales,
    chart = xl_chart("column",
      xl_chart_series(values = list(sheet = "Nope", cols = "qty")))))),
    '"Data"')
})

test_that("a range holding no data is refused, not plotted empty", {
  # Excel shows an empty chart with no complaint, so the emptiness has to be
  # caught here or it is delivered as a silently blank chart
  expect_error(chart_plan(list(Data = xl_sheet(crange_sales,
    chart = xl_chart("column", xl_chart_series(values = "A99:A200"))))),
    "selects no data")
  expect_error(chart_plan(list(Data = xl_sheet(crange_sales[0, ],
    chart = xl_chart("column", xl_chart_series(values = "A2:A3"))))),
    "selects no data")
})

test_that("a title or series name may come from a cell", {
  p <- chart_plan(list(Data = xl_sheet(crange_sales, chart = xl_chart("pie",
    xl_chart_series(values = list(cols = "qty"),
                    name = list(rows = 1, cols = "fruit")),
    title = list(rows = 1, cols = "fruit")))))
  expect_equal(p[[1L]]$title_range, c(1L, 0L, 1L, 0L))
  expect_equal(p[[1L]]$series[[1L]]$name_range, c(1L, 0L, 1L, 0L))
  # a string title stays a string
  q <- chart_plan(list(Data = xl_sheet(crange_sales, chart = xl_chart("pie",
    xl_chart_series(values = list(cols = "qty")), title = "Share"))))
  expect_equal(q[[1L]]$title, "Share")
  expect_null(q[[1L]]$title_range)
})

test_that("a chart is anchored to one cell", {
  expect_error(chart_plan(list(Data = xl_sheet(crange_sales,
    chart = xl_chart("column", xl_chart_series(values = list(cols = "qty")),
                     at = "A1:B2")))),
    "must name a single cell")
})

test_that("a chart is a victim of the drawing-id desync, like an image", {
  # charts produce a drawing so they never cause the desync, but they are
  # numbered from the same counter and so are subject to it.  Established with
  # a pure-C reprex before charts existed here; nothing else would notice.
  logo <- tempfile(fileext = ".png")
  writeBin(as.raw(c(
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
    0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00,
    0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
    0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82)), logo)
  hdr <- xl_sheet(crange_sales,
                  page = xl_page_setup(header = "&L&G",
                                       header_image = list(left = logo)))
  charted <- xl_sheet(crange_sales, chart = xl_chart("column",
    xl_chart_series(values = list(cols = "qty"))))
  expect_error(write_tmp(list(H = hdr, C = charted)),
               'sheet "C" has a chart')
  # the other order is fine, and is what the message recommends
  expect_silent(write_tmp(list(C = charted, H = hdr)))
})

# ── What reaches the file ─────────────────────────────────────────────────────

cfile <- function(..., df = crange_sales) {
  p <- write_tmp(list(Data = xl_sheet(df, ...)))
  d <- tempfile(); dir.create(d); utils::unzip(p, exdir = d)
  ch <- list.files(file.path(d, "xl/charts"), pattern = "^chart",
                   full.names = TRUE)
  list(files = list.files(d, recursive = TRUE),
       chart = if (length(ch)) paste(readLines(ch[1L], warn = FALSE),
                                     collapse = "") else "",
       sheet = xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE))
}

test_that("a chart reaches the file with its drawing and reference", {
  r <- cfile(chart = xl_chart("column",
    xl_chart_series(values = list(cols = "qty"),
                    categories = list(cols = "fruit"), name = "Qty"),
    title = "Fruit", at = "D2"))
  expect_true(any(grepl("^xl/charts/chart", r$files)))
  expect_true(any(grepl("^xl/drawings/drawing", r$files)))
  expect_match(r$sheet, "<drawing r:id=", fixed = TRUE)
  expect_match(r$chart, "<c:barChart>", fixed = TRUE)
  # the series points at the resolved range, absolute and sheet-qualified
  expect_match(r$chart, "<c:f>Data!$B$2:$B$4</c:f>", fixed = TRUE)
  expect_match(r$chart, "<c:f>Data!$A$2:$A$4</c:f>", fixed = TRUE)
  expect_match(r$chart, "Fruit", fixed = TRUE)
  expect_match(r$chart, "Qty", fixed = TRUE)
})

test_that("each chart type writes its own plot element", {
  kinds <- c(column = "barChart", bar = "barChart", line = "lineChart",
             pie = "pieChart", doughnut = "doughnutChart", area = "areaChart",
             scatter = "scatterChart", radar = "radarChart")
  for (ty in names(kinds)) {
    se <- if (identical(.CHART_FAMILY(ty), "scatter"))
      xl_chart_series(values = list(cols = "qty"),
                      categories = list(cols = "fruit"))
    else xl_chart_series(values = list(cols = "qty"))
    expect_match(cfile(chart = xl_chart(ty, se))$chart,
                 sprintf("<c:%s>", kinds[[ty]]), fixed = TRUE, label = ty)
  }
})

test_that("a series format reaches the chart's shape properties", {
  r <- cfile(chart = xl_chart("line", xl_chart_series(
    values = list(cols = "qty"),
    format = xl_border(all = "dashed", color = "red") +
             xl_fill(background = "yellow", transparency = 30))))
  # charts use 6-digit RGB, not the 8-digit ARGB of cell styles
  expect_match(r$chart, "<a:srgbClr val=\"FF0000\"/>", fixed = TRUE)
  expect_match(r$chart, "<a:srgbClr val=\"FFFF00\">", fixed = TRUE)
  expect_match(r$chart, "<a:prstDash val=\"dash\"/>", fixed = TRUE)
  # transparency 30 becomes 70% opacity
  expect_match(r$chart, "<a:alpha val=\"70000\"/>", fixed = TRUE)
})

test_that("a cross-sheet series names the other sheet", {
  p <- write_tmp(list(
    Chart = xl_sheet(data.frame(z = 1), chart = xl_chart("pie",
      xl_chart_series(values = list(sheet = "Data", cols = "qty")))),
    Data = crange_sales))
  d <- tempfile(); dir.create(d); utils::unzip(p, exdir = d)
  x <- paste(readLines(list.files(file.path(d, "xl/charts"), pattern = "^chart",
                                  full.names = TRUE)[1L], warn = FALSE),
             collapse = "")
  expect_match(x, "Data!", fixed = TRUE)
})

test_that("several charts on one sheet share its drawing", {
  r <- cfile(chart = list(
    xl_chart("column", xl_chart_series(values = list(cols = "qty")), at = "D2"),
    xl_chart("pie", xl_chart_series(values = list(cols = "qty")), at = "D20")))
  expect_equal(sum(grepl("^xl/charts/chart", r$files)), 2L)
  # one drawing per sheet, two charts anchored in it
  expect_equal(sum(grepl("^xl/drawings/drawing", r$files)), 1L)
})

test_that("placement and style options are written", {
  r <- cfile(chart = xl_chart("column",
    xl_chart_series(values = list(cols = "qty")),
    at = "A1", offset = c(20, 10), style = 12))
  expect_match(r$chart, "<c:style val=\"12\"/>", fixed = TRUE)
  d <- tempfile(); dir.create(d)
  p <- write_tmp(list(Data = xl_sheet(crange_sales, chart = xl_chart("column",
    xl_chart_series(values = list(cols = "qty")), at = "A1",
    offset = c(20, 10)))))
  utils::unzip(p, exdir = d)
  dr <- paste(readLines(file.path(d, "xl/drawings/drawing1.xml"), warn = FALSE),
              collapse = "")
  expect_match(dr, sprintf("<xdr:colOff>%d</xdr:colOff>", 20 * 9525),
               fixed = TRUE)
})

test_that("title = FALSE writes autoTitleDeleted", {
  expect_match(cfile(chart = xl_chart("pie",
    xl_chart_series(values = list(cols = "qty")), title = FALSE))$chart,
    "<c:autoTitleDeleted val=\"1\"/>", fixed = TRUE)
})

test_that("every relationship a chart adds resolves", {
  p <- write_tmp(list(Data = xl_sheet(crange_sales, chart = xl_chart("column",
    xl_chart_series(values = list(cols = "qty"))))))
  d <- tempfile(); dir.create(d); utils::unzip(p, exdir = d)
  bad <- character(0)
  for (rel in list.files(d, recursive = TRUE, pattern = "_rels/")) {
    x <- paste(readLines(file.path(d, rel), warn = FALSE), collapse = "")
    base <- dirname(dirname(rel))
    for (tg in regmatches(x, gregexpr('Target="[^"]+"[^>]*', x))[[1L]]) {
      if (grepl('TargetMode="External"', tg)) next
      target <- sub('Target="([^"]+)".*', "\\1", tg)
      q <- if (startsWith(target, "/")) sub("^/", "", target)
           else if (base == ".") target else file.path(base, target)
      while (grepl("/[^/]+/\\.\\./", q)) q <- sub("/[^/]+/\\.\\./", "/", q)
      if (!file.exists(file.path(d, q))) bad <- c(bad, target)
    }
  }
  expect_equal(bad, character(0))
})

test_that("a scatter series must have categories", {
  # not a style preference: libxlsxwriter's _chart_write_cat() reads
  # series->categories->has_string_cache before its own NULL guard, so a
  # scatter series without categories segfaults.  Found by bisecting a crash in
  # this very file, and reduced to a check that says what to do.
  for (ty in c("scatter", "scatter_straight", "scatter_straight_markers",
               "scatter_smooth", "scatter_smooth_markers"))
    expect_error(xl_chart(ty, xl_chart_series(values = "B1:B5")),
                 "no `categories`", label = ty)
  # with them, every scatter type writes
  for (ty in c("scatter", "scatter_straight", "scatter_smooth"))
    expect_true(inherits(xl_chart(ty, xl_chart_series(values = "B1:B5",
                                                      categories = "A1:A5")),
                         "xl_chart"), info = ty)
  # and no other family is affected
  expect_s3_class(xl_chart("column", xl_chart_series(values = "A1:A5")),
                  "xl_chart")
})

test_that("every chart type writes a file without crashing", {
  # the scatter crash reached the C layer, where a bad assumption is a
  # segfault rather than a failed expectation, so every type is exercised
  # end to end
  for (ty in names(.LXW_CHART_TYPE)) {
    se <- if (identical(.CHART_FAMILY(ty), "scatter"))
      xl_chart_series(values = list(cols = "qty"),
                      categories = list(cols = "fruit"))
    else xl_chart_series(values = list(cols = "qty"))
    expect_silent(write_tmp(list(Data = xl_sheet(crange_sales,
                                                 chart = xl_chart(ty, se)))))
  }
})

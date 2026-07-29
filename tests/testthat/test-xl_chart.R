# Charts.  This file covers the constructors, the chart-type feature matrix and
# range specs; what reaches the file is tested once apply_charts() lands.

# ── Chart types ───────────────────────────────────────────────────────────────

test_that("every chart type libxlsxwriter offers is reachable", {
  # the enum runs 1..22 with no gaps, so a type added upstream shows up here as
  # a count mismatch rather than being quietly unreachable
  expect_equal(length(.LXW_CHART_TYPE), 22L)
  expect_equal(sort(unname(.LXW_CHART_TYPE)), 1:22)
  for (ty in names(.LXW_CHART_TYPE))
    expect_s3_class(xl_chart(ty, xl_chart_series(values = "A1:A5")), "xl_chart")
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
                           xl_chart_series(values = "A1", smooth = TRUE)),
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

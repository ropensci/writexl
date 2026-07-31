# Translating an ordinary xl_format into the line / fill / pattern / font a
# chart understands.  The translation is lossy in known ways, and every loss is
# an error -- these tests exist mostly to pin that.

cf <- function(f) .chart_format_payload(f, "format")

# ── What translates ───────────────────────────────────────────────────────────

test_that("a font maps to the chart font's five properties", {
  # lxw_chart_font carries size, bold, italic, underline and colour -- and
  # nothing else, which is why the typeface is refused below
  ent <- cf(xl_font(bold = TRUE, italic = TRUE, size = 12, color = "navy",
                    underline = "single"))$font
  expect_equal(ent$size, 12)
  expect_equal(ent$bold, 1L)
  expect_equal(ent$italic, 1L)
  expect_equal(ent$underline, 1L)
  expect_equal(ent$color, xl_color("navy"))
})

test_that("Excel's five underline styles collapse to a flag", {
  # the chart font has a boolean, not a choice, so anything but "none" is on
  for (u in c("single", "double", "single-accounting", "double-accounting"))
    expect_equal(cf(xl_font(underline = u))$font$underline, 1L, label = u)
  expect_null(cf(xl_font(underline = "none", bold = TRUE))$font$underline)
})

test_that("a plain fill maps to the chart fill, with transparency", {
  ent <- cf(xl_fill(background = "red", transparency = 40))$fill
  expect_equal(ent$color, xl_color("red"))
  expect_equal(ent$transparency, 40L)
  # and "none" means no fill rather than a colour
  expect_equal(cf(xl_fill(pattern = "none"))$fill$none, 1L)
})

test_that("a real pattern maps to the chart pattern, not the fill", {
  p <- cf(xl_fill(pattern = "gray-125", foreground = "red",
                  background = "white"))
  expect_null(p$fill)
  expect_equal(p$pattern$fg_color, xl_color("red"))
  expect_equal(p$pattern$bg_color, xl_color("white"))
  expect_true(is.numeric(p$pattern$type))
})

test_that("a border maps to the chart line, with transparency", {
  ent <- cf(xl_border(all = "dashed", color = "gray", transparency = 25))$line
  expect_equal(ent$dash_type, 3L)
  expect_equal(ent$color, xl_color("gray"))
  expect_equal(ent$transparency, 25L)
  expect_equal(cf(xl_border(all = "none"))$line$none, 1L)
})

test_that("the four supported border styles map to their exact dash types", {
  # only exact equivalents are mapped; the table is short on purpose, since the
  # alternative is approximating a style the user asked for
  expect_equal(cf(xl_border(all = "thin"))$line$dash_type, 0L)      # solid
  expect_equal(cf(xl_border(all = "dashed"))$line$dash_type, 3L)
  expect_equal(cf(xl_border(all = "dash-dot"))$line$dash_type, 4L)
  expect_equal(cf(xl_border(all = "dotted"))$line$dash_type, 8L)
  expect_equal(length(.CHART_DASH), 4L)
})

test_that("groups combine into one payload", {
  p <- cf(xl_font(bold = TRUE) + xl_fill(background = "yellow") +
            xl_border(all = "thin", color = "black"))
  expect_named(p, c("font", "fill", "line"), ignore.order = TRUE)
})

# ── What is refused ───────────────────────────────────────────────────────────

test_that("every border style a chart cannot draw is refused, with a reason", {
  # enumerated: each of the 14 cell styles is either mapped or refused, so a
  # style added to writexl cannot silently fall through
  mapped <- names(.CHART_DASH)
  refused <- names(.CHART_DASH_UNSUPPORTED)
  all_styles <- setdiff(names(writexl:::.LXW$border), "none")
  expect_setequal(c(mapped, refused), all_styles)

  for (st in refused)
    expect_error(cf(xl_border(all = st)), "cannot be drawn on a chart",
                 label = st)
  # the width cases say so, and the pattern cases say something different
  expect_error(cf(xl_border(all = "thick")), "no width")
  expect_error(cf(xl_border(all = "slant-dash-dot")),
               "no chart dash pattern matching it")
  # and the message lists what would have worked
  expect_error(cf(xl_border(all = "medium")), "Chart lines accept")
})

test_that("the format groups a chart shape has no concept of are refused", {
  expect_error(cf(xl_num_format("0.0")), "set it on the axis or the data labels")
  expect_error(cf(xl_align(horizontal = "center")), "no text alignment")
  expect_error(cf(xl_protection(locked = FALSE)), "cannot be locked or hidden")
  # a typeface: lxw_chart_font has no font name field at all
  expect_error(cf(xl_font(name = "Arial")), "carries no typeface")
})

test_that("a chart has one line, so differing sides are refused", {
  expect_error(cf(xl_border(left = "thin", right = "dashed")),
               "a chart has one line")
  expect_error(cf(xl_border(all = "thin", left_color = "red",
                            right_color = "blue")),
               "a chart has one line")
  # all four the same is fine
  expect_silent(cf(xl_border(all = "thin", color = "red")))
})

test_that("a pattern without both colours is refused", {
  expect_error(cf(xl_fill(pattern = "gray-125")), "give both")
  expect_error(cf(xl_fill(pattern = "gray-125", foreground = "red")),
               "give both")
})

test_that("an empty or non-format argument is refused", {
  expect_error(cf(xl_format()), "is an empty format")
  expect_error(cf("red"), "must be an xl_format object")
  expect_null(.chart_format_payload(NULL, "format"))
})

# ── transparency reaches the format objects ───────────────────────────────────

test_that("transparency is carried by xl_fill and xl_border", {
  expect_equal(unclass(xl_fill(background = "red",
                               transparency = 40))$fill$transparency, 40L)
  expect_equal(unclass(xl_border(all = "thin",
                                 transparency = 25))$border$transparency, 25L)
  expect_error(xl_fill(background = "red", transparency = 101),
               "`transparency` must be between 0 and 100")
  expect_error(xl_border(all = "thin", transparency = -1),
               "`transparency` must be between 0 and 100")
  # NA means unset, as everywhere else
  expect_null(unclass(xl_fill(background = "red"))$fill$transparency)
})

test_that("transparency does not disturb a cell format", {
  # it is inert for cells -- Excel has no transparency there -- so it must not
  # change the styles a normal write produces
  df <- data.frame(a = 1:2)
  plain <- styles_string(list(D = xl_sheet(df,
    cols = xl_col_spec("a", format = xl_fill(background = "red")))))
  with_t <- styles_string(list(D = xl_sheet(df,
    cols = xl_col_spec("a", format = xl_fill(background = "red",
                                             transparency = 40)))))
  expect_equal(plain, with_t)
})

# ── What each chart part will accept ──────────────────────────────────────────

test_that("a series takes no font, and says where the font belongs", {
  # libxlsxwriter's series has a line, a fill and a pattern but no font: a
  # series is a shape, not text.  Dropping the font silently would leave the
  # caller with an unstyled chart and no reason why.
  expect_error(.chart_format_payload(xl_font(bold = TRUE), "format",
                                     accept = c("line", "fill", "pattern"),
                                     part = "a chart series"),
               "a shape and has no text")
  expect_error(.chart_format_payload(xl_font(bold = TRUE), "format",
                                     accept = c("line", "fill", "pattern"),
                                     part = "a chart series"),
               "xl_border() or xl_fill()", fixed = TRUE)
  # the rest of the same format still translates
  p <- .chart_format_payload(xl_fill(background = "red"), "format",
                             accept = c("line", "fill", "pattern"),
                             part = "a chart series")
  expect_equal(p$fill$color, xl_color("red"))
})

test_that("a title takes only a font, and says so for each other group", {
  ttl <- function(f) .chart_format_payload(f, "title_format", accept = "font",
                                           part = "a chart title")
  expect_equal(ttl(xl_font(size = 14))$font$size, 14)
  expect_error(ttl(xl_border(all = "thin")), "takes no line")
  expect_error(ttl(xl_fill(background = "red")), "takes no fill")
  expect_error(ttl(xl_fill(pattern = "gray-125", foreground = "red",
                           background = "white")), "takes no fill")
  expect_error(ttl(xl_border(all = "thin")), "Use xl_font()", fixed = TRUE)
})

test_that("an empty format names the groups the part could use", {
  expect_error(.chart_format_payload(xl_format(), "title_format",
                                     accept = "font", part = "a chart title"),
               "give xl_font()", fixed = TRUE)
})

test_that("black reaches a chart, rather than reading as no colour at all", {
  # every chart colour is truth-tested in libxlsxwriter -- `if (line->color)` --
  # so plain black, which xl_color() gives as 0, would be taken for "unset" and
  # the part drawn with Excel's default.  LXW_COLOR_BLACK is 0x1000000 for this
  # reason, and LXW_COLOR_MASK strips the bit again on the way out.
  expect_equal(.chart_color(xl_color("black")), 0x1000000L)
  expect_equal(.chart_color(xl_color("navy")), xl_color("navy"))
  expect_null(.chart_color(NULL))
  for (f in list(xl_border(all = "thin", color = "black"),
                 xl_fill(background = "black"),
                 xl_font(color = "black")))
    expect_true(any(vapply(cf(f), function(e) identical(e$color, 0x1000000L),
                           logical(1))))
  # a pattern's two colours go the same way
  p <- cf(xl_fill(pattern = "gray-125", foreground = "black",
                  background = "white"))$pattern
  expect_equal(p$fg_color, 0x1000000L)
})

test_that("a black chart line is written as a black line", {
  t <- tempfile(fileext = ".xlsx")
  df <- data.frame(a = c("x", "y"), b = c(1, 2), stringsAsFactors = FALSE)
  write_xlsx(list(Data = xl_sheet(df, chart = xl_chart("column",
    xl_chart_series(values = list(cols = "b")),
    plot_area_format = xl_border(all = "thin", color = "black")))), t)
  d <- tempfile(); dir.create(d); utils::unzip(t, exdir = d)
  x <- paste(readLines(list.files(file.path(d, "xl/charts"), pattern = "^chart",
                                  full.names = TRUE)[1L], warn = FALSE),
             collapse = "")
  plot_area <- substr(x, regexpr("<c:plotArea>", x), regexpr("</c:plotArea>", x))
  expect_match(plot_area, '<a:srgbClr val="000000"/>', fixed = TRUE)
})

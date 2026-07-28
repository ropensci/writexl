# Conditional formatting.  The rules land in <conditionalFormatting> in the
# worksheet part, and their formats in <dxfs> in styles.xml -- differential
# formats, not the ordinary cell styles.

cf <- function(..., df = data.frame(v = c(50, 150, 250))) {
  x <- xlsx_part(write_tmp(list(D = xl_sheet(df, conditional = list(...)))),
                 "xl/worksheets/sheet1.xml", raw = TRUE)
  m <- regmatches(x, regexpr("<conditionalFormatting.*</conditionalFormatting>", x))
  if (!length(m)) NA_character_ else m
}

cf_styles <- function(...) {
  p <- write_tmp(list(D = xl_sheet(data.frame(v = c(50, 150, 250)),
                                   conditional = list(...))))
  s <- xlsx_part(p, "xl/styles.xml", raw = TRUE)
  m <- regmatches(s, regexpr("<dxfs.*?</dxfs>", s))
  if (!length(m)) NA_character_ else m
}

# ── Simple cell rules ─────────────────────────────────────────────────────────

test_that("a comparison rule writes cellIs with an operator", {
  x <- cf(xl_cond_cell("A2:A4", criteria = ">", value = 100,
                       format = xl_fill(background = "red")))
  expect_match(x, 'type="cellIs"', fixed = TRUE)
  expect_match(x, 'operator="greaterThan"', fixed = TRUE)
  expect_match(x, "<formula>100</formula>", fixed = TRUE)
  expect_match(x, 'sqref="A2:A4"', fixed = TRUE)
})

test_that("each comparison criteria maps to its Excel operator", {
  ops <- c("==" = "equal", "!=" = "notEqual", ">" = "greaterThan",
           "<" = "lessThan", ">=" = "greaterThanOrEqual",
           "<=" = "lessThanOrEqual")
  for (k in names(ops))
    expect_match(cf(xl_cond_cell("A2:A4", criteria = k, value = 100)),
                 sprintf('operator="%s"', ops[[k]]), fixed = TRUE, label = k)
})

test_that("between uses both bounds", {
  x <- cf(xl_cond_cell("A2:A4", criteria = "between", min = 10, max = 20))
  expect_match(x, 'operator="between"', fixed = TRUE)
  expect_match(x, "<formula>10</formula>", fixed = TRUE)
  expect_match(x, "<formula>20</formula>", fixed = TRUE)
})

test_that("the type is inferred from the criteria", {
  # a text criteria implies a text rule without naming the type
  expect_match(cf(xl_cond_cell("A2:A4", criteria = "contains", value = "urgent")),
               'type="containsText"', fixed = TRUE)
  expect_match(cf(xl_cond_cell("A2:A4", criteria = "today")),
               'type="timePeriod"', fixed = TRUE)
  expect_match(cf(xl_cond_cell("A2:A4", criteria = "above")),
               'type="aboveAverage"', fixed = TRUE)
})

test_that("the rules that need no criteria are written", {
  for (ty in c("duplicate", "unique", "blanks", "no_blanks", "errors",
               "no_errors"))
    expect_match(cf(xl_cond_cell("A2:A4", type = ty)), "<cfRule", fixed = TRUE,
                 label = ty)
  expect_match(cf(xl_cond_cell("A2:A4", type = "duplicate")),
               'type="duplicateValues"', fixed = TRUE)
  expect_match(cf(xl_cond_cell("A2:A4", type = "unique")),
               'type="uniqueValues"', fixed = TRUE)
})

test_that("a formula rule carries its formula", {
  x <- cf(xl_cond_cell("A2:A4", type = "formula", value = "=$A2>100"))
  expect_match(x, 'type="expression"', fixed = TRUE)
  expect_match(x, "$A2&gt;100", fixed = TRUE)
})

test_that("top and bottom rules are written, with and without percent", {
  expect_match(cf(xl_cond_cell("A2:A4", type = "top", value = 10)),
               'type="top10"', fixed = TRUE)
  expect_match(cf(xl_cond_cell("A2:A4", criteria = "percent", value = 10, type = "top")),
               'percent="1"', fixed = TRUE)
  expect_match(cf(xl_cond_cell("A2:A4", type = "bottom", value = 5)),
               'bottom="1"', fixed = TRUE)
})

# ── Formats land in dxfs ──────────────────────────────────────────────────────

test_that("a rule's format becomes a differential format, not a cell style", {
  # this is the part worth pinning: conditional formats use <dxfs>, so the
  # shared format registry has to reach somewhere other than cellXfs
  d <- cf_styles(xl_cond_cell("A2:A4", criteria = ">", value = 100,
                              format = xl_fill(background = "red")))
  expect_match(d, "<dxf>", fixed = TRUE)
  expect_match(d, "FFFF0000", ignore.case = TRUE)
  expect_match(cf(xl_cond_cell("A2:A4", criteria = ">", value = 100,
                               format = xl_fill(background = "red"))),
               'dxfId="0"', fixed = TRUE)
})

test_that("several rules get distinct dxf ids and ascending priorities", {
  x <- cf(xl_cond_cell("A2:A4", criteria = ">", value = 100, format = xl_fill(background = "red")),
          xl_cond_cell("A2:A4", type = "duplicate",
                       format = xl_font(bold = TRUE)))
  expect_match(x, 'dxfId="0"', fixed = TRUE)
  expect_match(x, 'dxfId="1"', fixed = TRUE)
  expect_match(x, 'priority="1"', fixed = TRUE)
  expect_match(x, 'priority="2"', fixed = TRUE)
  expect_match(cf_styles(
    xl_cond_cell("A2:A4", criteria = ">", value = 100, format = xl_fill(background = "red")),
    xl_cond_cell("A2:A4", type = "duplicate", format = xl_font(bold = TRUE))),
    'count="2"', fixed = TRUE)
})

test_that("a rule with no format is still written", {
  expect_match(cf(xl_cond_cell("A2:A4", criteria = ">", value = 100)), "<cfRule", fixed = TRUE)
})

# ── The type / criteria pairing ───────────────────────────────────────────────

test_that("a criteria illegal for its type is refused, naming both", {
  # libxlsxwriter accepts these and Excel then ignores the rule, so the check
  # has to be ours
  expect_error(xl_cond_cell("A2:A4", criteria = "contains", value = "x", type = "cell"),
               'not valid for type = "cell"')
  expect_error(xl_cond_cell("A2:A4", criteria = ">", value = 1, type = "text"),
               'not valid for type = "text"')
  expect_error(xl_cond_cell("A2:A4", criteria = "today", type = "average"),
               'not valid for type = "average"')
  # the message lists what would be valid *for that type*
  expect_error(xl_cond_cell("A2:A4", criteria = "contains", value = "x", type = "cell"),
               "not between", fixed = TRUE)
  expect_error(xl_cond_cell("A2:A4", criteria = ">", value = 1, type = "text"),
               "begins with", fixed = TRUE)
})

test_that("criteria is refused for the rules that carry their own meaning", {
  for (ty in c("duplicate", "unique", "blanks", "errors", "formula"))
    expect_error(xl_cond_cell("A2:A4", criteria = ">", value = 1, type = ty),
                 "does not apply", label = ty)
})

test_that("criteria is required where it decides the rule", {
  for (ty in c("cell", "text", "time_period", "average"))
    expect_error(xl_cond_cell("A2:A4", type = ty, value = 1),
                 "`criteria` is required", label = ty)
})

test_that("between needs both bounds and comparisons need a value", {
  expect_error(xl_cond_cell("A2:A4", criteria = "between", min = 1),
               "needs both `min` and `max`")
  expect_error(xl_cond_cell("A2:A4", criteria = ">"), "needs a `value`")
  expect_error(xl_cond_cell("A2:A4", type = "formula"), "needs a `value`")
})

test_that("bad enum values and ranges are rejected", {
  expect_error(xl_cond_cell("A2:A4", criteria = "≈", value = 1), "criteria")
  expect_error(xl_cond_cell("A2:A4", criteria = ">", value = 1, type = "rainbow"), "type")
  expect_error(xl_cond_cell(), "must name the cells")
  expect_error(xl_cond_cell("A2:A4", criteria = ">", value = 1, format = "red"),
               "must be an xl_format")
})

# ── Wiring ────────────────────────────────────────────────────────────────────

test_that("a single rule need not be wrapped in a list", {
  x <- xlsx_part(write_tmp(list(D = xl_sheet(data.frame(v = c(1, 2, 3)),
    conditional = xl_cond_cell("A2:A4", criteria = ">", value = 1)))),
    "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_match(x, "<cfRule", fixed = TRUE)
})

test_that("non-conditional objects are rejected", {
  expect_error(write_tmp(list(D = xl_sheet(data.frame(v = 1),
                                           conditional = "A1:B2"))),
               "must be an xl_conditional object")
})

test_that("stop_if_true and multi_range reach the file", {
  expect_match(cf(xl_cond_cell("A2:A4", criteria = ">", value = 100, stop_if_true = TRUE)),
               'stopIfTrue="1"', fixed = TRUE)
  expect_match(cf(xl_cond_cell("A2:A4", criteria = ">", value = 100, multi_range = "A2:A4 C2:C4")),
               'sqref="A2:A4 C2:C4"', fixed = TRUE)
})

test_that("a rule accepts a data-frame-relative spec", {
  df <- data.frame(v = c(1, 2, 3))
  x <- cf(xl_cond_cell(list(rows = 1:3, cols = "v"), criteria = ">", value = 1), df = df)
  expect_match(x, 'sqref="A2:A4"', fixed = TRUE)
})

test_that("conditional formatting leaves the cells alone and keeps streaming", {
  skip_if_not_installed("readxl")
  df <- data.frame(v = c(50, 150, 250))
  p <- write_tmp(list(D = xl_sheet(df,
    conditional = xl_cond_cell("A2:A4", criteria = ">", value = 100,
                               format = xl_fill(background = "red")))))
  expect_equal(as.data.frame(readxl::read_xlsx(p)), df)
  # written at packaging time, so row streaming stays on
  expect_true(is.null(xlsx_part(p, "xl/sharedStrings.xml")))
})

test_that("the print method runs", {
  expect_output(print(xl_cond_cell("A2:A4", criteria = ">", value = 1)), "xl_conditional")
})

# ── Colour scales ─────────────────────────────────────────────────────────────

test_that("a three-colour scale writes three stops and three colours", {
  x <- cf(xl_cond_scale("A2:A4", colors = c("red", "yellow", "green")))
  expect_match(x, 'type="colorScale"', fixed = TRUE)
  expect_match(x, "FFFF0000", ignore.case = TRUE)   # red
  expect_match(x, "FFFFFF00", ignore.case = TRUE)   # yellow
  expect_match(x, "FF00FF00", ignore.case = TRUE)   # green
  expect_equal(lengths(regmatches(x, gregexpr("<cfvo ", x))), 3L)
})

test_that("a two-colour scale writes two stops", {
  x <- cf(xl_cond_scale("A2:A4", colors = c("white", "blue")))
  expect_match(x, 'type="colorScale"', fixed = TRUE)
  expect_equal(lengths(regmatches(x, gregexpr("<cfvo ", x))), 2L)
})

test_that("scale stops accept values and rule types", {
  x <- cf(xl_cond_scale("A2:A4", colors = c("red", "green"),
                        values = c(10, 90),
                        rule_types = c("number", "number")))
  expect_match(x, 'type="num"', fixed = TRUE)
  expect_match(x, 'val="10"', fixed = TRUE)
  expect_match(x, 'val="90"', fixed = TRUE)
})

test_that("a scale needs two or three colours", {
  expect_error(xl_cond_scale("A2:A4", colors = "red"), "2 or 3 colours")
  expect_error(xl_cond_scale("A2:A4", colors = c("a", "b", "c", "d")),
               "2 or 3 colours")
  expect_error(xl_cond_scale("A2:A4", colors = c("red", "green"),
                             values = 1), "one entry per colour")
  expect_error(xl_cond_scale("A2:A4", colors = c("red", "green"),
                             rule_types = "number"), "one entry per colour")
  expect_error(xl_cond_scale(), "must name the cells")
})

# ── Data bars ─────────────────────────────────────────────────────────────────

test_that("a data bar is written with its colour", {
  x <- cf(xl_cond_bar("A2:A4", color = "steelblue"))
  expect_match(x, 'type="dataBar"', fixed = TRUE)
  expect_match(x, "FF4682B4", ignore.case = TRUE)
})

test_that("bar options reach the file", {
  expect_match(cf(xl_cond_bar("A2:A4", color = "green", bar_only = TRUE)),
               "dataBar", fixed = TRUE)
  # the 2010 extensions carry the solid/direction/axis settings
  x <- cf(xl_cond_bar("A2:A4", color = "green", solid = TRUE,
                      direction = "right to left", axis = "midpoint"))
  expect_match(x, "x14:", fixed = TRUE)
})

test_that("bar ends accept values and rule types", {
  x <- cf(xl_cond_bar("A2:A4", color = "green", values = c(0, 100),
                      rule_types = c("number", "number")))
  expect_match(x, 'type="num"', fixed = TRUE)
})

test_that("bar arguments are validated", {
  expect_error(xl_cond_bar("A2:A4", values = 1), "two ends")
  expect_error(xl_cond_bar("A2:A4", rule_types = "number"), "two ends")
  expect_error(xl_cond_bar("A2:A4", direction = "sideways"), "direction")
  expect_error(xl_cond_bar("A2:A4", axis = "diagonal"), "axis")
  expect_error(xl_cond_bar(), "must name the cells")
})

# ── Icon sets ─────────────────────────────────────────────────────────────────

test_that("an icon set is written", {
  expect_match(cf(xl_cond_icons("A2:A4", style = "3_traffic_lights")),
               'type="iconSet"', fixed = TRUE)
})

test_that("every icon style maps to its Excel name", {
  # the style is an index into libxlsxwriter's enum, and three of the
  # conditional type constants lack the TYPE_ prefix -- getting the offset
  # wrong would silently write a colour scale instead, so enumerate them all
  expected <- c(
    "3_arrows" = "3Arrows", "3_arrows_gray" = "3ArrowsGray",
    "3_flags" = "3Flags", "3_traffic_lights" = NA,   # the default: no attribute
    "3_traffic_lights_rimmed" = "3TrafficLights2", "3_signs" = "3Signs",
    "3_symbols_circled" = "3Symbols", "3_symbols" = "3Symbols2",
    "4_arrows" = "4Arrows", "4_arrows_gray" = "4ArrowsGray",
    "4_red_to_black" = "4RedToBlack", "4_ratings" = "4Rating",
    "4_traffic_lights" = "4TrafficLights", "5_arrows" = "5Arrows",
    "5_arrows_gray" = "5ArrowsGray", "5_ratings" = "5Rating",
    "5_quarters" = "5Quarters")
  for (st in names(expected)) {
    x <- cf(xl_cond_icons("A2:A4", style = st))
    if (is.na(expected[[st]]))
      expect_match(x, "<iconSet>", fixed = TRUE, label = st)
    else
      expect_match(x, sprintf('iconSet="%s"', expected[[st]]), fixed = TRUE,
                   label = st)
  }
})

test_that("all 17 icon styles are covered by the map", {
  # a newly added style in a libxlsxwriter update should fail here rather than
  # be silently unreachable from R
  expect_equal(length(.LXW_COND_ICONS), 17L)
})

test_that("icon options are written", {
  expect_match(cf(xl_cond_icons("A2:A4", icons_only = TRUE)),
               'showValue="0"', fixed = TRUE)
  expect_match(cf(xl_cond_icons("A2:A4", reverse = TRUE)),
               'reverse="1"', fixed = TRUE)
})

test_that("icon arguments are validated", {
  expect_error(xl_cond_icons("A2:A4", style = "6_smileys"), "style")
  expect_error(xl_cond_icons(), "must name the cells")
})

test_that("icon sets embed no image parts", {
  # Excel draws these itself, so nothing goes into xl/media
  p <- write_tmp(list(D = xl_sheet(data.frame(v = c(10, 50, 90)),
                                   conditional = xl_cond_icons("A2:A4"))))
  d <- tempfile(); dir.create(d); utils::unzip(p, exdir = d)
  expect_false(dir.exists(file.path(d, "xl", "media")))
})

# ── The four clusters together ────────────────────────────────────────────────

test_that("all four kinds coexist on one sheet", {
  skip_if_not_installed("readxl")
  df <- data.frame(a = c(10, 50, 90), b = c(10, 50, 90),
                   cc = c(10, 50, 90), d = c(10, 50, 90))
  p <- write_tmp(list(D = xl_sheet(df, conditional = list(
    xl_cond_cell("A2:A4", criteria = ">", value = 20,
                 format = xl_fill(background = "red")),
    xl_cond_scale("B2:B4"),
    xl_cond_bar("C2:C4", color = "steelblue"),
    xl_cond_icons("D2:D4")))))
  x <- xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE)
  for (ty in c("cellIs", "colorScale", "dataBar", "iconSet"))
    expect_match(x, sprintf('type="%s"', ty), fixed = TRUE, label = ty)
  expect_equal(as.data.frame(readxl::read_xlsx(p)), df)
})

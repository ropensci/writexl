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

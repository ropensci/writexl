# Data validation.  libxlsxwriter's 17 validation types are collapsed to five
# kinds plus list/custom/any, with the formula and datetime variants inferred
# from the R type of the limit.

dv <- function(..., df = data.frame(qty = 1:3)) {
  x <- xlsx_part(write_tmp(list(D = xl_sheet(df, validation = list(...)))),
                 "xl/worksheets/sheet1.xml", raw = TRUE)
  m <- regmatches(x, regexpr("<dataValidations.*</dataValidations>", x))
  if (!length(m)) NA_character_ else m
}

# ── Numeric and length kinds ──────────────────────────────────────────────────

test_that("an integer bound writes a whole-number validation", {
  x <- dv(xl_validation("A2:A4", type = "integer", min = 1, max = 10))
  expect_match(x, 'type="whole"', fixed = TRUE)
  expect_match(x, "<formula1>1</formula1>", fixed = TRUE)
  expect_match(x, "<formula2>10</formula2>", fixed = TRUE)
  expect_match(x, 'sqref="A2:A4"', fixed = TRUE)
})

test_that("min and max together imply between", {
  # criteria may be omitted when both bounds are given
  a <- dv(xl_validation("A2:A4", type = "integer", min = 1, max = 10))
  b <- dv(xl_validation("A2:A4", type = "integer", criteria = "between",
                        min = 1, max = 10))
  expect_equal(a, b)
})

test_that("the single-limit criteria are written with an operator", {
  x <- dv(xl_validation("A2:A4", type = "integer", criteria = ">", value = 5))
  expect_match(x, 'operator="greaterThan"', fixed = TRUE)
  expect_match(x, "<formula1>5</formula1>", fixed = TRUE)
  expect_match(dv(xl_validation("A2:A4", type = "decimal", criteria = "<=",
                                value = 2.5)),
               'operator="lessThanOrEqual"', fixed = TRUE)
})

test_that("decimal and length kinds are distinct types", {
  expect_match(dv(xl_validation("A2:A4", type = "decimal", criteria = ">",
                                value = 0)),
               'type="decimal"', fixed = TRUE)
  expect_match(dv(xl_validation("A2:A4", type = "length", criteria = "<=",
                                value = 10)),
               'type="textLength"', fixed = TRUE)
})

test_that("a formula limit selects the formula variant", {
  # the type is inferred from the limit being a "=" string
  x <- dv(xl_validation("A2:A4", type = "integer", criteria = ">",
                        value = "=$B$1"))
  expect_match(x, 'type="whole"', fixed = TRUE)
  expect_match(x, "<formula1>$B$1</formula1>", fixed = TRUE)
})

# ── Dates and times ───────────────────────────────────────────────────────────

test_that("a Date bound writes a date validation with a serial number", {
  x <- dv(xl_validation("A2:A4", type = "date", criteria = ">=",
                        value = as.Date("2024-01-01")))
  expect_match(x, 'type="date"', fixed = TRUE)
  # 2024-01-01 is Excel serial 45292
  expect_match(x, "<formula1>45292</formula1>", fixed = TRUE)
})

test_that("a POSIXct bound carries the time of day", {
  x <- dv(xl_validation("A2:A4", type = "date", criteria = ">=",
                        value = as.POSIXct("2024-01-01 12:00:00", tz = "UTC")))
  expect_match(x, 'type="date"', fixed = TRUE)
  expect_match(x, "45292.5", fixed = TRUE)   # midday
})

test_that("a date range uses both bounds", {
  x <- dv(xl_validation("A2:A4", type = "date",
                        min = as.Date("2024-01-01"),
                        max = as.Date("2024-12-31")))
  expect_match(x, "<formula1>45292</formula1>", fixed = TRUE)
  expect_match(x, "<formula2>45657</formula2>", fixed = TRUE)
})

# ── Dropdown lists ────────────────────────────────────────────────────────────

test_that("a list becomes a quoted CSV formula", {
  x <- dv(xl_validation("A2:A4", list = c("open", "high", "close")))
  expect_match(x, 'type="list"', fixed = TRUE)
  expect_match(x, '<formula1>"open,high,close"</formula1>', fixed = TRUE)
})

test_that("a list given as a range formula is used as-is", {
  x <- dv(xl_validation("A2:A4", list = "=$H$1:$H$5"))
  expect_match(x, 'type="list"', fixed = TRUE)
  expect_match(x, "<formula1>$H$1:$H$5</formula1>", fixed = TRUE)
})

test_that("a single-choice list is allowed", {
  expect_match(dv(xl_validation("A2:A4", list = "only")),
               '<formula1>"only"</formula1>', fixed = TRUE)
})

test_that("the dropdown arrow can be turned off", {
  expect_match(dv(xl_validation("A2:A4", list = c("a", "b"), dropdown = FALSE)),
               'showDropDown="1"', fixed = TRUE)   # Excel stores it inverted
})

test_that("a list whose joined length exceeds 255 characters is refused", {
  # Excel's limit is on the comma-joined string, not the number of choices
  long <- rep(strrep("x", 20), 20)          # 20 * 21 - 1 = 419 characters
  expect_error(xl_validation("A2:A4", list = long), "joined by commas")
  expect_error(xl_validation("A2:A4", list = long), "255")
  # and just under the limit is fine
  ok <- rep(strrep("x", 20), 12)            # 251 characters
  expect_s3_class(xl_validation("A2:A4", list = ok), "xl_validation")
})

# ── custom and any ────────────────────────────────────────────────────────────

test_that("a custom formula validation is written", {
  x <- dv(xl_validation("A2:A4", type = "custom", value = "=A2>A1"))
  expect_match(x, 'type="custom"', fixed = TRUE)
  expect_match(x, "<formula1>A2&gt;A1</formula1>", fixed = TRUE)
})

test_that("type = any carries only the messages", {
  x <- dv(xl_validation("A2:A4", type = "any", input_title = "Hint",
                        input_message = "Anything goes"))
  expect_match(x, 'promptTitle="Hint"', fixed = TRUE)
  expect_match(x, 'prompt="Anything goes"', fixed = TRUE)
})

# ── Messages and toggles ──────────────────────────────────────────────────────

test_that("input and error text reach the file", {
  x <- dv(xl_validation("A2:A4", type = "integer", criteria = ">", value = 0,
                        input_title = "Quantity", input_message = "A number",
                        error_title = "Bad", error_message = "Must exceed 0"))
  expect_match(x, 'promptTitle="Quantity"', fixed = TRUE)
  expect_match(x, 'prompt="A number"', fixed = TRUE)
  expect_match(x, 'errorTitle="Bad"', fixed = TRUE)
  expect_match(x, 'error="Must exceed 0"', fixed = TRUE)
})

test_that("error_type selects the alert style", {
  base <- xl_validation("A2:A4", type = "integer", criteria = ">", value = 0)
  expect_false(grepl("errorStyle", dv(base), fixed = TRUE))   # stop is default
  expect_match(dv(xl_validation("A2:A4", type = "integer", criteria = ">",
                                value = 0, error_type = "warning")),
               'errorStyle="warning"', fixed = TRUE)
  expect_match(dv(xl_validation("A2:A4", type = "integer", criteria = ">",
                                value = 0, error_type = "information")),
               'errorStyle="information"', fixed = TRUE)
})

test_that("the on-by-default toggles are only written when turned off", {
  on <- dv(xl_validation("A2:A4", type = "integer", criteria = ">", value = 0))
  expect_match(on, 'allowBlank="1"', fixed = TRUE)
  off <- dv(xl_validation("A2:A4", type = "integer", criteria = ">", value = 0,
                          ignore_blank = FALSE, show_input = FALSE,
                          show_error = FALSE))
  expect_false(grepl('allowBlank="1"', off, fixed = TRUE))
  expect_false(grepl('showInputMessage="1"', off, fixed = TRUE))
})

# ── Validation of the arguments themselves ────────────────────────────────────

test_that("criteria is required for the comparison kinds", {
  for (k in c("integer", "decimal", "date", "time", "length"))
    expect_error(xl_validation("A2:A4", type = k, value = 1),
                 "`criteria` is required", label = k)
})

test_that("criteria is refused for the kinds that carry their own meaning", {
  expect_error(xl_validation("A2:A4", list = c("a", "b"), criteria = ">"),
               "does not apply")
  expect_error(xl_validation("A2:A4", type = "custom", value = "=A1",
                             criteria = ">"), "does not apply")
  expect_error(xl_validation("A2:A4", type = "any", criteria = ">"),
               "does not apply")
})

test_that("between needs both bounds and the others need a value", {
  expect_error(xl_validation("A2:A4", type = "integer", criteria = "between",
                             min = 1), "needs both `min` and `max`")
  expect_error(xl_validation("A2:A4", type = "integer",
                             criteria = "not between", max = 9),
               "needs both `min` and `max`")
  expect_error(xl_validation("A2:A4", type = "integer", criteria = ">"),
               "needs a `value`")
})

test_that("custom needs its formula and list needs its choices", {
  expect_error(xl_validation("A2:A4", type = "custom"), "needs a `value`")
  expect_error(xl_validation("A2:A4", list = character(0)),
               "at least one choice")
  expect_error(xl_validation("A2:A4", list = c("a", NA)), "must not be NA")
  expect_error(xl_validation("A2:A4", list = 1:3), "character vector")
})

test_that("titles and messages are capped at Excel's limits", {
  expect_error(xl_validation("A2:A4", type = "any",
                             input_title = strrep("x", 33)),
               "at most 32 characters")
  expect_error(xl_validation("A2:A4", type = "any",
                             error_title = strrep("x", 33)),
               "at most 32 characters")
  expect_error(xl_validation("A2:A4", type = "any",
                             input_message = strrep("x", 256)),
               "at most 255 characters")
  # exactly at the limit is fine
  expect_s3_class(xl_validation("A2:A4", type = "any",
                                input_title = strrep("x", 32)),
                  "xl_validation")
})

test_that("the limits count characters, not bytes", {
  s <- strrep("é", 32)          # 64 bytes, 32 characters
  expect_gt(nchar(s, type = "bytes"), 32L)
  expect_s3_class(xl_validation("A2:A4", type = "any", input_title = s),
                  "xl_validation")
})

test_that("bad types and ranges are rejected", {
  expect_error(xl_validation("A2:A4", type = "colour"), "type")
  expect_error(xl_validation(), "must name the cells")
  expect_error(write_tmp(list(D = xl_sheet(data.frame(a = 1),
                                           validation = "A1:B2"))),
               "must be an xl_validation object")
})

# ── Wiring ────────────────────────────────────────────────────────────────────

test_that("a single validation need not be wrapped in a list", {
  x <- xlsx_part(write_tmp(list(D = xl_sheet(data.frame(qty = 1:3),
    validation = xl_validation("A2:A4", type = "integer", criteria = ">",
                               value = 0)))),
    "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_match(x, 'type="whole"', fixed = TRUE)
})

test_that("several validations coexist", {
  x <- dv(xl_validation("A2:A4", type = "integer", min = 1, max = 10),
          xl_validation("B2:B4", list = c("a", "b")))
  expect_match(x, 'count="2"', fixed = TRUE)
})

test_that("a validation accepts a data-frame-relative spec", {
  df <- data.frame(a = 1:3, b = 1:3)
  x <- dv(xl_validation(list(rows = 1:3, cols = "a"), type = "integer",
                        criteria = ">", value = 0), df = df)
  # data rows 1:3 are sheet rows 2:4 once the header is counted
  expect_match(x, 'sqref="A2:A4"', fixed = TRUE)
})

test_that("validation does not disturb the cells or need streaming off", {
  skip_if_not_installed("readxl")
  df <- data.frame(qty = 1:3)
  p <- write_tmp(list(D = xl_sheet(df,
    validation = xl_validation("A2:A4", type = "integer", min = 1, max = 10))))
  expect_equal(as.data.frame(readxl::read_xlsx(p)), df)
  # validations are written at packaging time, so row streaming stays on
  expect_true(is.null(xlsx_part(p, "xl/sharedStrings.xml")))
})

test_that("the print method runs", {
  expect_output(print(xl_validation("A2:A4", list = c("a", "b"))),
                "xl_validation")
})

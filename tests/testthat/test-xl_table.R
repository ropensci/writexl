# Worksheet tables.  This file covers the constructors and their validation;
# what reaches the file is tested once the C side lands.

# ── Style parsing ─────────────────────────────────────────────────────────────

test_that("a style string parses to the type and number C wants", {
  expect_equal(.parse_table_style("medium 9"), list(type = 2L, number = 9L))
  expect_equal(.parse_table_style("dark 11"), list(type = 3L, number = 11L))
  expect_equal(.parse_table_style("light 21"), list(type = 1L, number = 21L))
  # "none" is Light 0 in libxlsxwriter's numbering, not a fourth type
  expect_equal(.parse_table_style("none"), list(type = 1L, number = 0L))
  expect_equal(.parse_table_style("light 0"), .parse_table_style("none"))
  # case and spacing are forgiving
  expect_equal(.parse_table_style("  MEDIUM   9 "), list(type = 2L, number = 9L))
  # the default matches Excel's own default table style
  expect_equal(.parse_table_style(NULL), list(type = 2L, number = 9L))
})

test_that("each style type enforces its own number range", {
  # libxlsxwriter only warns and silently substitutes Medium 9, so writexl has
  # to be the one that refuses -- and the ranges genuinely differ per type
  expect_error(.parse_table_style("light 22"), "light styles are numbered 0 to 21")
  expect_error(.parse_table_style("medium 0"), "medium styles are numbered 1 to 28")
  expect_error(.parse_table_style("medium 29"), "medium styles are numbered 1 to 28")
  expect_error(.parse_table_style("dark 0"), "dark styles are numbered 1 to 11")
  expect_error(.parse_table_style("dark 12"), "dark styles are numbered 1 to 11")
  # the boundaries themselves are accepted
  expect_silent(.parse_table_style("light 0"))
  expect_silent(.parse_table_style("light 21"))
  expect_silent(.parse_table_style("medium 1"))
  expect_silent(.parse_table_style("medium 28"))
  expect_silent(.parse_table_style("dark 1"))
  expect_silent(.parse_table_style("dark 11"))
})

test_that("a malformed style is rejected", {
  expect_error(.parse_table_style("rainbow 3"), 'must be "light", "medium" or "dark"')
  expect_error(.parse_table_style("medium"), 'a type and number like "medium 9"')
  expect_error(.parse_table_style("medium 9 extra"), 'a type and number like')
  expect_error(.parse_table_style("medium nine"), "numbered 1 to 28")
  expect_error(.parse_table_style(9), "must be a single string")
  expect_error(.parse_table_style(NA_character_), "must be a single string")
})

# ── Table names ───────────────────────────────────────────────────────────────

test_that("Excel's table-name rules are enforced", {
  expect_error(.check_table_name("My Table"), "spaces or punctuation")
  expect_error(.check_table_name("a-b"), "spaces or punctuation")
  expect_error(.check_table_name("2024sales"), "may not start with a digit")
  expect_error(.check_table_name("C"), 'may not be "C" or "R"')
  expect_error(.check_table_name("r"), 'may not be "C" or "R"')
  expect_error(.check_table_name(""), "non-empty string")
  expect_error(.check_table_name(strrep("a", 256L)), "limit of 255")
  # libxlsxwriter does not reject this one, but Excel does
  expect_error(.check_table_name("A1"), "looks like a cell reference")
  expect_error(.check_table_name("XFD1048576"), "looks like a cell reference")
})

test_that("valid table names pass", {
  for (nm in c("Sales", "Sales_2024", "T.1", "a", "Table1x", strrep("a", 255L)))
    expect_silent(.check_table_name(nm))
  # a letter-and-digit name that is not a cell reference is fine
  expect_silent(.check_table_name("Q1data"))
})

test_that("the punctuation set matches libxlsxwriter's strpbrk set", {
  # enumerated rather than sampled: every character worksheet.c rejects must be
  # rejected here, or a name would reach the file and prompt a repair
  bad <- strsplit(" !\"#$%&'()*+,-/:;<=>?@[\\]^`{|}~", "", fixed = TRUE)[[1L]]
  for (ch in bad)
    expect_error(.check_table_name(paste0("a", ch, "b")),
                 "spaces or punctuation", label = ch)
})

# ── Constructors ──────────────────────────────────────────────────────────────

test_that("xl_table carries its options", {
  t <- xl_table(name = "Sales", style = "light 9", total_row = TRUE,
                banded_columns = TRUE, first_column = TRUE)
  p <- unclass(t)
  expect_s3_class(t, "xl_table")
  expect_equal(p$name, "Sales")
  expect_equal(p$style_type, 1L)
  expect_equal(p$style_number, 9L)
  expect_true(p$total_row)
  expect_true(p$banded_columns)
  expect_true(p$first_column)
  expect_false(p$last_column)
})

test_that("turning off the header row also turns off the autofilter", {
  # libxlsxwriter forces this itself; the object must not claim otherwise
  expect_false(unclass(xl_table(header_row = FALSE))$autofilter)
  expect_false(unclass(xl_table(header_row = FALSE, autofilter = TRUE))$autofilter)
  expect_true(unclass(xl_table())$autofilter)
})

test_that("a total without a total row is refused, naming the column", {
  expect_error(xl_table(columns = xl_table_column("qty", total = "sum")),
               'column "qty" gives a total')
  expect_error(xl_table(columns = xl_table_column("qty", total_label = "Total")),
               "no total row")
  expect_silent(xl_table(total_row = TRUE,
                         columns = xl_table_column("qty", total = "sum")))
})

test_that("name = NA with a column formula warns", {
  # structured references name the table, so leaving Excel to number it is a
  # real hazard -- but an allowed one, so it warns rather than errors
  expect_warning(xl_table(name = NA,
                          columns = xl_table_column("q", formula = "=1")),
                 "refers to the table by name")
  # without a formula there is nothing to break
  expect_silent(xl_table(name = NA))
  # and a real name is never a hazard
  expect_silent(xl_table(name = "Sales",
                         columns = xl_table_column("q", formula = "=1")))
})

test_that("xl_table_column carries its options and validates them", {
  cc <- unclass(xl_table_column("qty", header = "Quantity", total = "sum",
                                format = xl_font(bold = TRUE)))
  expect_equal(cc$col, "qty")
  expect_equal(cc$header, "Quantity")
  expect_equal(cc$total, "sum")
  expect_true(is_xl_format(cc$format))

  expect_error(xl_table_column(), "must name or index")
  expect_error(xl_table_column("q", formula = "SUM(1)"), "must start with '='")
  expect_error(xl_table_column("q", total = "median"), "`total` must be one of")
  expect_error(xl_table_column("q", total = "sum", total_label = "T"),
               "either `total` or `total_label`, not both")
  expect_error(xl_table_column("q", header = 1), "single non-NA string")
  expect_error(xl_table_column("q", format = "bold"), "must be an xl_format")
  expect_error(xl_table_column("q", header_format = "bold"),
               "`header_format` must be an xl_format")
})

test_that("every documented total function is accepted", {
  # the docs quantify over the list, so enumerate it
  for (f in names(.LXW_TABLE_FUNCTION))
    expect_silent(xl_table_column("q", total = f))
  expect_equal(length(.LXW_TABLE_FUNCTION), 8L)
  # the enum is sparse, which is the reason callers get names instead
  expect_equal(unname(.LXW_TABLE_FUNCTION[["std_dev"]]), 107L)
  expect_equal(unname(.LXW_TABLE_FUNCTION[["sum"]]), 109L)
})

test_that("columns normalise from one or a list", {
  one <- xl_table(columns = xl_table_column("a"))
  many <- xl_table(columns = list(xl_table_column("a"), xl_table_column("b")))
  expect_length(unclass(one)$columns, 1L)
  expect_length(unclass(many)$columns, 2L)
  expect_null(unclass(xl_table())$columns)
  expect_error(xl_table(columns = "a"), "must be an xl_table_column object")
  expect_error(xl_table(columns = list(xl_table_column("a"), "b")),
               "`columns\\[\\[2\\]\\]` must be an xl_table_column object")
})

test_that("flag arguments are validated", {
  for (arg in c("header_row", "autofilter", "banded_rows", "banded_columns",
                "first_column", "last_column", "total_row")) {
    args <- stats::setNames(list(NA), arg)
    expect_error(do.call(xl_table, args), sprintf("`%s` must be TRUE or FALSE", arg),
                 label = arg)
  }
  expect_error(xl_table(name = c("a", "b")), "single string, NULL or NA")
})

test_that("the print methods run", {
  expect_output(print(xl_table(name = "Sales")), "Sales")
  expect_output(print(xl_table()), "unnamed")
  expect_output(print(xl_table(name = NA)), "named by Excel")
  expect_output(print(xl_table(total_row = TRUE)), "total row")
  expect_output(print(xl_table_column("qty", total = "sum")), "total=sum")
  expect_output(print(xl_table_column("qty", formula = "=1")), "formula")
})

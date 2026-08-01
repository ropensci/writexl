# Worksheet tables: the constructors, their validation, and what reaches the
# file.

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

# ── What reaches the file ─────────────────────────────────────────────────────

sdf <- data.frame(fruit = c("apple", "banana", "cherry"), qty = c(5, 150, 300),
                  stringsAsFactors = FALSE)

# The table part for a sheet built with `...` passed to xl_sheet().
table_xml <- function(..., df = sdf) {
  p <- write_tmp(list(Sales = xl_sheet(df, ...)))
  xlsx_part(p, "xl/tables/table1.xml", raw = TRUE)
}

# The text a header cell actually shows, resolved through the shared strings.
header_text <- function(..., df = sdf) {
  p <- write_tmp(list(Sales = xl_sheet(df, ...)))
  w <- xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE)
  ss <- xlsx_part(p, "xl/sharedStrings.xml", raw = TRUE)
  strings <- regmatches(ss, gregexpr("<t>[^<]*</t>", ss))[[1L]]
  strings <- sub("</t>$", "", sub("^<t>", "", strings))
  row1 <- regmatches(w, regexpr('<row r="1".*?</row>', w))
  idx <- as.integer(sub(".*<v>([0-9]+)</v>.*", "\\1",
                        regmatches(row1, gregexpr("<v>[0-9]+</v>", row1))[[1L]]))
  strings[idx + 1L]
}

test_that("a table's columns take the data frame's column names", {
  # the whole point of the commit: _write_table_column_data() would otherwise
  # default these to "Column 1", "Column 2", and Excel rejects a file whose
  # table definition and header cells disagree
  x <- table_xml(table = xl_table())
  expect_match(x, '<tableColumn id="1" name="fruit"/>', fixed = TRUE)
  expect_match(x, '<tableColumn id="2" name="qty"/>', fixed = TRUE)
  expect_no_match(x, "Column 1", fixed = TRUE)
  # and the header cells say the same thing, which is what Excel checks
  expect_equal(header_text(table = xl_table()), c("fruit", "qty"))
})

test_that("an overridden header reaches both the definition and the cell", {
  x <- table_xml(table = xl_table(columns = xl_table_column("qty",
                                                            header = "Quantity")))
  expect_match(x, 'name="Quantity"', fixed = TRUE)
  expect_equal(header_text(table = xl_table(
    columns = xl_table_column("qty", header = "Quantity"))),
    c("fruit", "Quantity"))
})

test_that("the range defaults to the sheet's used range", {
  expect_match(table_xml(table = xl_table()), 'ref="A1:B4"', fixed = TRUE)
  # a total row extends it by one, and the autofilter stays on the data only
  x <- table_xml(table = xl_table(total_row = TRUE))
  expect_match(x, 'ref="A1:B5"', fixed = TRUE)
  expect_match(x, '<autoFilter ref="A1:B4"/>', fixed = TRUE)
  expect_match(x, 'totalsRowCount="1"', fixed = TRUE)
})

test_that("total functions and labels are written", {
  x <- table_xml(table = xl_table(total_row = TRUE, columns = list(
    xl_table_column("fruit", total_label = "Total"),
    xl_table_column("qty", total = "sum"))))
  expect_match(x, 'totalsRowLabel="Total"', fixed = TRUE)
  expect_match(x, 'totalsRowFunction="sum"', fixed = TRUE)
})

test_that("a column carrying only a total label does not confuse the total", {
  # regression: `$total` partial-matches `total_label` once NULLs are dropped,
  # so a label-only column looked up the label in the function table and errored
  x <- table_xml(table = xl_table(total_row = TRUE,
                                  columns = xl_table_column("fruit",
                                                            total_label = "Total")))
  expect_match(x, 'totalsRowLabel="Total"', fixed = TRUE)
  expect_no_match(x, "totalsRowFunction", fixed = TRUE)
})

test_that("style, banding and highlight flags are written", {
  expect_match(table_xml(table = xl_table(style = "dark 3")),
               'name="TableStyleDark3"', fixed = TRUE)
  # "none" (Light 0) is written as a tableStyleInfo with no name at all, not
  # as a style called TableStyleLight0
  x0 <- table_xml(table = xl_table(style = "none"))
  expect_match(x0, "<tableStyleInfo showFirstColumn", fixed = TRUE)
  expect_no_match(x0, "TableStyle", fixed = TRUE)
  x <- table_xml(table = xl_table(first_column = TRUE, banded_columns = TRUE,
                                  banded_rows = FALSE))
  expect_match(x, 'showFirstColumn="1"', fixed = TRUE)
  expect_match(x, 'showColumnStripes="1"', fixed = TRUE)
  expect_match(x, 'showRowStripes="0"', fixed = TRUE)
})

test_that("turning off the header row removes the autofilter too", {
  x <- table_xml(table = xl_table(header_row = FALSE))
  expect_match(x, 'headerRowCount="0"', fixed = TRUE)
  expect_no_match(x, "<autoFilter", fixed = TRUE)
  # and the autofilter can be dropped on its own
  expect_no_match(table_xml(table = xl_table(autofilter = FALSE)),
                  "<autoFilter", fixed = TRUE)
})

test_that("a column formula is written as a calculated column", {
  d <- sdf
  d$double <- NA_real_
  x <- table_xml(df = d, table = xl_table(
    name = "Sales",
    columns = xl_table_column("double",
                              formula = "=Sales[[#This Row],[qty]]*2")))
  expect_match(x, "<calculatedColumnFormula>", fixed = TRUE)
  expect_match(x, "Sales[[#This Row],[qty]]*2", fixed = TRUE)
})

test_that("a column format reaches the data cells", {
  # libxlsxwriter applies a table column's format only to cells it writes
  # itself -- the total cell and a calculated column -- never to data already
  # written, so left to it this argument would silently do nothing
  p <- write_tmp(list(Sales = xl_sheet(sdf, table = xl_table(
    columns = xl_table_column("qty", format = xl_num_format("$#,##0.00"))))))
  s <- xlsx_part(p, "xl/styles.xml", raw = TRUE)
  w <- xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_match(s, "$#,##0.00", fixed = TRUE)
  # the column carries the style, and so does the data cell
  expect_match(w, '<col min="2" max="2"', fixed = TRUE)
  expect_match(regmatches(w, regexpr('<row r="2".*?</row>', w)),
               '<c r="B2" s=', fixed = TRUE)
})

test_that("a column format reaches both the column plan and the table part", {
  # Upstream (jmcnamara/libxlsxwriter#520) confirms the column plan is the
  # supported way to format a table's data cells, and adds that the table
  # column's own format is still needed "for strict correctness with Excel",
  # where it becomes dataDxfId.  Neither half is redundant; this pins both, so
  # a later tidy-up cannot drop the one that looks superfluous.
  p <- write_tmp(list(Sales = xl_sheet(sdf, table = xl_table(
    columns = xl_table_column("qty", format = xl_num_format("$#,##0.00"))))))
  w <- xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE)
  tbl <- xlsx_part(p, "xl/tables/table1.xml", raw = TRUE)
  expect_match(w, '<col min="2" max="2"', fixed = TRUE)      # the column plan
  expect_match(tbl, "dataDxfId=", fixed = TRUE)              # strict correctness
})

test_that("a column format merges over an xl_col_spec format", {
  # the table column is the more specific of the two, so it wins on conflict
  # while leaving the rest of the column's format in place
  p <- write_tmp(list(Sales = xl_sheet(sdf,
    cols = xl_col_spec("qty", format = xl_font(bold = TRUE)),
    table = xl_table(columns = xl_table_column("qty",
                                               format = xl_num_format("0.000"))))))
  s <- xlsx_part(p, "xl/styles.xml", raw = TRUE)
  expect_match(s, "0.000", fixed = TRUE)
  expect_match(s, "<b/>", fixed = TRUE)
})

test_that("a header format cascades over the workbook header format", {
  # an italic header should stay bold and centred like every other header,
  # rather than replacing the default outright
  p <- write_tmp(list(Sales = xl_sheet(sdf, table = xl_table(
    columns = xl_table_column("qty", header_format = xl_font(italic = TRUE))))))
  s <- xlsx_part(p, "xl/styles.xml", raw = TRUE)
  fonts <- regmatches(s, regexpr("<fonts.*?</fonts>", s))
  # some font carries both, which only happens if the two were merged
  expect_match(fonts, "<b/><i/>|<i/><b/>")
})

test_that("a table turns off constant memory", {
  # worksheet_add_table() returns LXW_ERROR_FEATURE_NOT_SUPPORTED in optimize
  # mode, so this is a hard requirement rather than a preference
  cm <- .resolve_constant_memory(list(D = sdf), list(),
                                 list(list(tables = list(1))))
  expect_equal(cm$on, 0L)
  expect_match(cm$reasons, "does not support tables while row streaming")
  # request = TRUE isolates the feature's effect from the size rule
  expect_equal(.resolve_constant_memory(list(D = sdf), list(),
                                        list(list(tables = NULL)),
                                        request = TRUE)$on, 1L)
})

test_that("a table needs a data row", {
  expect_error(write_tmp(list(D = xl_sheet(sdf[0, ], table = xl_table()))),
               "at least one data row")
  # one data row plus a total row still leaves one non-total row
  expect_silent(write_tmp(list(D = xl_sheet(sdf[1, , drop = FALSE],
                                            table = xl_table(total_row = TRUE)))))
})

test_that("two tables on one sheet both reach the file", {
  p <- write_tmp(list(Sales = xl_sheet(sdf, table = list(
    xl_table(range = "A1:A4"), xl_table(range = "B1:B4")))))
  expect_match(xlsx_part(p, "xl/tables/table1.xml", raw = TRUE),
               'ref="A1:A4"', fixed = TRUE)
  expect_match(xlsx_part(p, "xl/tables/table2.xml", raw = TRUE),
               'ref="B1:B4"', fixed = TRUE)
})

test_that("a range reaching past the data columns is refused", {
  # there is no header name for the extra column, and "Column 3" is exactly the
  # generic caption this design exists to make unreachable
  expect_error(write_tmp(list(D = xl_sheet(sdf, table = xl_table(range = "A1:D4")))),
               "writexl has no header name for the extra column")
  # the full width is fine
  expect_silent(write_tmp(list(D = xl_sheet(sdf, table = xl_table(range = "A1:B4")))))
})

test_that("a header row that would land on data is refused", {
  # worksheet_add_table() writes its captions into the range's first row, so a
  # table starting below row 1 would silently replace values with column names
  expect_error(write_tmp(list(D = xl_sheet(sdf, table = xl_table(range = "A2:B4")))),
               "would overwrite data")
  # ... unless the table has no header row, which is the legitimate way to
  # cover the data rows only
  expect_silent(write_tmp(list(D = xl_sheet(sdf, table = xl_table(
    range = "A2:B4", header_row = FALSE)))))
  # and with no header written at all, any first row is fine
  expect_silent(write_tmp(list(D = xl_sheet(sdf, table = xl_table(range = "A2:B4"))),
                          col_names = FALSE))
})

test_that("a column outside the table's range is refused", {
  expect_error(write_tmp(list(D = xl_sheet(sdf, table = xl_table(
    range = "A1:A4", columns = xl_table_column("qty"))))),
    'outside the table\'s range')
})

# ── Conflicts ─────────────────────────────────────────────────────────────────

test_that("a range renders back to A1 notation", {
  expect_equal(.range_a1(c(0L, 0L, 3L, 1L)), "A1:B4")
  expect_equal(.range_a1(c(1L, 2L, 5L, 27L)), "C2:AB6")   # past Z
  expect_equal(.range_a1(c(0L, 0L, 0L, 0L)), "A1:A1")
})

test_that("range overlap is inclusive on both edges", {
  expect_true(.ranges_overlap(c(0, 0, 3, 1), c(0, 0, 3, 1)))
  expect_true(.ranges_overlap(c(0, 0, 3, 1), c(3, 1, 5, 2)))   # touching corner
  expect_false(.ranges_overlap(c(0, 0, 3, 0), c(0, 1, 3, 1)))  # side by side
  expect_false(.ranges_overlap(c(0, 0, 1, 1), c(2, 0, 3, 1)))  # stacked
})

test_that("a table and a sheet autofilter over the same cells is refused", {
  expect_error(write_tmp(list(D = xl_sheet(sdf, table = xl_table(),
                                           autofilter = TRUE))),
               "brings its own filter dropdown")
  # the message must name the remedy that works: turning the table's own
  # autofilter off leaves the sheet-level one over the same cells
  expect_error(write_tmp(list(D = xl_sheet(sdf,
                                           table = xl_table(autofilter = FALSE),
                                           autofilter = TRUE))),
               "Drop the sheet's `autofilter`")
  # an autofilter outside the table is fine
  wide <- cbind(sdf, note = c("a", "b", "c"))
  expect_silent(write_tmp(list(D = xl_sheet(wide,
                                            table = xl_table(range = "A1:B4"),
                                            autofilter = "C1:C4"))))
})

test_that("a filter on a table column is refused", {
  # libxlsxwriter: "Filter conditions within the table are not supported"
  expect_error(write_tmp(list(D = xl_sheet(sdf, table = xl_table(),
                                           filter = xl_filter("qty", ">", 100)))),
               'covers column "qty", which `filter` also filters')
})

test_that("a merge inside a table is refused", {
  expect_error(write_tmp(list(D = xl_sheet(sdf, table = xl_table(),
                                           merge = xl_merge("A1:B1", "x")))),
               "does not allow merged cells inside a table")
  # a merge below the table is fine
  expect_silent(write_tmp(list(D = xl_sheet(sdf,
                                            table = xl_table(range = "A1:B4"),
                                            merge = xl_merge("A6:B6", "Note")))))
})

test_that("two overlapping tables are refused", {
  expect_error(write_tmp(list(D = xl_sheet(sdf, table = list(
    xl_table(range = "A1:B4"), xl_table(range = "B1:B4"))))),
    "tables at A1:B4 and B1:B4 overlap")
  # side by side is fine
  expect_silent(write_tmp(list(D = xl_sheet(sdf, table = list(
    xl_table(range = "A1:A4"), xl_table(range = "B1:B4"))))))
})

# ── Name resolution across the workbook ───────────────────────────────────────

tdf <- data.frame(a = 1:2)
tsheet <- function(x) xl_sheet(tdf, table = x)
resolved <- function(elems) .resolve_table_names(elems, names(elems))

test_that("a sheet name is sanitized into a usable table name", {
  expect_equal(.sanitize_table_name("Sales"), "Sales")
  expect_equal(.sanitize_table_name("My Sheet"), "My_Sheet")   # space is legal
  expect_equal(.sanitize_table_name("a-b/c"), "a_b_c")         # in a sheet name
  expect_equal(.sanitize_table_name("Q1.data"), "Q1.data")     # "." is allowed
  expect_equal(.sanitize_table_name(""), "Table")
  # the three reserved shapes get cleared rather than passed through
  expect_equal(.sanitize_table_name("2024"), "T_2024")
  expect_equal(.sanitize_table_name("A1"), "A1_tbl")
  expect_equal(.sanitize_table_name("C"), "C_tbl")
  # whatever comes out must itself be a legal table name
  for (s in c("Sales", "My Sheet", "a-b/c", "", "2024", "A1", "C", "R"))
    expect_silent(.check_table_name(.sanitize_table_name(s)))
})

test_that("an unnamed table takes its name from the sheet", {
  expect_equal(resolved(list(Sales = tsheet(xl_table()))), list("Sales"))
  expect_equal(resolved(list(`My Sheet` = tsheet(xl_table()))),
               list("My_Sheet"))
})

test_that("generated names do not collide", {
  # regression: comparing an original-case name against a folded set let every
  # generated name through, so two tables on one sheet both came out "Sales"
  expect_equal(resolved(list(Sales = tsheet(list(xl_table(), xl_table())))),
               list(c("Sales", "Sales_2")))
  # ... including against an explicit name, which is a fixed point
  expect_equal(
    resolved(list(Sales = tsheet(list(xl_table(name = "Sales"), xl_table())))),
    list(c("Sales", "Sales_2")))
  # and case must not hide a collision, since Excel folds table names
  expect_equal(
    resolved(list(Sales = tsheet(list(xl_table(name = "sales"), xl_table())))),
    list(c("sales", "Sales_2")))
})

test_that("name = NA is left for Excel to number", {
  expect_equal(resolved(list(Sales = tsheet(xl_table(name = NA)))),
               list(NA_character_))
  # and it does not take part in uniqueness, having no name to collide
  expect_equal(
    resolved(list(S = tsheet(list(xl_table(name = NA), xl_table())))),
    list(c(NA, "S")))
})

test_that("a duplicate explicit name is refused, naming both sheets", {
  # libxlsxwriter never checks this, and Excel repairs the file
  expect_error(
    resolved(list(A = tsheet(xl_table(name = "Totals")),
                  B = tsheet(xl_table(name = "Totals")))),
    'both named "Totals" \\(on sheet "A" and sheet "B"\\)')
  # Excel compares case-insensitively, so this is the same collision
  expect_error(
    resolved(list(A = tsheet(xl_table(name = "Totals")),
                  B = tsheet(xl_table(name = "totals")))),
    "unique across the workbook")
  # the same name twice on one sheet is caught too
  expect_error(
    resolved(list(A = tsheet(list(xl_table(name = "T"), xl_table(name = "T"))))),
    "unique across the workbook")
})

test_that("a workbook with no tables resolves to nothing", {
  expect_equal(resolved(list(A = xl_sheet(tdf), B = xl_sheet(tdf))),
               list(character(0), character(0)))
})

test_that("duplicate table names are caught when writing", {
  expect_error(
    write_tmp(list(A = tsheet(xl_table(name = "Totals")),
                   B = tsheet(xl_table(name = "Totals")))),
    "unique across the workbook")
})

test_that("tables normalise from one or a list", {
  expect_error(write_tmp(list(A = xl_sheet(tdf, table = "t"))),
               "must be an xl_table object")
  expect_error(write_tmp(list(A = xl_sheet(tdf, table = list(xl_table(), "t")))),
               "`table\\[\\[2\\]\\]` must be an xl_table object")
})

test_that("the print methods run", {
  expect_output(print(xl_table(name = "Sales")), "Sales")
  expect_output(print(xl_table()), "unnamed")
  expect_output(print(xl_table(name = NA)), "named by Excel")
  expect_output(print(xl_table(total_row = TRUE)), "total row")
  expect_output(print(xl_table_column("qty", total = "sum")), "total=sum")
  expect_output(print(xl_table_column("qty", formula = "=1")), "formula")
})

test_that("a cached total is written for readers that do not evaluate formulas", {
  df <- data.frame(fruit = c("apple", "pear"), qty = c(3, 4),
                   stringsAsFactors = FALSE)
  sheet <- xl_sheet(df, table = xl_table(total_row = TRUE, columns = list(
    xl_table_column("fruit", total_label = "Total"),
    xl_table_column("qty", total = "sum", total_value = 7))))
  p <- write_tmp(list(D = sheet))
  # readxl reads the cached value, so without one the total row reads blank
  got <- as.data.frame(readxl::read_xlsx(p))
  expect_equal(got$qty[3L], 7)
  expect_equal(got$fruit[3L], "Total")
  # the formula is still what Excel recalculates on open
  expect_match(xlsx_part(p, "xl/tables/table1.xml", raw = TRUE),
               'totalsRowFunction="sum"', fixed = TRUE)
})

test_that("a cached total needs the function it caches", {
  expect_error(xl_table_column("qty", total_value = 7),
               "caches the result of `total`", fixed = TRUE)
  expect_error(xl_table_column("qty", total = "sum", total_value = "7"),
               "must be a single non-NA number")
  expect_error(xl_table_column("qty", total = "sum", total_value = c(1, 2)),
               "must be a single non-NA number")
  expect_error(xl_table_column("qty", total = "sum", total_value = NA_real_),
               "must be a single non-NA number")
  expect_equal(unclass(xl_table_column("q", total = "sum",
                                       total_value = 7))$total_value, 7)
})

test_that("a generated table name is trimmed and made unique", {
  # Excel caps a table name at 255 and compares them case-insensitively
  long <- paste(rep("a", 300), collapse = "")
  expect_equal(nchar(.sanitize_table_name(long)), 250L)
  expect_equal(.unique_table_name("Sales", c("sales", "sales_2")), "Sales_3")
  expect_equal(.unique_table_name("Sales", character(0)), "Sales")
})

test_that("a table needs a row that is neither header nor total", {
  df <- data.frame(a = 1, b = 2)
  expect_error(write_tmp(list(D = xl_sheet(df, table = xl_table(
    range = "A1:B2", total_row = TRUE)))),
    "at least one row that is not the header")
})

test_that("a sheet with no table has no column formats to collect", {
  expect_equal(.table_column_formats(data.frame(a = 1), data.frame(a = 1)),
               list(NULL))
  expect_null(.check_table_conflicts(xl_sheet(data.frame(a = 1)),
                                     data.frame(a = 1), 1L, list()))
})

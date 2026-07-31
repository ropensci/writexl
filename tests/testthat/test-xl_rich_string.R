# Rich strings: one cell, several differently formatted runs.

sheet_xml <- function(path) xlsx_part(path, "xl/worksheets/sheet1.xml", raw = TRUE)

# A one-cell sheet whose column B holds the given rich string.
# constant_memory = TRUE throughout: these tests are about the inline-string
# form used while streaming, which a frame this small would not otherwise get
# (row streaming is only chosen when it would save enough memory to matter).
rich_sheet <- function(rs, ...) {
  df <- data.frame(x = 1L)
  df$r <- xl_cell_general(value = rs, ...)
  write_tmp(df, constant_memory = TRUE)
}

# ── xl_rich_run() ─────────────────────────────────────────────────────────────

test_that("xl_rich_run validates its text", {
  expect_error(xl_rich_run(1), "single non-NA string")
  expect_error(xl_rich_run(NA_character_), "single non-NA string")
  expect_error(xl_rich_run(c("a", "b")), "single non-NA string")
  # Excel cannot represent an empty run, and libxlsxwriter rejects one outright
  expect_error(xl_rich_run(""), "must not be empty")
})

test_that("xl_rich_run validates its format", {
  expect_error(xl_rich_run("a", format = "bold"), "must be an xl_format")
  expect_s3_class(xl_rich_run("a", xl_font(bold = TRUE)), "xl_rich_run")
  expect_s3_class(xl_rich_run("a"), "xl_rich_run")
})

test_that("a run warns about format properties it cannot render", {
  # a run is drawn with a font only, so every other group is ignored
  expect_warning(xl_rich_run("a", xl_fill(background = "yellow")), "fill")
  expect_warning(xl_rich_run("a", xl_border(all = "thin")), "border")
  expect_warning(xl_rich_run("a", xl_align(horizontal = "center")), "align")
  expect_warning(xl_rich_run("a", xl_num_format("0.00")), "num_format")
  expect_warning(xl_rich_run("a", xl_protection(locked = FALSE)), "protection")
  expect_warning(xl_rich_run("a", xl_format(quote_prefix = TRUE)), "quote_prefix")
  # the whole font group is supported, so none of it warns
  expect_silent(xl_rich_run("a", xl_font(bold = TRUE, italic = TRUE,
                                         underline = "single", strikeout = TRUE,
                                         script = "sub", color = "red",
                                         name = "Arial", size = 9)))
})

# ── xl_rich_string() ──────────────────────────────────────────────────────────

test_that("a rich string needs at least two runs", {
  expect_error(xl_rich_string("only one"), "at least 2 runs")
  expect_error(xl_rich_string(), "at least 2 runs")
  expect_error(xl_rich_string(xl_rich_run("a", xl_font(bold = TRUE))),
               "ordinary character value")
  expect_s3_class(xl_rich_string("a", "b"), "xl_rich_string")
})

test_that("bare strings become unformatted runs and lists are flattened", {
  rs <- xl_rich_string("a", xl_rich_run("b", xl_font(bold = TRUE)), "c")
  expect_length(unclass(rs), 3L)
  expect_null(unclass(rs)[[1L]]$format)
  expect_true(is_xl_format(unclass(rs)[[2L]]$format))

  # a character vector contributes one run per element
  expect_length(unclass(xl_rich_string(c("a", "b"))), 2L)
  # nested lists splice in, so runs can be built with lapply()
  built <- lapply(c("x", "y"), xl_rich_run)
  expect_length(unclass(xl_rich_string("pre", built)), 3L)
})

test_that("xl_rich_string rejects run types it cannot use", {
  expect_error(xl_rich_string("a", 42), "must be a string or an xl_rich_run")
  expect_error(xl_rich_string("a", character(0)), "at least one string")
})

test_that("rich string helpers behave", {
  rs <- xl_rich_string("a", "b")
  expect_true(is_xl_rich_string(rs))
  expect_false(is_xl_rich_string("a"))
  # format() gives the plain concatenated text, which auto_colwidth measures
  expect_equal(format(rs), "ab")
  expect_output(print(rs), "2 runs")
  expect_output(print(xl_rich_run("a")), "xl_rich_run")
})

# ── Writing ───────────────────────────────────────────────────────────────────

test_that("a rich string writes one cell with per-run fonts", {
  p <- rich_sheet(xl_rich_string("This is ",
                                 xl_rich_run("bold", xl_font(bold = TRUE)),
                                 " text"))
  s <- sheet_xml(p)
  # three runs, the middle one bold, in the inline-string form used while row
  # streaming is on
  expect_match(s, '<c r="B2" t="inlineStr">', fixed = TRUE)
  expect_match(s, "<r><rPr><b/>", fixed = TRUE)
  expect_match(s, "<t>bold</t>", fixed = TRUE)
  # leading/trailing spaces in a run must be preserved
  expect_match(s, 'xml:space="preserve"', fixed = TRUE)
  # and it reads back as the concatenated text
  skip_if_not_installed("readxl")
  expect_equal(as.data.frame(readxl::read_xlsx(p))$r, "This is bold text")
})

test_that("a rich string is stored in the shared table when streaming is off", {
  # the two storage forms are different code paths in libxlsxwriter, so both
  # need covering; nothing writexl writes turns streaming off on its own here
  local_mocked_bindings(
    .resolve_constant_memory = function(dfs, props, ...)
      list(on = 0L, reasons = "forced off by test")
  )
  p <- rich_sheet(xl_rich_string("This is ",
                                 xl_rich_run("bold", xl_font(bold = TRUE)),
                                 " text"))
  ss <- xlsx_part(p, "xl/sharedStrings.xml", raw = TRUE)
  expect_match(ss, "<r><rPr><b/>", fixed = TRUE)
  expect_match(ss, "<t>bold</t>", fixed = TRUE)
  # the cell now references the table rather than carrying the text inline
  expect_match(sheet_xml(p), '<c r="B2" t="s">', fixed = TRUE)
  skip_if_not_installed("readxl")
  expect_equal(as.data.frame(readxl::read_xlsx(p))$r, "This is bold text")
})

test_that("a cell-level format applies alongside the run fonts", {
  p <- rich_sheet(xl_rich_string("a", xl_rich_run("b", xl_font(bold = TRUE))),
                  format = xl_fill(background = "yellow"))
  # the cell carries a style ...
  expect_match(sheet_xml(p), '<c r="B2" s="', fixed = TRUE)
  # ... and the fill reached styles.xml, which a run alone could not do
  expect_match(xlsx_part(p, "xl/styles.xml", raw = TRUE), "FFFFFF00",
               ignore.case = TRUE)
})

test_that("a rich string counts as one cell and recycles", {
  rs <- xl_rich_string("a", "b")
  expect_equal(length(xl_cell_general(value = rs)), 1L)
  # recycled across a longer column, like any other length-1 value
  df <- data.frame(x = 1:3)
  df$r <- xl_cell_general(value = rs)
  expect_equal(length(df$r), 3L)
  expect_true(file.exists(write_tmp(df)))
})

test_that("a rich string cannot share a cell with a formula or hyperlink", {
  rs <- xl_rich_string("a", "b")
  expect_error(xl_cell_general(value = rs, formula = "=A1"),
               "rich string `value` with a `formula`")
  expect_error(xl_cell_general(value = rs, hyperlink = "https://example.com"),
               "rich string `value` with a `hyperlink`")
  # an NA hyperlink is not a hyperlink, so it is allowed
  expect_s3_class(xl_cell_general(value = rs, hyperlink = NA),
                  "xl_cell_general")
})

test_that("a rich string works with a comment and with auto_colwidth", {
  df <- data.frame(x = 1L)
  df$r <- xl_cell_general(value = xl_rich_string("hello ", "world"),
                          comment = "note")
  p <- write_tmp(list(S = xl_sheet(df, auto_colwidth = TRUE)))
  expect_match(xlsx_part(p, "xl/comments1.xml", raw = TRUE), "note")
  # auto_colwidth must measure the concatenated run text, not treat the cell as
  # empty: "hello world" is 11 characters wide.  (The stored number carries
  # libxlsxwriter's own padding, so compare against the text width rather than
  # pinning that constant.)
  w <- regmatches(sheet_xml(p),
                  regexec('<col min="2" max="2" width="([0-9.]+)"',
                          sheet_xml(p)))[[1L]]
  expect_length(w, 2L)
  expect_gt(as.numeric(w[2L]), nchar("hello world"))
})

test_that("a rich string is accepted by the cell value type check", {
  # the unsupported-type guard must not reject it for being a list
  expect_s3_class(xl_cell_general(value = xl_rich_string("a", "b")),
                  "xl_cell_general")
  # mixed with ordinary values in one column
  cells <- c(xl_cell_general(value = xl_rich_string("a", "b")),
             xl_cell_general(value = 1.5))
  expect_equal(length(cells), 2L)
  df <- data.frame(x = 1:2)
  df$r <- cells
  expect_true(file.exists(write_tmp(df)))
})

# ── as.character() ────────────────────────────────────────────────────────────

test_that("as.character() drops the runs' fonts and keeps the text", {
  rs <- xl_rich_string("This is ", xl_rich_run("bold", xl_font(bold = TRUE)),
                       " text")
  expect_equal(as.character(rs), "This is bold text")
  # which is what a data frame column of them shows
  expect_equal(as.character(xl_cell_general(value = rs)), "This is bold text")
})

test_that("as.character() of a cell is what the cell displays", {
  expect_equal(as.character(xl_cell_general(value = c(1.5, 2))), c("1.5", "2"))
  expect_equal(as.character(xl_cell_general(value = "a")), "a")
  # a formula's displayed value comes from Excel, so writexl has none to give
  expect_equal(as.character(xl_cell_general(formula = "=SUM(A1:A2)")),
               NA_character_)
  expect_equal(as.character(xl_cell_general(value = NA)), NA_character_)
})

test_that("a run takes its value from anything that can render itself", {
  expect_equal(xl_rich_run(xl_cell_general(value = "a"))$value, "a")
  expect_equal(xl_rich_run(factor("a"))$value, "a")
  # a run is one font by definition, so a rich string cannot be one
  expect_error(xl_rich_run(xl_rich_string("a", "b")), "cannot be a rich string")
})

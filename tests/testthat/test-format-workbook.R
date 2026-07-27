# Tests for the workbook layer: xl_properties(), xl_workbook(), and the
# de-hardcoded default formats.

test_that("xl_properties validates its arguments", {
  expect_error(xl_properties(default_format = "x"), "must be an xl_format")
  expect_error(xl_properties(header_format = 1), "must be an xl_format")
  expect_error(xl_properties(custom = list(1, 2)), "must be a named list")
  expect_error(xl_properties(window_size = 100), "length-2")
  expect_error(xl_properties(names = list("=1")), "named list")
  p <- xl_properties(title = "T", author = "A")
  expect_s3_class(p, "xl_properties")
})

test_that("xl_workbook validates and stores its parts", {
  expect_error(xl_workbook(1:3), "must be a data frame")
  expect_error(xl_workbook(data.frame(a = 1), properties = list()),
               "must be an xl_properties")
  wb <- xl_workbook(data.frame(a = 1), col_names = FALSE)
  expect_s3_class(wb, "xl_workbook")
  expect_false(wb$col_names)
  expect_length(wb$sheets, 1L)
})

test_that("print methods run", {
  expect_output(print(xl_properties(title = "T", author = "A", read_only = TRUE)),
                "xl_properties")
  expect_output(print(xl_workbook(data.frame(a = 1))), "xl_workbook")
})

test_that("document metadata is written to docProps/core.xml", {
  wb <- xl_workbook(data.frame(a = 1),
                    properties = xl_properties(title = "My Report",
                                               author = "Bill",
                                               company = "HP",
                                               subject = "Q1"))
  tmp <- write_tmp(wb)
  core <- xlsx_part(tmp, "docProps/core.xml", raw = TRUE)
  expect_match(core, "My Report")
  expect_match(core, "Bill")
  expect_match(core, "Q1")
})

test_that("custom properties of every scalar type are written", {
  wb <- xl_workbook(data.frame(a = 1),
                    properties = xl_properties(custom = list(
                      Project  = "Alpha",     # string
                      Revision = 3L,          # integer
                      Score    = 9.5,         # number
                      Final    = TRUE         # boolean
                    )))
  tmp <- write_tmp(wb)
  cust <- xlsx_part(tmp, "docProps/custom.xml", raw = TRUE)
  expect_match(cust, "Project");  expect_match(cust, "Alpha")
  expect_match(cust, "Revision"); expect_match(cust, "<vt:i4>3")
  expect_match(cust, "Score");    expect_match(cust, "9.5")
  expect_match(cust, "Final");    expect_match(cust, "<vt:bool>true")
})

test_that("read_only, window size, and defined names reach workbook.xml", {
  wb <- xl_workbook(data.frame(a = 1),
                    properties = xl_properties(read_only = TRUE,
                                               window_size = c(1000L, 600L),
                                               names = list(tax = "=0.2")))
  tmp <- write_tmp(wb)
  wx <- xlsx_part(tmp, "xl/workbook.xml", raw = TRUE)
  expect_match(wx, "readOnlyRecommended")
  expect_match(wx, 'windowWidth="15000"')   # set_size converts to twips (x15)
  expect_match(wx, 'windowHeight="9000"')
  expect_match(wx, 'name="tax"')
})

test_that("default_format cascades under every cell and header", {
  wb <- xl_workbook(data.frame(a = 1),
                    properties = xl_properties(default_format = xl_font(name = "Arial")))
  s <- styles_string(wb)
  expect_match(s, "Arial")
})

test_that("header_format and hyperlink_format can be overridden", {
  wb <- xl_workbook(
    data.frame(a = 1),
    properties = xl_properties(
      header_format = xl_font(bold = TRUE, color = "white") + xl_fill(background = "navy")
    )
  )
  s <- styles_string(wb)
  expect_match(s, "FF000080", ignore.case = TRUE)  # navy header fill

  df <- data.frame(a = 1L)
  df$h <- xl_hyperlink_cell("http://x.com")
  wb2 <- xl_workbook(list(S = df),
                     properties = xl_properties(
                       hyperlink_format = xl_font(color = "red", underline = "double")))
  s2 <- styles_string(wb2)
  expect_match(s2, "FFFF0000", ignore.case = TRUE)  # red hyperlink
})

test_that("date_format can be overridden at the workbook level", {
  wb <- xl_workbook(data.frame(d = as.Date("2020-01-01")),
                    properties = xl_properties(date_format = xl_num_format("dd/mm/yyyy")))
  expect_match(styles_string(wb), "dd/mm/yyyy", fixed = TRUE)
})

test_that("date column width and header row height are configurable", {
  wb <- xl_workbook(data.frame(d = as.Date("2020-01-01")),
                    properties = xl_properties(date_col_width = 30,
                                               header_row_height = 40))
  tmp <- write_tmp(wb)
  w <- xlsx_part(tmp, "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_match(w, 'width="30')     # date column widened (Excel pads the value)
  expect_match(w, 'ht="40"')       # header row height
  # datetime width too
  wb2 <- xl_workbook(data.frame(t = as.POSIXct("2020-01-01 00:00", tz = "UTC")),
                     properties = xl_properties(datetime_col_width = 28))
  w2 <- xlsx_part(write_tmp(wb2), "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_match(w2, 'width="28')
})

test_that("layout dimensions are validated", {
  expect_error(xl_properties(date_col_width = -1), "non-negative")
  expect_error(xl_properties(header_row_height = "tall"), "non-negative")
  expect_error(xl_properties(datetime_col_width = c(1, 2)), "single")
})

test_that("workbook col_names / format_headers override write_xlsx args", {
  wb <- xl_workbook(data.frame(a = 1:2), col_names = FALSE)
  tmp <- write_tmp(wb, col_names = TRUE)   # workbook wins -> no header
  rd <- readxl::read_xlsx(tmp, col_names = FALSE)
  expect_equal(as.integer(rd[[1]]), 1:2)
})

test_that("long and duplicate sheet names are still handled", {
  long <- list(data.frame(a = 1))
  names(long) <- paste(rep("x", 40), collapse = "")
  expect_warning(write_tmp(long), "Truncating")
  dup <- list(data.frame(a = 1), data.frame(b = 2))
  names(dup) <- c("S", "S")
  expect_warning(write_tmp(dup), "Deduplicating")
})

test_that("de-hardcoded defaults match the previous behavior (regression)", {
  # plain data frame, no workbook customization
  df <- data.frame(name = c("a", "b"), when = as.Date("2020-01-01") + 0:1)
  df$link <- xl_hyperlink_cell(c("http://x.com", "http://y.com"))
  s <- styles_string(df)
  expect_match(s, "<b/>")                          # bold header
  expect_match(s, 'horizontal="center"')           # centered header
  expect_match(s, "yyyy", fixed = TRUE)            # date number format
  expect_match(s, "FF0000FF", ignore.case = TRUE)  # blue hyperlink font
})

# --- constant memory --------------------------------------------------------

test_that("constant memory resolves to on", {
  cm <- .resolve_constant_memory(list(data.frame(a = 1)), xl_properties())
  expect_equal(cm$on, 1L)
  expect_equal(cm$reasons, character(0))
  expect_equal(.properties_payload(xl_properties())$constant_memory, 1L)
  expect_equal(.properties_payload(xl_properties(), 0L)$constant_memory, 0L)
})

test_that("workbooks write the same content with constant memory on and off", {
  skip_if_not_installed("readxl")
  df <- data.frame(txt = c("alpha", "beta", "alpha"),
                   num = c(1.5, 2, 3),
                   flag = c(TRUE, FALSE, NA),
                   when = as.Date("2020-01-01") + 0:2,
                   stringsAsFactors = FALSE)
  on_path <- write_tmp(list(D = xl_sheet(df, autofilter = TRUE, freeze = "A2")))
  # No feature writexl writes today turns the mode off, so drive the off path by
  # mocking the resolver.  It must stay exercised: the C side and libxlsxwriter
  # behave differently with row streaming disabled, and a phase that needs the
  # mode off (worksheet tables, multi-cell array formulas) will rely on it.
  off_path <- local({
    local_mocked_bindings(
      .resolve_constant_memory = function(elems, props)
        list(on = 0L, reasons = "forced off by test")
    )
    write_tmp(list(D = xl_sheet(df, autofilter = TRUE, freeze = "A2")))
  })

  # Content must match; the bytes need not -- with constant memory off, strings
  # go to the shared string table instead of being written inline.
  rd_on  <- as.data.frame(readxl::read_xlsx(on_path))
  rd_off <- as.data.frame(readxl::read_xlsx(off_path))
  expect_equal(rd_off, rd_on)
  expect_equal(rd_off$txt, df$txt)
  expect_equal(as.Date(rd_off$when), df$when)

  # ... and the worksheet features survive either way
  for (p in c(on_path, off_path)) {
    w <- xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE)
    expect_match(w, '<autoFilter ref="A1:D4"')
    expect_match(w, "<pane")
  }

  # the storage difference itself, so the off path cannot rot unnoticed
  expect_null(xlsx_part(on_path, "xl/sharedStrings.xml"))
  expect_s3_class(xlsx_part(off_path, "xl/sharedStrings.xml"), "xml_document")
})

# ── Hyperlink styling opt-out ────────────────────────────────────────────────

hyperlink_book <- function(hf) {
  df <- data.frame(x = 1L)
  df$h <- xl_hyperlink_cell("https://example.com")
  write_tmp(xl_workbook(list(S = df),
                        properties = xl_properties(hyperlink_format = hf)))
}

test_that("hyperlink_format = NULL writes an unstyled hyperlink", {
  # libxlsxwriter substitutes its own blue-underline format whenever a cell
  # reaches worksheet_write_url_opt() with none, so clearing that default is
  # the only way to get an unstyled link
  s <- xlsx_part(hyperlink_book(NULL), "xl/styles.xml", raw = TRUE)
  expect_false(grepl("<u/>", s, fixed = TRUE))
})

test_that("an empty xl_format() is not the opt-out", {
  # it is the neutral element of the cascade, so the default still applies
  s <- xlsx_part(hyperlink_book(xl_format()), "xl/styles.xml", raw = TRUE)
  expect_true(grepl("<u/>", s, fixed = TRUE))
})

test_that("the default hyperlink format is still blue and underlined", {
  s <- xlsx_part(hyperlink_book(xl_properties()$hyperlink_format),
                 "xl/styles.xml", raw = TRUE)
  expect_true(grepl("<u/>", s, fixed = TRUE))
  expect_match(s, "FF0000FF", ignore.case = TRUE)
})

test_that("hyperlink_format rejects a non-format that is not NULL", {
  expect_error(xl_properties(hyperlink_format = "blue"),
               "must be an xl_format object or NULL")
  expect_s3_class(xl_properties(hyperlink_format = NULL), "xl_properties")
})

# ── Datetime custom properties ───────────────────────────────────────────────

test_that("Date and POSIXct custom properties are written as datetimes", {
  wb <- xl_workbook(data.frame(a = 1),
                    properties = xl_properties(custom = list(
                      Released = as.Date("2024-03-17"),
                      Built    = as.POSIXct("2024-03-17 14:35:09", tz = "UTC")
                    )))
  cust <- xlsx_part(write_tmp(wb), "docProps/custom.xml", raw = TRUE)
  # vt:filetime, not vt:lpwstr -- these used to be stringified
  expect_match(cust, "<vt:filetime>2024-03-17T00:00:00Z</vt:filetime>",
               fixed = TRUE)
  expect_match(cust, "<vt:filetime>2024-03-17T14:35:09Z</vt:filetime>",
               fixed = TRUE)
  expect_false(grepl("<vt:lpwstr>2024-03-17", cust, fixed = TRUE))
})

test_that("a datetime custom property follows the workbook time zone rule", {
  # one zone across the workbook: the zone is dropped and wall-clock written,
  # exactly as for cells, rather than the property using a second policy
  wb <- xl_workbook(
    data.frame(t = as.POSIXct("2024-03-17 09:00:00", tz = "Australia/Perth")),
    properties = xl_properties(custom = list(
      Built = as.POSIXct("2024-03-17 14:35:09", tz = "Australia/Perth")))
  )
  cust <- xlsx_part(write_tmp(wb), "docProps/custom.xml", raw = TRUE)
  expect_match(cust, "2024-03-17T14:35:09Z", fixed = TRUE)
})

test_that("the other custom property types are unaffected", {
  wb <- xl_workbook(data.frame(a = 1),
                    properties = xl_properties(custom = list(
                      S = "txt", I = 3L, N = 9.5, B = TRUE,
                      D = as.Date("2020-01-02"))))
  cust <- xlsx_part(write_tmp(wb), "docProps/custom.xml", raw = TRUE)
  expect_match(cust, "<vt:lpwstr>txt")
  expect_match(cust, "<vt:i4>3")
  expect_match(cust, "9.5")
  expect_match(cust, "<vt:bool>true")
  expect_match(cust, "<vt:filetime>2020-01-02T00:00:00Z", fixed = TRUE)
})

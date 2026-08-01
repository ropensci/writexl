# Page setup: how a worksheet prints.  None of this touches the cells, so the
# assertions are all against the print-related parts of the worksheet XML.

page_xml <- function(page, df = data.frame(x = 1:3)) {
  xlsx_part(write_tmp(list(D = xl_sheet(df, page = page))),
            "xl/worksheets/sheet1.xml", raw = TRUE)
}

# ── Construction and validation ───────────────────────────────────────────────

test_that("an empty page setup is inert", {
  expect_s3_class(xl_page_setup(), "xl_page_setup")
  expect_length(unclass(xl_page_setup()), 0L)
  # a sheet with an empty setup writes the same print XML as one with none
  plain <- xlsx_part(write_tmp(list(D = xl_sheet(data.frame(x = 1:3)))),
                     "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_equal(page_xml(xl_page_setup()), plain)
})

test_that("orientation is validated and applied", {
  expect_error(xl_page_setup(orientation = "sideways"), "orientation")
  expect_match(page_xml(xl_page_setup(orientation = "landscape")),
               'orientation="landscape"', fixed = TRUE)
  expect_match(page_xml(xl_page_setup(orientation = "portrait")),
               'orientation="portrait"', fixed = TRUE)
})

test_that("paper accepts names case-insensitively and raw integers", {
  expect_match(page_xml(xl_page_setup(paper = "A4")), 'paperSize="9"',
               fixed = TRUE)
  expect_match(page_xml(xl_page_setup(paper = "a4")), 'paperSize="9"',
               fixed = TRUE)
  expect_match(page_xml(xl_page_setup(paper = "letter")), 'paperSize="1"',
               fixed = TRUE)
  # an unnamed specialist size is still reachable by index (37 = Monarch)
  expect_match(page_xml(xl_page_setup(paper = 37L)), 'paperSize="37"',
               fixed = TRUE)
  expect_error(xl_page_setup(paper = "A9"), "must be one of")
  expect_error(xl_page_setup(paper = 999L), "paper")
})

test_that("margins accept one value, four values, or a named subset", {
  expect_match(page_xml(xl_page_setup(margins = 1)),
               'left="1" right="1" top="1" bottom="1"', fixed = TRUE)
  expect_match(page_xml(xl_page_setup(margins = c(0.5, 0.6, 0.7, 0.8))),
               'left="0.5" right="0.6" top="0.7" bottom="0.8"', fixed = TRUE)
  # a named subset leaves the other sides at Excel's defaults
  m <- page_xml(xl_page_setup(margins = c(left = 1.25)))
  expect_match(m, 'left="1.25"', fixed = TRUE)
  expect_match(m, 'right="0.7"', fixed = TRUE)
})

test_that("margins are validated", {
  expect_error(xl_page_setup(margins = "wide"), "numeric")
  expect_error(xl_page_setup(margins = c(1, 2)), "one value, four values")
  expect_error(xl_page_setup(margins = c(middle = 1)), "unknown `margins`")
  expect_error(xl_page_setup(margins = -1), "non-negative")
  # zero is a legitimate margin
  expect_match(page_xml(xl_page_setup(margins = 0)), 'left="0"', fixed = TRUE)
})

test_that("scale is bounded to Excel's range", {
  expect_error(xl_page_setup(scale = 5), "scale")
  expect_error(xl_page_setup(scale = 500), "scale")
  expect_match(page_xml(xl_page_setup(scale = 150)), 'scale="150"', fixed = TRUE)
})

test_that("fit_to sets the page counts and the fitToPage flag", {
  # 0 means "as many pages as needed in that direction"
  x <- page_xml(xl_page_setup(fit_to = c(width = 1, height = 0)))
  expect_match(x, 'fitToHeight="0"', fixed = TRUE)
  expect_match(x, '<pageSetUpPr fitToPage="1"/>', fixed = TRUE)

  y <- page_xml(xl_page_setup(fit_to = c(2, 3)))
  expect_match(y, 'fitToWidth="2"', fixed = TRUE)
  expect_match(y, 'fitToHeight="3"', fixed = TRUE)

  # a single value is the width, height defaulting to 1 page
  expect_match(page_xml(xl_page_setup(fit_to = 1)),
               '<pageSetUpPr fitToPage="1"/>', fixed = TRUE)
})

test_that("fit_to is validated", {
  expect_error(xl_page_setup(fit_to = "one page"), "non-negative numbers")
  expect_error(xl_page_setup(fit_to = c(1, 2, 3)), "non-negative numbers")
  expect_error(xl_page_setup(fit_to = -1), "non-negative numbers")
  expect_error(xl_page_setup(fit_to = c(depth = 1)), "unknown `fit_to`")
})

# ── The flag options ──────────────────────────────────────────────────────────

test_that("centring is applied on each axis independently", {
  expect_match(page_xml(xl_page_setup(center_horizontally = TRUE)),
               'horizontalCentered="1"', fixed = TRUE)
  expect_match(page_xml(xl_page_setup(center_vertically = TRUE)),
               'verticalCentered="1"', fixed = TRUE)
  both <- page_xml(xl_page_setup(center_horizontally = TRUE,
                                 center_vertically = TRUE))
  expect_match(both, 'horizontalCentered="1"', fixed = TRUE)
  expect_match(both, 'verticalCentered="1"', fixed = TRUE)
  # FALSE is not TRUE: nothing is written
  expect_false(grepl("horizontalCentered",
                     page_xml(xl_page_setup(center_horizontally = FALSE)),
                     fixed = TRUE))
})

test_that("the print options are applied", {
  expect_match(page_xml(xl_page_setup(black_and_white = TRUE)),
               'blackAndWhite="1"', fixed = TRUE)
  expect_match(page_xml(xl_page_setup(row_col_headers = TRUE)),
               'headings="1"', fixed = TRUE)
  expect_match(page_xml(xl_page_setup(across = TRUE)),
               'pageOrder="overThenDown"', fixed = TRUE)
})

test_that("first_page sets the starting page number", {
  x <- page_xml(xl_page_setup(first_page = 3))
  expect_match(x, 'firstPageNumber="3"', fixed = TRUE)
  # Excel needs the companion flag or it ignores the number
  expect_match(x, 'useFirstPageNumber="1"', fixed = TRUE)
})

test_that("page_view opens the sheet in page-layout view", {
  expect_match(page_xml(xl_page_setup(page_view = TRUE)),
               'view="pageLayout"', fixed = TRUE)
})

# ── Wiring ────────────────────────────────────────────────────────────────────

test_that("settings combine without interfering", {
  x <- page_xml(xl_page_setup(orientation = "landscape", paper = "legal",
                              margins = 0.25, scale = 80,
                              center_horizontally = TRUE, first_page = 2))
  expect_match(x, 'orientation="landscape"', fixed = TRUE)
  expect_match(x, 'paperSize="5"', fixed = TRUE)
  expect_match(x, 'left="0.25"', fixed = TRUE)
  expect_match(x, 'scale="80"', fixed = TRUE)
  expect_match(x, 'horizontalCentered="1"', fixed = TRUE)
  expect_match(x, 'firstPageNumber="2"', fixed = TRUE)
})

test_that("page setup does not disturb the cells", {
  skip_if_not_installed("readxl")
  df <- data.frame(a = 1:3, b = c("x", "y", "z"), stringsAsFactors = FALSE)
  p <- write_tmp(list(D = xl_sheet(df, page = xl_page_setup(
    orientation = "landscape", margins = 1, fit_to = c(1, 0)))))
  expect_equal(as.data.frame(readxl::read_xlsx(p)), df)
})

test_that("`page` must be an xl_page_setup", {
  expect_error(write_tmp(list(D = xl_sheet(data.frame(x = 1), page = "landscape"))),
               "must be an xl_page_setup")
})

test_that("the print method runs", {
  expect_output(print(xl_page_setup(orientation = "landscape", margins = 1)),
                "xl_page_setup")
  expect_output(print(xl_page_setup()), "0 settings")
})

# ── Headers and footers ───────────────────────────────────────────────────────

test_that("header and footer text is written with Excel's codes intact", {
  x <- page_xml(xl_page_setup(header = "&LQ1 report&RPage &P of &N",
                              footer = "&Cdraft"))
  # the & codes survive, XML-escaped
  expect_match(x, "<oddHeader>&amp;LQ1 report&amp;RPage &amp;P of &amp;N</oddHeader>",
               fixed = TRUE)
  expect_match(x, "<oddFooter>&amp;Cdraft</oddFooter>", fixed = TRUE)
})

test_that("a header or footer works without a margin", {
  # no margin means the plain setter rather than the _opt form
  x <- page_xml(xl_page_setup(header = "&Cplain"))
  expect_match(x, "<oddHeader>&amp;Cplain</oddHeader>", fixed = TRUE)
  expect_match(x, 'header="0.3"', fixed = TRUE)   # Excel's default retained
})

test_that("header and footer margins reach pageMargins", {
  x <- page_xml(xl_page_setup(header = "&Ca", footer = "&Cb",
                              header_margin = 0.5, footer_margin = 0.6))
  expect_match(x, 'header="0.5"', fixed = TRUE)
  expect_match(x, 'footer="0.6"', fixed = TRUE)
})

test_that("a header/footer margin must be greater than zero", {
  # libxlsxwriter applies the margin only when it is > 0, so a zero would be
  # silently ignored rather than doing what it says
  expect_error(xl_page_setup(header_margin = 0), "greater than 0")
  expect_error(xl_page_setup(footer_margin = 0), "greater than 0")
  expect_error(xl_page_setup(header_margin = -1), "header_margin")
})

test_that("header and footer length is capped at 255 characters", {
  ok <- strrep("x", 255)
  expect_s3_class(xl_page_setup(header = ok), "xl_page_setup")
  expect_error(xl_page_setup(header = strrep("x", 256)),
               "at most 255 characters")
  expect_error(xl_page_setup(footer = strrep("x", 256)),
               "at most 255 characters")
})

test_that("the length limit counts characters, not bytes", {
  # libxlsxwriter measures with lxw_utf8_strlen(); a 200-character string of
  # 3-byte characters is well over 255 bytes but under the limit
  s <- strrep("\u00e9", 200)
  expect_gt(nchar(s, type = "bytes"), 255L)
  expect_s3_class(xl_page_setup(header = s), "xl_page_setup")
})

# ── Header and footer images ──────────────────────────────────────────────────

# A minimal valid 1x1 PNG, as in test-xl_image.R.
HF_PNG <- as.raw(c(
  0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
  0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
  0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00,
  0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
  0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82))

hf_logo <- function() {
  p <- tempfile(fileext = ".png")
  writeBin(HF_PNG, p)
  p
}

# Write a sheet and list every part of the resulting zip.
hf_parts <- function(page) {
  p <- write_tmp(list(S = xl_sheet(data.frame(a = 1:3), page = page)))
  d <- tempfile()
  dir.create(d)
  utils::unzip(p, exdir = d)
  list(files = list.files(d, recursive = TRUE),
       sheet = xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE))
}

test_that("both spellings of the image placeholder are counted", {
  expect_equal(.hf_placeholders("&L&G"), 1L)
  expect_equal(.hf_placeholders("&G&C&G"), 2L)
  expect_equal(.hf_placeholders("&L&[Picture]"), 1L)
  expect_equal(.hf_placeholders("&G&[Picture]"), 2L)
  expect_equal(.hf_placeholders("plain"), 0L)
  expect_equal(.hf_placeholders(NULL), 0L)
})

test_that("a header image reaches the file with everything Excel needs", {
  # the image alone is not enough: Excel renders a header image through a VML
  # drawing, and without the legacyDrawingHF reference nothing appears
  r <- hf_parts(xl_page_setup(header = "&L&G",
                              header_image = list(left = hf_logo())))
  expect_true("xl/media/image1.png" %in% r$files)
  expect_true("xl/drawings/vmlDrawing1.vml" %in% r$files)
  expect_true("xl/drawings/_rels/vmlDrawing1.vml.rels" %in% r$files)
  expect_match(r$sheet, "<legacyDrawingHF r:id=", fixed = TRUE)
})

test_that("a footer image works, and may be given in memory", {
  r <- hf_parts(xl_page_setup(footer = "&R&G",
                              footer_image = list(right = HF_PNG)))
  expect_true("xl/media/image1.png" %in% r$files)
  expect_match(r$sheet, "<legacyDrawingHF r:id=", fixed = TRUE)
})

test_that("all three positions can carry an image", {
  a <- hf_logo()
  b <- tempfile(fileext = ".png"); writeBin(c(HF_PNG, as.raw(0)), b)
  cimg <- tempfile(fileext = ".png"); writeBin(c(HF_PNG, as.raw(c(0, 0))), cimg)
  r <- hf_parts(xl_page_setup(
    header = "&L&G&C&G&R&G",
    header_image = list(left = a, center = b, right = cimg)))
  # three distinct images, so three media files -- identical ones would be
  # deduplicated by libxlsxwriter and prove nothing
  expect_true(all(sprintf("xl/media/image%d.png", 1:3) %in% r$files))
})

test_that("placeholders and images must agree in number", {
  # libxlsxwriter enforces this too, but its error does not say which side is
  # short, and that is the only thing the caller needs to know
  expect_error(xl_page_setup(header = "&LPlain",
                             header_image = list(left = hf_logo())),
               "gives 1 image\\(s\\) but `header` has 0 image placeholder")
  expect_error(xl_page_setup(header = "&L&G"),
               "has 1 image placeholder.*but `header_image` gives 0")
  expect_error(xl_page_setup(footer = "&C&G"), "gives 0 image")
  # the message says how to fix the commoner direction
  expect_error(xl_page_setup(header = "&LPlain",
                             header_image = list(left = hf_logo())),
               "add &G where each image should sit")
})

test_that("header image positions are validated", {
  expect_error(xl_page_setup(header = "&L&G",
                             header_image = list(middle = hf_logo())),
               "unknown position\\(s\\): middle")
  expect_error(xl_page_setup(header = "&L&G", header_image = hf_logo()),
               "must be named by position")
  expect_error(xl_page_setup(header = "&L&G",
                             header_image = list(left = "nope.png")),
               "file does not exist")
  # the image itself goes through the same checks as xl_image()
  bad <- tempfile(fileext = ".png"); writeLines("not an image", bad)
  expect_error(xl_page_setup(header = "&L&G", header_image = list(left = bad)),
               "is not a PNG, JPEG, GIF or BMP")
})

test_that("a header without an image is unaffected", {
  r <- hf_parts(xl_page_setup(header = "&CPlain"))
  expect_match(r$sheet, "<oddHeader>&amp;CPlain</oddHeader>", fixed = TRUE)
  expect_false(any(grepl("vmlDrawing", r$files)))
  # a literal ampersand or an unrelated code is still fine
  expect_s3_class(xl_page_setup(header = "R&&D&Cmid"), "xl_page_setup")
})

test_that("header and footer do not disturb the cells", {
  skip_if_not_installed("readxl")
  df <- data.frame(a = 1:2, stringsAsFactors = FALSE)
  p <- write_tmp(list(D = xl_sheet(df, page = xl_page_setup(
    header = "&Ctitle", footer = "&Rpage &P", header_margin = 0.4))))
  expect_equal(as.data.frame(readxl::read_xlsx(p)), df)
})

# ── Print area and repeat rows/columns ───────────────────────────────────────
#
# These write defined names into workbook.xml, not into the worksheet part.

defined_names <- function(page, df = data.frame(x = 1:30)) {
  wb <- xlsx_part(write_tmp(list(D = xl_sheet(df, page = page))),
                  "xl/workbook.xml", raw = TRUE)
  m <- regmatches(wb, regexpr("<definedNames>.*</definedNames>", wb))
  if (!length(m)) NA_character_ else m
}

test_that("print_area becomes a Print_Area defined name", {
  expect_match(defined_names(xl_page_setup(print_area = "A1:B20")),
               "_xlnm.Print_Area", fixed = TRUE)
  expect_match(defined_names(xl_page_setup(print_area = "A1:B20")),
               "D!$A$1:$B$20", fixed = TRUE)
})

test_that("print_area accepts a data-frame-relative spec", {
  df <- data.frame(a = 1:5, b = 1:5, cc = 1:5)
  x <- defined_names(xl_page_setup(print_area = list(rows = 1:3, cols = c("a", "b"))),
                     df = df)
  # rows 1:3 are data rows, so sheet rows 2:4 with the header
  expect_match(x, "D!$A$2:$B$4", fixed = TRUE)
})

test_that("repeat_rows and repeat_cols become Print_Titles", {
  # a count repeats that many sheet rows, starting at the header
  expect_match(defined_names(xl_page_setup(repeat_rows = 1)),
               "D!$1:$1", fixed = TRUE)
  expect_match(defined_names(xl_page_setup(repeat_rows = 2)),
               "D!$1:$2", fixed = TRUE)
  expect_match(defined_names(xl_page_setup(repeat_cols = 1)),
               "D!$A:$A", fixed = TRUE)
  # a range string works too
  expect_match(defined_names(xl_page_setup(repeat_rows = "1:2")),
               "D!$1:$2", fixed = TRUE)
  expect_match(defined_names(xl_page_setup(repeat_cols = "A:B")),
               "D!$A:$B", fixed = TRUE)
  # both together
  expect_match(defined_names(xl_page_setup(repeat_rows = 1, repeat_cols = "A:A")),
               "_xlnm.Print_Titles", fixed = TRUE)
})

test_that("repeat arguments reject the wrong kind of range", {
  expect_error(xl_page_setup(repeat_rows = "A:B"), "row range")
  expect_error(xl_page_setup(repeat_cols = "1:2"), "column range")
  expect_error(xl_page_setup(repeat_rows = 0), "repeat_rows")
  expect_error(xl_page_setup(repeat_rows = list(1)), "count or range string")
})

test_that("a print area outside the data still resolves", {
  # a print area is a sheet range, not restricted to the written cells
  expect_match(defined_names(xl_page_setup(print_area = "A1:Z100")),
               "D!$A$1:$Z$100", fixed = TRUE)
})

# ── Page breaks ───────────────────────────────────────────────────────────────

breaks_xml <- function(page, df = data.frame(x = 1:30)) page_xml(page, df)

test_that("page breaks are written 0-based, one before the named row", {
  # h_breaks = 21 means the new page starts at row 21, which is 0-based 20
  x <- breaks_xml(xl_page_setup(h_breaks = c(11, 21)))
  expect_match(x, '<brk id="10"', fixed = TRUE)
  expect_match(x, '<brk id="20"', fixed = TRUE)
  expect_match(x, 'manualBreakCount="2"', fixed = TRUE)
})

test_that("vertical breaks accept numbers and column letters", {
  expect_match(breaks_xml(xl_page_setup(v_breaks = 2)), "<colBreaks",
               fixed = TRUE)
  expect_match(breaks_xml(xl_page_setup(v_breaks = 2)), '<brk id="1"', fixed = TRUE)
  # column "C" is the third column, so a break before it is 0-based 2
  expect_match(breaks_xml(xl_page_setup(v_breaks = "C")), '<brk id="2"',
               fixed = TRUE)
})

test_that("breaks are sorted and deduplicated", {
  x <- breaks_xml(xl_page_setup(h_breaks = c(21, 11, 21)))
  expect_match(x, 'manualBreakCount="2"', fixed = TRUE)
  # sorted: 10 must appear before 20
  expect_lt(regexpr('<brk id="10"', x, fixed = TRUE),
            regexpr('<brk id="20"', x, fixed = TRUE))
})

test_that("a break at position 1 is rejected", {
  # 0-based 0 terminates libxlsxwriter's array, so this would silently drop
  # every break rather than being merely useless
  expect_error(xl_page_setup(h_breaks = 1), "2 or greater")
  expect_error(xl_page_setup(v_breaks = 1), "2 or greater")
  expect_error(xl_page_setup(h_breaks = c(5, 1)), "2 or greater")
})

test_that("the break count is capped at Excel's limit", {
  expect_s3_class(xl_page_setup(h_breaks = 2:1024), "xl_page_setup")   # 1023
  expect_error(xl_page_setup(h_breaks = 2:1025), "at most 1023")
})

test_that("breaks are validated", {
  expect_error(xl_page_setup(h_breaks = "row 5"), "numeric vector")
  expect_error(xl_page_setup(h_breaks = c(5, NA)), "numeric vector")
})

test_that("combining breaks with fit_to warns", {
  # Excel ignores manual breaks when fit-to-page is on
  expect_warning(xl_page_setup(h_breaks = 10, fit_to = c(1, 0)),
                 "ignores manual page breaks")
  expect_silent(xl_page_setup(h_breaks = 10))
})

test_that("print area and breaks together leave the cells alone", {
  skip_if_not_installed("readxl")
  df <- data.frame(a = 1:10, stringsAsFactors = FALSE)
  p <- write_tmp(list(D = xl_sheet(df, page = xl_page_setup(
    print_area = "A1:A11", repeat_rows = 1, h_breaks = 6))))
  expect_equal(as.data.frame(readxl::read_xlsx(p)), df)
})

test_that("a header or footer image position may be given only once", {
  logo <- png_file()
  expect_error(xl_page_setup(header = "&L&G",
                             header_image = list(left = logo, left = logo)),
               "each position may be given only once")
  expect_error(xl_page_setup(header = "&L&G",
                             header_image = list(middle = logo)),
               "unknown position")
})

test_that("a header image may be an in-memory picture", {
  skip_if_not(isTRUE(capabilities("png")))
  p <- write_tmp(list(D = xl_sheet(data.frame(a = 1:3), page = xl_page_setup(
    header = "&L&G",
    header_image = list(left = as.raster(matrix("red", 2, 2)))))))
  d <- tempfile(); dir.create(d); utils::unzip(p, exdir = d)
  expect_true(any(grepl("^xl/media/", list.files(d, recursive = TRUE))))
})

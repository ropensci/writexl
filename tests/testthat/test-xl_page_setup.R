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
  expect_false(grepl('horizontalCentered',
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

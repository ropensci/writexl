# Cross-function API consistency.
#
# An argument that appears in several exported functions should mean the same
# thing and appear in the same relative order, so that reading one signature
# teaches you the others.  This is a mechanical gate rather than a convention in
# prose, because a convention in prose is exactly what got violated: `format`
# was briefly moved after `author` in xl_comment(), breaking the idiom that
# `format` follows the content arguments everywhere else.

# Argument names and positions for every exported function.
api_arg_positions <- function() {
  ns <- asNamespace("writexl")
  out <- list()
  for (f in sort(getNamespaceExports("writexl"))) {
    obj <- get(f, envir = ns)
    if (!is.function(obj)) next
    a <- names(formals(obj))
    a <- a[a != "..."]
    if (length(a)) out[[f]] <- a
  }
  out
}

# Same name, unrelated concept.  These are allowed to disagree on order because
# they are not really the same argument: a comment *box's* pixel width and a
# worksheet *column's* character width have nothing to do with each other.
FALSE_FRIENDS <- list(
  c("xl_comment", "xl_col_spec"),
  c("xl_comment", "xl_row_spec"),
  # `name` in xl_hyperlink() is the deprecated alias for `value`, kept last on
  # purpose; in xl_chart_series() it is the series label shown in the legend.
  c("xl_hyperlink", "xl_chart_series"),
  c("xl_hyperlink", "xl_chart_axis"),
  # `position` in xl_image() is how the picture behaves when rows and columns
  # resize; in xl_chart_axis() it is whether the categories sit on the tick
  # marks or between them.  Different vocabularies, not one argument.
  c("xl_image", "xl_chart_axis")
)

is_false_friend <- function(a, b) {
  any(vapply(FALSE_FRIENDS,
             function(p) setequal(p, c(a, b)), logical(1)))
}

test_that("shared arguments keep the same relative order across functions", {
  api <- api_arg_positions()
  fns <- names(api)
  conflicts <- character(0)
  for (i in seq_along(fns)) for (j in seq_along(fns)) {
    if (j <= i) next
    fa <- fns[i]; fb <- fns[j]
    if (is_false_friend(fa, fb)) next
    common <- intersect(api[[fa]], api[[fb]])
    if (length(common) < 2L) next
    oa <- order(match(common, api[[fa]]))
    ob <- order(match(common, api[[fb]]))
    if (!identical(oa, ob))
      conflicts <- c(conflicts, sprintf(
        "%s(%s) vs %s(%s)", fa, paste(common[oa], collapse = ", "),
        fb, paste(common[ob], collapse = ", ")))
  }
  expect_equal(conflicts, character(0))
})

test_that("`format` follows the content arguments, everywhere it appears", {
  # The idiom across the package: identify the content, then style it, then
  # everything else.  The number is how many content arguments precede it.
  expected <- c(xl_num_format = 1L,   # `format` *is* the content here
                xl_formula = 2L, xl_rich_run = 2L, xl_comment = 2L,
                xl_merge = 3L, xl_hyperlink = 3L, xl_hyperlink_cell = 3L,
                xl_cell_general = 4L)
  api <- api_arg_positions()
  got <- vapply(names(expected), function(f) {
    p <- match("format", api[[f]])
    if (is.na(p)) NA_integer_ else as.integer(p)
  }, integer(1))
  expect_equal(got, expected)
})

test_that("data comes first in the functions that take data", {
  api <- api_arg_positions()
  expect_equal(api$write_xlsx[1L], "x")
  expect_equal(api$xl_sheet[1L], "data")
  expect_equal(api$xl_workbook[1L], "sheets")
  # and the range/target identifies what is being acted on
  expect_equal(api$xl_merge[1L], "range")
  expect_equal(api$xl_col_spec[1L], "cols")
  expect_equal(api$xl_row_spec[1L], "rows")
})

test_that("the object-taking sheet arguments are named for their constructor", {
  # xl_sheet(page = xl_page_setup(), view = xl_sheet_view(), merge = xl_merge())
  # -- the argument name should make the constructor obvious
  api <- api_arg_positions()
  for (a in c("page", "view", "merge"))
    expect_true(a %in% api$xl_sheet, label = a)
})

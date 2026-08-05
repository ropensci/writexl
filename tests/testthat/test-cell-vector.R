# An xl_cell_general is a vector, not a list.
#
# The distinction is not cosmetic.  `[<-.data.frame` branches on
# is.list(value): a list is read as a *list of columns*, so a list-backed cell
# vector could not be assigned with `df[, j] <- cells` at all -- it scattered
# one record per column and dropped the class, which is how writexl 2.0.0 broke
# a reverse dependency that had worked since 1.5.4.
#
# The carrier is an integer index; the per-cell records ride in an attribute,
# the shape factor uses.  These tests pin the vector behaviour rather than the
# representation, except where the representation *is* the behaviour.

test_that("a cell vector is not a list", {
  x <- xl_cell_general(value = 1:3)
  expect_false(is.list(x))
  expect_type(unclass(x), "integer")
  expect_length(x, 3L)
  # the records are the content, and there is one per cell
  expect_length(.cell_records(x), 3L)
})

test_that("mixed types survive: the carrier constrains nothing", {
  x <- xl_cell_general(value = list(1.5, "text", TRUE, as.Date("2024-01-01")))
  expect_equal(vapply(.cell_records(x), function(r) class(r$value)[1L],
                      character(1)),
               c("numeric", "character", "logical", "Date"))
  # and they reach the file as their own types.  The Date arrives as its Excel
  # serial rather than a formatted date: a mixed cell column carries no column
  # number format, which is behaviour this rewrite inherited unchanged --
  # verified identical against the pre-rewrite tree, and pinned here so that
  # fixing it is a deliberate act rather than an accident.
  p <- write_tmp(data.frame(a = 1:4, b = x))
  got <- as.data.frame(readxl::read_xlsx(p))
  expect_equal(got$b, c("1.5", "text", "TRUE", "45292"))
})

# ── Assignment into a data frame ──────────────────────────────────────────────

test_that("`df[, j] <- cells` keeps the class, which is the whole point", {
  d <- data.frame(a = 1:3)
  d[, 2] <- xl_cell_general(value = 4:6)
  expect_s3_class(d[[2]], "xl_cell_general")
  expect_length(.cell_records(d[[2]]), 3L)
})

test_that("the pattern that broke a reverse dependency works", {
  # BioMonTools seeds a one-row frame with a formula cell, then fills rows in
  # one at a time.  Under a list-backed cell vector the seed warned, dropped
  # the class, and left a bare list column that write_xlsx() refused.
  NOTES <- data.frame(matrix(ncol = 3))
  NOTES[, 2] <- xl_formula('=""')
  expect_s3_class(NOTES[[2]], "xl_cell_general")

  NOTES[1, 1] <- "Title"
  # writing a formula string as a value is what this pattern does; the warning
  # is the subject of its own test below
  expect_warning(NOTES[3, 2] <- "=LEFT(A1,1)", "not a formula")
  expect_s3_class(NOTES[[2]], "xl_cell_general")
  # the carrier and the records stay in step: a stale record list is how this
  # corrupts silently rather than loudly
  expect_length(.cell_records(NOTES[[2]]), length(NOTES[[2]]))
  expect_equal(nrow(NOTES), length(NOTES[[2]]))
  expect_equal(.cell_records(NOTES[[2]])[[3L]]$value, "=LEFT(A1,1)")

  expect_silent(write_tmp(list(NOTES = NOTES)))
})

test_that("`df[i, j] <- value` reaches the record, not the carrier", {
  d <- data.frame(a = 1:3)
  d[, 2] <- xl_cell_general(value = c(10, 20, 30))
  d[2, 2] <- 99
  expect_s3_class(d[[2]], "xl_cell_general")
  expect_equal(.cell_records(d[[2]])[[2L]]$value, 99)
  # unchanged cells keep theirs
  expect_equal(.cell_records(d[[2]])[[1L]]$value, 10)
  # and an xl_cell_general on the right replaces the cell outright
  d[3, 2] <- xl_cell_general(formula = "=A1+1")
  expect_equal(.cell_records(d[[2]])[[3L]]$formula, "=A1+1")
})

test_that("the other ways a cell column gets built still work", {
  x <- xl_cell_general(value = 1:3)
  d <- data.frame(a = 1:3)
  d$b <- x
  expect_s3_class(d$b, "xl_cell_general")
  expect_s3_class(data.frame(a = 1:3, b = x)$b, "xl_cell_general")
  d2 <- data.frame(a = 1:3)
  d2[["c"]] <- x
  expect_s3_class(d2$c, "xl_cell_general")
})

# ── Vector semantics ──────────────────────────────────────────────────────────

test_that("subsetting, concatenation and repetition act on cells", {
  x <- xl_cell_general(value = c(1, 2, 3))
  expect_length(x[2:3], 2L)
  expect_equal(x[2:3][[1L]]$value, 2)
  expect_length(c(x, x), 6L)
  expect_length(rep(x, 2L), 6L)
  expect_length(rep(x, length.out = 5L), 5L)
  # `[[` is one cell's record, never the carrier's integer
  expect_type(x[[1L]], "list")
  expect_equal(x[[1L]]$value, 1)
})

test_that("a cell vector round-trips through a write unchanged by the rewrite", {
  # the representation changed; what reaches the file must not have
  df <- data.frame(a = 1:3)
  df$b <- xl_cell_general(value = list(1, "two", NA),
                          formula = c(NA, NA, "=A2+1"),
                          format = xl_font(bold = TRUE))
  p <- write_tmp(list(D = df))
  x <- xlsx_part(p, "xl/worksheets/sheet1.xml", raw = TRUE)
  expect_match(x, "A2+1", fixed = TRUE)
  expect_match(xlsx_part(p, "xl/styles.xml", raw = TRUE), "<b/>", fixed = TRUE)
  got <- as.data.frame(readxl::read_xlsx(p))
  expect_equal(got$b[1L], "1")
  expect_equal(got$b[2L], "two")
})

test_that("assigning a bare string writes a string, not a formula", {
  # writexl's model is per cell: a formula is xl_formula(), anything else is a
  # value.  writexl 1.5.4 carried the "this is a formula column" idea on the
  # column's class, so a string assigned into one became a formula; there is no
  # column-level kind here, and `d[i, j] <- "=A1"` writes the six characters.
  d <- data.frame(a = 1:2)
  d[, 2] <- xl_formula(c("=A1", "=A2"))
  expect_warning(d[2, 2] <- "=SUM(A1:A2)", "not a formula")
  expect_equal(.cell_records(d[[2]])[[2L]]$value, "=SUM(A1:A2)")
  expect_true(is.na(.cell_records(d[[2]])[[2L]]$formula))
  # the formula spelling is explicit, and still available
  d[2, 2] <- xl_formula("=SUM(A1:A2)")
  expect_equal(.cell_records(d[[2]])[[2L]]$formula, "=SUM(A1:A2)")
})

test_that("a frame extended past a cell column pads it with empty cells", {
  # R extends rows behind our back, through xpdrows.data.frame(); the records
  # are padded on read so the new cells are blank rather than absent
  d <- data.frame(matrix(ncol = 2))
  d[, 1] <- xl_cell_general(value = 1)
  d[4, 2] <- "x"
  expect_equal(nrow(d), 4L)
  expect_length(d[[1]], 4L)
  expect_length(.cell_records(d[[1]]), 4L)
  expect_true(is.na(.cell_records(d[[1]])[[3L]]$value))
  expect_silent(write_tmp(list(D = d)))
})

test_that("`[[<-` replaces one cell's record", {
  x <- xl_cell_general(value = c(1, 2, 3))
  x[[2L]] <- .cell_records(xl_cell_general(formula = "=A1"))[[1L]]
  expect_s3_class(x, "xl_cell_general")
  expect_equal(x[[2L]]$formula, "=A1")
  expect_equal(x[[1L]]$value, 1)
  expect_length(x, 3L)
})

test_that("`length<-` grows and shrinks a cell vector", {
  x <- xl_cell_general(value = c(1, 2, 3))
  length(x) <- 5L
  expect_length(x, 5L)
  expect_length(.cell_records(x), 5L)
  # the new cells are empty rather than absent
  expect_true(is.na(x[[5L]]$value))
  length(x) <- 2L
  expect_length(x, 2L)
  expect_equal(x[[2L]]$value, 2)
})

test_that("writing a formula string as a value says so", {
  # a cell column has no column-wide "these are all formulas", so this writes
  # the characters.  That is the rule, but it is a quiet way to lose every
  # formula on a sheet, so it warns -- and points at the spelling that works.
  d <- data.frame(a = 1:2)
  d[, 2] <- xl_cell_general(value = c(1, 2))
  expect_warning(d[2, 2] <- "=SUM(A1)", "Use xl_formula")
  expect_warning(d[2, 2] <- "=SUM(A1)", "df[[j]] <- xl_formula", fixed = TRUE)

  # and stays quiet everywhere the intent is already clear
  expect_silent(d[2, 2] <- "plain text")
  expect_silent(d[2, 2] <- 99)
  expect_silent(d[2, 2] <- NA)
  expect_silent(d[2, 2] <- xl_formula("=SUM(A1)"))
  expect_silent(d[2, 2] <- xl_cell_general(formula = "=SUM(A1)"))
  # constructing a text cell that looks like a formula is deliberate and
  # documented, so it is not the assignment's business to second-guess it
  expect_silent(xl_cell_general(value = "=SUM(A1)"))
})

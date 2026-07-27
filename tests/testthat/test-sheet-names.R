# Sheet names are repaired to Excel's rules rather than erroring; see
# R/sheet_names.R for the rules and the ordering.

# The sheet names actually written to a workbook.
written_names <- function(x) {
  p <- suppressWarnings(write_tmp(x))
  wb <- xlsx_part(p, "xl/workbook.xml", raw = TRUE)
  m <- regmatches(wb, gregexpr('<sheet name="[^"]*"', wb))[[1L]]
  sub('"$', "", sub('^<sheet name="', "", m))
}

one <- function(...) {
  l <- list(...)
  lapply(l, function(i) data.frame(a = 1))
}

# ── Length ────────────────────────────────────────────────────────────────────

test_that("an over-long name is truncated to exactly 31 characters", {
  nm <- strrep("x", 40)
  s <- setNames(list(data.frame(a = 1)), nm)
  expect_warning(write_tmp(s), "Truncating")
  # the old code truncated to 29 while claiming 31
  expect_equal(written_names(s), strrep("x", 31))
  expect_equal(nchar(written_names(s)), 31L)
})

test_that("a name of exactly 31 characters is left alone", {
  nm <- strrep("x", 31)
  s <- setNames(list(data.frame(a = 1)), nm)
  expect_silent(write_tmp(s))
  expect_equal(written_names(s), nm)
})

test_that("a name of 32 characters loses exactly one character", {
  nm <- strrep("x", 32)
  s <- setNames(list(data.frame(a = 1)), nm)
  expect_equal(written_names(s), strrep("x", 31))
})

# ── Invalid characters ────────────────────────────────────────────────────────

test_that("characters Excel forbids are replaced", {
  s <- setNames(list(data.frame(a = 1)), "2024/Q1")
  expect_warning(write_tmp(s), "Replacing characters")
  expect_equal(written_names(s), "2024-Q1")
})

test_that("every forbidden character is handled", {
  for (ch in c("[", "]", ":", "*", "?", "/", "\\")) {
    s <- setNames(list(data.frame(a = 1)), paste0("a", ch, "b"))
    expect_equal(written_names(s), "a-b", label = ch)
  }
})

test_that("a name of nothing but forbidden characters still writes", {
  s <- setNames(list(data.frame(a = 1)), "///")
  expect_equal(written_names(s), "---")
})

test_that("leading and trailing apostrophes are stripped", {
  expect_equal(written_names(setNames(list(data.frame(a = 1)), "'q1'")), "q1")
  expect_equal(written_names(setNames(list(data.frame(a = 1)), "'q1")), "q1")
  expect_equal(written_names(setNames(list(data.frame(a = 1)), "q1'")), "q1")
  # an interior apostrophe is legal and must survive
  expect_equal(written_names(setNames(list(data.frame(a = 1)), "q1's")), "q1's")
})

# ── Duplicates ────────────────────────────────────────────────────────────────

test_that("duplicate names are made unique", {
  s <- setNames(one(1, 2), c("S", "S"))
  expect_warning(write_tmp(s), "Deduplicating")
  expect_equal(written_names(s), c("S", "S_1"))
})

test_that("three-way duplicates are all made unique", {
  s <- setNames(one(1, 2, 3), c("S", "S", "S"))
  expect_equal(written_names(s), c("S", "S_1", "S_2"))
})

test_that("dedup keeps names within the 31-character maximum", {
  # the suffix must shorten the base, not push the name over the limit
  nm <- strrep("x", 31)
  s <- setNames(one(1, 2), c(nm, nm))
  got <- written_names(s)
  expect_equal(nchar(got), c(31L, 31L))
  expect_equal(got[2L], paste0(strrep("x", 29), "_1"))
  expect_false(anyDuplicated(got) > 0L)
})

test_that("collisions created by sanitizing are caught", {
  # "a/b" and "a-b" both sanitize to "a-b"; only a dedup pass that runs
  # *after* sanitizing sees that collision
  s <- setNames(one(1, 2), c("a/b", "a-b"))
  got <- written_names(s)
  expect_equal(got, c("a-b", "a-b_1"))
  expect_false(anyDuplicated(got) > 0L)
})

test_that("collisions created by truncating are caught", {
  a <- paste0(strrep("x", 31), "AAA")
  b <- paste0(strrep("x", 31), "BBB")
  got <- written_names(setNames(one(1, 2), c(a, b)))
  expect_equal(nchar(got), c(31L, 31L))
  expect_false(anyDuplicated(got) > 0L)
})

# ── Empty and absent names ────────────────────────────────────────────────────

test_that("unnamed input is left for libxlsxwriter to name", {
  # a bare data frame has no sheet name, and must keep getting the default
  expect_equal(written_names(data.frame(a = 1)), "Sheet1")
  expect_equal(written_names(one(1, 2)), c("Sheet1", "Sheet2"))
})

test_that("an empty or NA name is left for libxlsxwriter to name", {
  expect_equal(written_names(setNames(list(data.frame(a = 1)), "")), "Sheet1")
  expect_equal(written_names(setNames(list(data.frame(a = 1)), NA)), "Sheet1")
  # mixed: the named one keeps its name
  expect_equal(written_names(setNames(one(1, 2), c("", "Data"))),
               c("Sheet1", "Data"))
})

# ── Warning content ───────────────────────────────────────────────────────────

test_that("a rename warning names the original and the replacement", {
  s <- setNames(list(data.frame(a = 1)), "2024/Q1")
  expect_warning(write_tmp(s), '"2024/Q1" -> "2024-Q1"', fixed = TRUE)
})

test_that("many renames are summarised rather than listed in full", {
  nms <- paste0("a/", seq_len(8))
  s <- setNames(one(1, 2, 3, 4, 5, 6, 7, 8), nms)
  expect_warning(write_tmp(s), "and 3 more", fixed = TRUE)
})

# ── The C-side backstop ───────────────────────────────────────────────────────

test_that("a name the R side failed to repair is rejected in C", {
  # drive the C validation directly by bypassing .resolve_sheet_names(); this
  # is the guard against a future R-side regression, so it must actually fire
  local_mocked_bindings(.resolve_sheet_names = function(nms, n) rep("a/b", n))
  expect_error(write_tmp(setNames(list(data.frame(a = 1)), "ok")),
               "cannot contain invalid characters")
})

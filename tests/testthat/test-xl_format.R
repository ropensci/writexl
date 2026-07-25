# Tests for the shared format engine (R/xl_format.R)

test_that("xl_color accepts names, hex, and integers", {
  expect_equal(xl_color("red"), 0xFF0000)
  expect_equal(xl_color("navy"), xl_color("#000080"))
  expect_equal(xl_color("#0000FF"), 0x0000FF)
  expect_equal(xl_color("00FF00"), 0x00FF00)   # bare 6-hex
  expect_equal(xl_color(255L), 255L)
  expect_equal(xl_color(0xFFFFFF), 0xFFFFFF)
  expect_identical(xl_color(NA), NA_integer_)
})

test_that("xl_color rejects bad input", {
  expect_error(xl_color("notacolor"), "Unknown color")
  expect_error(xl_color(c("red", "blue")), "single color")
  expect_error(xl_color(-1L), "range")
  expect_error(xl_color(0xFFFFFF + 1L), "range")
})

test_that("xl_font sets and validates properties", {
  f <- xl_font(name = "Arial", size = 14, color = "red", bold = TRUE,
               italic = TRUE, underline = "double", strikeout = TRUE,
               script = "super")
  fo <- unclass(f)$font
  expect_equal(fo$name, "Arial")
  expect_equal(fo$size, 14)
  expect_equal(fo$color, xl_color("red"))
  expect_true(fo$bold)
  expect_equal(fo$underline, "double")
  expect_equal(fo$script, "super")
  expect_true(is_xl_format(f))
})

test_that("xl_font handles rarely-used properties", {
  f <- xl_font(family = 2, charset = 1, outline = TRUE, shadow = TRUE,
               condense = TRUE, extend = TRUE, scheme = "minor", theme = 1,
               color_indexed = 8, font_only = TRUE)
  p <- writexl:::.xl_format_payload(f)
  expect_equal(p$font_family, 2L)
  expect_equal(p$font_charset, 1L)
  expect_true(p$font_outline)
  expect_true(p$font_shadow)
  expect_true(p$font_condense)
  expect_true(p$font_extend)
  expect_equal(p$font_scheme, "minor")
  expect_equal(p$theme, 1L)
  expect_equal(p$color_indexed, 8L)
  expect_true(p$font_only)
})

test_that("xl_font validation errors are typed and specific", {
  expect_error(xl_font(underline = "wiggly"), "underline")
  expect_error(xl_font(size = 0), "between 1 and 409")
  expect_error(xl_font(size = 500), "between 1 and 409")
  expect_error(xl_font(bold = "yes"), "bold")
  expect_error(xl_font(script = "middle"), "script")
  expect_error(xl_font(name = 1), "name")
  expect_error(xl_font(color = "notacolor"), "color")
})

test_that("empty xl_font is an empty format", {
  f <- xl_font()
  expect_true(is_xl_format(f))
  expect_null(unclass(f)$font)
})

test_that("xl_fill maps background/foreground/pattern and auto-solids", {
  f <- xl_fill(background = "yellow")
  expect_equal(unclass(f)$fill$pattern, "solid")   # auto solid
  f2 <- xl_fill(background = "yellow", foreground = "black", pattern = "light-gray")
  fi <- unclass(f2)$fill
  expect_equal(fi$pattern, "light-gray")            # explicit pattern preserved
  p <- writexl:::.xl_format_payload(f2)
  expect_equal(p$bg_color, xl_color("yellow"))
  expect_equal(p$fg_color, xl_color("black"))
  expect_equal(p$pattern, 4L)
  expect_error(xl_fill(pattern = "sparkle"), "pattern")
})

test_that("xl_border 'all' and 'color' fan out to four sides", {
  b <- xl_border(all = "thin", color = "gray")
  u <- unclass(b)$border
  expect_equal(u$left, "thin"); expect_equal(u$right, "thin")
  expect_equal(u$top, "thin");  expect_equal(u$bottom, "thin")
  expect_equal(u$left_color, xl_color("gray"))
  # per-side overrides 'all'
  b2 <- xl_border(all = "thin", top = "thick")
  expect_equal(unclass(b2)$border$top, "thick")
  expect_equal(unclass(b2)$border$bottom, "thin")
})

test_that("xl_border diagonal and payload", {
  b <- xl_border(diagonal = "up-down", diagonal_style = "hair",
                 diagonal_color = "red", bottom = "double", bottom_color = "blue")
  p <- writexl:::.xl_format_payload(b)
  expect_equal(p$diag_type, 3L)
  expect_equal(p$diag_border, 7L)
  expect_equal(p$diag_color, xl_color("red"))
  expect_equal(p$border_bottom, 6L)
  expect_equal(p$bottom_color, xl_color("blue"))
  expect_error(xl_border(all = "wavy"), "all")
  expect_error(xl_border(diagonal = "sideways"), "diagonal")
})

test_that("xl_align validates and maps alignment", {
  a <- xl_align(horizontal = "center", vertical = "top", wrap = TRUE,
                rotation = 45, indent = 2, shrink = TRUE, reading_order = "rtl")
  p <- writexl:::.xl_format_payload(a)
  expect_equal(p$align_h, 2L)
  expect_equal(p$align_v, 8L)
  expect_true(p$text_wrap)
  expect_equal(p$rotation, 45L)
  expect_equal(p$indent, 2L)
  expect_true(p$shrink)
  expect_equal(p$reading_order, 2L)
  expect_equal(writexl:::.xl_format_payload(xl_align(rotation = 270))$rotation, 270L)
  expect_error(xl_align(horizontal = "middle"), "horizontal")
  expect_error(xl_align(vertical = "sideways"), "vertical")
  expect_error(xl_align(rotation = 120), "270")
})

test_that("xl_num_format takes a string or an index", {
  expect_equal(writexl:::.xl_format_payload(xl_num_format("#,##0.00"))$num_format, "#,##0.00")
  expect_equal(writexl:::.xl_format_payload(xl_num_format(index = 3))$num_format_index, 3L)
  expect_error(xl_num_format(format = 1), "format")
  expect_error(xl_num_format(index = 999), "between 0 and 255")
})

test_that("xl_protection inverts locked and sets hidden", {
  p <- writexl:::.xl_format_payload(xl_protection(locked = FALSE, hidden = TRUE))
  expect_true(p$unlocked)
  expect_true(p$hidden)
  # locked = TRUE is Excel's default -> no unlocked call
  expect_null(writexl:::.xl_format_payload(xl_protection(locked = TRUE))$unlocked)
  expect_error(xl_protection(locked = "no"), "locked")
})

test_that("xl_format merges groups property-by-property, right-biased", {
  f <- xl_format(xl_font(bold = TRUE), xl_border(bottom = "thin"),
                 xl_num_format("#,##0.00"))
  u <- unclass(f)
  expect_true(u$font$bold)
  expect_equal(u$border$bottom, "thin")
  expect_equal(u$num_format$format, "#,##0.00")
  # accumulate within a group
  f2 <- xl_font(bold = TRUE) + xl_font(italic = TRUE)
  expect_true(unclass(f2)$font$bold)
  expect_true(unclass(f2)$font$italic)
  # right operand wins on conflict
  f3 <- xl_font(color = "red") + xl_font(color = "blue")
  expect_equal(unclass(f3)$font$color, xl_color("blue"))
})

test_that("xl_format scalar flags and NULL handling", {
  f <- xl_format(xl_font(bold = TRUE), quote_prefix = TRUE, hyperlink = TRUE)
  p <- writexl:::.xl_format_payload(f)
  expect_true(p$quote_prefix)
  expect_true(p$set_hyperlink)
  expect_true(is_xl_format(xl_format()))              # empty
  expect_true(is_xl_format(xl_format(NULL, xl_font(bold = TRUE), NULL)))
  expect_error(xl_format("not a format"), "must be xl_format")
})

test_that("+ operator handles NULL and rejects non-formats", {
  f <- xl_font(bold = TRUE)
  expect_identical(f + NULL, f)
  expect_identical(NULL + f, f)
  expect_error(f + 1, "must be xl_format")
})

test_that("print methods run for coverage", {
  expect_output(print(xl_font(bold = TRUE, color = "red")), "font")
  expect_output(print(xl_format()), "empty")
  expect_output(print(xl_format(xl_font(bold = TRUE), quote_prefix = TRUE)), "quote_prefix")
})

test_that("format registry deduplicates and reserves 0 for none", {
  reg <- writexl:::.new_format_registry()
  i1 <- writexl:::.register_format(reg, xl_font(bold = TRUE))
  i2 <- writexl:::.register_format(reg, xl_font(bold = TRUE))
  i3 <- writexl:::.register_format(reg, xl_font(italic = TRUE))
  i0 <- writexl:::.register_format(reg, NULL)
  iE <- writexl:::.register_format(reg, xl_format())    # empty -> 0
  expect_equal(i1, 1L)
  expect_equal(i2, 1L)
  expect_equal(i3, 2L)
  expect_equal(i0, 0L)
  expect_equal(iE, 0L)
  expect_length(reg$table, 2L)
})

test_that("full font payload round-trips every property", {
  f <- xl_font(name = "Arial", size = 12, color = "red", bold = TRUE,
               italic = TRUE, underline = "single", strikeout = TRUE,
               script = "sub")
  p <- writexl:::.xl_format_payload(f)
  expect_equal(p$font_name, "Arial")
  expect_equal(p$font_size, 12)
  expect_equal(p$font_color, xl_color("red"))
  expect_true(p$bold); expect_true(p$italic); expect_true(p$strikeout)
  expect_equal(p$underline, 1L)
  expect_equal(p$script, 2L)
})

test_that("full border payload round-trips every side and color", {
  b <- xl_border(left = "thin", right = "medium", top = "thick", bottom = "double",
                 left_color = "red", right_color = "green", top_color = "blue",
                 bottom_color = "black")
  p <- writexl:::.xl_format_payload(b)
  expect_equal(p$border_left, 1L)
  expect_equal(p$border_right, 2L)
  expect_equal(p$border_top, 5L)
  expect_equal(p$border_bottom, 6L)
  expect_equal(p$left_color, xl_color("red"))
  expect_equal(p$right_color, xl_color("green"))
  expect_equal(p$top_color, xl_color("blue"))
  expect_equal(p$bottom_color, xl_color("black"))
})

test_that("numeric validators reject non-numeric and non-scalar input", {
  expect_error(xl_font(size = "big"), "single number")
  expect_error(xl_font(size = c(1, 2)), "single number")
  expect_error(xl_font(family = "two"), "single number")
  expect_error(xl_align(indent = -1), "between 0")
  expect_error(xl_align(rotation = -200), "270")
})

test_that("merge_xl_format tolerates NULL operands directly", {
  f <- xl_font(bold = TRUE)
  expect_identical(writexl:::merge_xl_format(NULL, f), f)
  expect_identical(writexl:::merge_xl_format(f, NULL), f)
})

test_that(".fmt_kv tolerates NULL values", {
  expect_type(writexl:::.fmt_kv(list(a = NULL)), "character")
})

test_that("payload key is order-independent", {
  a <- writexl:::.payload_key(list(bold = TRUE, italic = TRUE))
  b <- writexl:::.payload_key(list(italic = TRUE, bold = TRUE))
  expect_equal(a, b)
  expect_equal(writexl:::.payload_key(list()), "")
})

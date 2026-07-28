# Images.  This file covers the constructor and its validation; what reaches
# the file is tested once apply_images() lands.

# A minimal valid 1x1 PNG.  Hardcoded rather than shipped as a fixture or drawn
# with png(), which needs a working graphics device that CRAN machines may not
# have.
PNG_1x1 <- as.raw(c(
  0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
  0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
  0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00,
  0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
  0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82))

png_file <- function() {
  p <- tempfile(fileext = ".png")
  writeBin(PNG_1x1, p)
  p
}

# ── Format detection ──────────────────────────────────────────────────────────

test_that("every format Excel reads is recognised by its leading bytes", {
  # enumerated rather than sampled: a format added to .IMAGE_MAGIC without a
  # case here fails the count check below
  expect_equal(.image_format(PNG_1x1), "png")
  expect_equal(.image_format(as.raw(c(0xFF, 0xD8, 0xFF, 0xE0))), "jpeg")
  expect_equal(.image_format(charToRaw("GIF89a")), "gif")
  expect_equal(.image_format(charToRaw("GIF87a")), "gif")
  expect_equal(.image_format(charToRaw("BM  ")), "bmp")
  expect_equal(length(.IMAGE_MAGIC), 4L)
})

test_that("anything else is not an image", {
  expect_null(.image_format(charToRaw("nope")))
  expect_null(.image_format(charToRaw("<?xml version=\"1.0\"?>")))
  expect_null(.image_format(raw(0)))
  # a truncated header must not match on a prefix
  expect_null(.image_format(as.raw(c(0x89, 0x50))))
})

test_that("the format comes from the bytes, not the extension", {
  # a .png that is really a JPEG would otherwise be misfiled in the zip
  p <- tempfile(fileext = ".png")
  writeBin(as.raw(c(0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10)), p)
  expect_equal(.check_image(p), "jpeg")
  # and a real image with no extension at all is still accepted
  q <- tempfile()
  writeBin(PNG_1x1, q)
  expect_equal(.check_image(q), "png")
})

test_that("a non-image file is refused, naming it", {
  p <- tempfile(fileext = ".png")
  writeLines("not an image", p)
  expect_error(xl_image(p), "is not a PNG, JPEG, GIF or BMP")
})

test_that("SVG is refused with a reason", {
  # libxlsxwriter declines SVG because Excel stores it converted to PNG, so the
  # error should say that rather than just listing the formats
  s <- tempfile(fileext = ".svg")
  writeLines("<svg xmlns=\"http://www.w3.org/2000/svg\"/>", s)
  expect_error(xl_image(s), "SVG is not supported")
  expect_error(xl_image(s), "save the plot as a PNG")
})

# ── The image itself ──────────────────────────────────────────────────────────

test_that("an image may be a path or a raw vector", {
  expect_s3_class(xl_image(png_file()), "xl_image")
  expect_s3_class(xl_image(PNG_1x1), "xl_image")
  expect_equal(unclass(xl_image(PNG_1x1))$image_format, "png")
})

test_that("a missing or unreadable source is refused", {
  expect_error(xl_image(), "must be an image path or a raw vector")
  expect_error(xl_image("no_such_file.png"), "file does not exist")
  expect_error(xl_image(tempdir()), "is a directory, not an image")
  expect_error(xl_image(raw(0)), "empty raw vector")
  expect_error(xl_image(42), "single file path or a raw vector")
})

# ── Options ───────────────────────────────────────────────────────────────────

test_that("scale and offset take one number or two", {
  expect_equal(unclass(xl_image(png_file(), scale = 0.5))$scale, c(0.5, 0.5))
  expect_equal(unclass(xl_image(png_file(), scale = c(2, 3)))$scale, c(2, 3))
  expect_equal(unclass(xl_image(png_file(), offset = c(4, 5)))$offset, c(4, 5))
  expect_null(unclass(xl_image(png_file()))$offset)

  expect_error(xl_image(png_file(), scale = -1), "`scale` must be positive")
  expect_error(xl_image(png_file(), scale = 0), "`scale` must be positive")
  expect_error(xl_image(png_file(), scale = c(1, 2, 3)), "one number or two")
  expect_error(xl_image(png_file(), scale = "big"), "one number or two")
  expect_error(xl_image(png_file(), offset = c(1, NA)), "one number or two")
})

test_that("every object position is accepted", {
  for (p in names(.LXW_OBJECT_POSITION))
    expect_equal(unclass(xl_image(png_file(), position = p))$position, p)
  expect_equal(length(.LXW_OBJECT_POSITION), 5L)
  expect_error(xl_image(png_file(), position = "wander"), "`position` must be one of")
})

test_that("a format without embed is refused rather than ignored", {
  # cell_format reaches libxlsxwriter only through worksheet_embed_image_opt(),
  # so accepting it for a floating image would be accepting a silent no-op
  expect_error(xl_image(png_file(), format = xl_font(bold = TRUE)),
               "needs `embed = TRUE`")
  expect_s3_class(xl_image(png_file(), embed = TRUE,
                           format = xl_font(bold = TRUE)), "xl_image")
  expect_error(xl_image(png_file(), embed = TRUE, format = "bold"),
               "must be an xl_format")
})

test_that("a url on an embedded image is refused", {
  # the image occupies the cell, so there is nowhere for the link to live
  expect_error(xl_image(png_file(), embed = TRUE, url = "https://example.com"),
               "not available for an embedded image")
  expect_s3_class(xl_image(png_file(), url = "https://example.com"), "xl_image")
})

test_that("text and flag arguments are validated", {
  expect_error(xl_image(png_file(), description = 1), "single non-NA string")
  expect_error(xl_image(png_file(), url = c("a", "b")), "single non-NA string")
  expect_error(xl_image(png_file(), tip = NA), "single non-NA string")
  expect_error(xl_image(png_file(), decorative = "yes"), "`decorative` must be")
  expect_error(xl_image(png_file(), embed = "yes"), "`embed` must be")
  # NA means "unset" throughout writexl, so it is accepted and drops out
  expect_null(unclass(xl_image(png_file(), decorative = NA))$decorative)
  expect_null(unclass(xl_image(png_file(), embed = NA))$embed)
})

test_that("images normalise from one or a list", {
  expect_length(.image_list(xl_image(png_file())), 1L)
  expect_length(.image_list(list(xl_image(png_file()), xl_image(PNG_1x1))), 2L)
  expect_length(.image_list(NULL), 0L)
  expect_error(.image_list("logo.png"), "must be an xl_image object")
  expect_error(.image_list(list(xl_image(PNG_1x1), "x")),
               "`image\\[\\[2\\]\\]` must be an xl_image object")
})

test_that("the print method runs", {
  expect_output(print(xl_image(png_file())), "xl_image")
  expect_output(print(xl_image(PNG_1x1)), "67 bytes")
  expect_output(print(xl_image(png_file(), embed = TRUE)), "embedded")
})

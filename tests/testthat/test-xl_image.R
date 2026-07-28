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
  # the argument is `image`, not `file`: it takes more than a filename
  expect_equal(names(formals(xl_image))[1L], "image")
})

test_that("a missing or unreadable source is refused", {
  expect_error(xl_image(), "must be a file path, a raw vector")
  expect_error(xl_image("no_such_file.png"), "file does not exist")
  expect_error(xl_image(tempdir()), "is a directory, not an image")
  expect_error(xl_image(raw(0)), "empty raw vector")
  expect_error(xl_image(42), "must be a file path, a raw vector")
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

# ── In-memory pictures ────────────────────────────────────────────────────────

# A PNG's real pixel size, read from its IHDR chunk -- no decoder needed.
png_dims <- function(bytes) {
  be <- function(i) sum(as.integer(bytes[i]) * c(256^3, 256^2, 256, 1))
  c(width = be(17:20), height = be(21:24))
}

test_that("a raster is encoded at its own pixel size", {
  skip_if_not(isTRUE(capabilities("png")))
  # dim(raster) is rows x cols, and a PNG is width x height -- so a 2x3 raster
  # must come out 3 wide and 2 tall, not the other way round
  r <- as.raster(matrix(c("red", "green", "blue", "black", "white", "grey"),
                        nrow = 2, byrow = TRUE))
  expect_equal(dim(r), c(2L, 3L))
  b <- unclass(xl_image(r))$image
  expect_true(is.raw(b))
  expect_equal(.image_format(b), "png")
  expect_equal(unname(png_dims(b)), c(3, 2))
})

test_that("everything as.raster() accepts is encoded", {
  skip_if_not(isTRUE(capabilities("png")))
  # colour matrix
  m <- matrix(c("#FF0000", "#00FF00", "#0000FF", "#FFFFFF"), nrow = 2)
  expect_equal(unname(png_dims(unclass(xl_image(m))$image)), c(2, 2))
  # numeric matrix (greyscale)
  g <- matrix(c(0, 0.5, 1, 0.25, 0.75, 0.1), nrow = 2)
  expect_equal(unname(png_dims(unclass(xl_image(g))$image)), c(3, 2))
  # RGB array
  a <- array(0.5, dim = c(4, 5, 3))
  expect_equal(unname(png_dims(unclass(xl_image(a))$image)), c(5, 4))
})

test_that("a nativeRaster is encoded", {
  skip_if_not(isTRUE(capabilities("png")))
  nr <- structure(as.integer(c(-1, -16777216, -65536, -16711936)),
                  dim = c(2L, 2L), class = "nativeRaster")
  expect_equal(unname(png_dims(unclass(xl_image(nr))$image)), c(2, 2))
})

test_that("an in-memory image is encoded once, at construction", {
  skip_if_not(isTRUE(capabilities("png")))
  # the object carries bytes, not the raster, so a later change to the raster
  # cannot alter what gets written
  r <- as.raster(matrix("red", 1, 1))
  im <- xl_image(r)
  expect_true(is.raw(unclass(im)$image))
  expect_equal(unclass(im)$image_format, "png")
})

test_that("a degenerate raster is refused", {
  skip_if_not(isTRUE(capabilities("png")))
  expect_error(xl_image(matrix(character(0), nrow = 0, ncol = 0)),
               "no pixels to write")
})

test_that("encoding a raster leaves no device open", {
  skip_if_not(isTRUE(capabilities("png")))
  before <- grDevices::dev.cur()
  xl_image(as.raster(matrix("red", 2, 2)))
  expect_equal(grDevices::dev.cur(), before)
})

# ── What reaches the file ─────────────────────────────────────────────────────

img_parts <- function(..., df = data.frame(a = 1:4)) {
  p <- write_tmp(list(S = xl_sheet(df, ...)))
  d <- tempfile(); dir.create(d); utils::unzip(p, exdir = d)
  part <- function(f) {
    q <- file.path(d, f)
    if (file.exists(q)) paste(readLines(q, warn = FALSE), collapse = "") else ""
  }
  list(files = list.files(d, recursive = TRUE),
       drawing = part("xl/drawings/drawing1.xml"),
       sheet = part("xl/worksheets/sheet1.xml"),
       rels = part("xl/drawings/_rels/drawing1.xml.rels"))
}

test_that("a floating image reaches the drawing and the media", {
  r <- img_parts(image = xl_image(png_file(), at = "C2"))
  expect_true("xl/media/image1.png" %in% r$files)
  expect_true("xl/drawings/drawing1.xml" %in% r$files)
  expect_match(r$sheet, "<drawing r:id=", fixed = TRUE)
  # anchored to C2, which is 0-based col 2, row 1
  expect_match(r$drawing, "<xdr:col>2</xdr:col>", fixed = TRUE)
  expect_match(r$drawing, "<xdr:row>1</xdr:row>", fixed = TRUE)
})

test_that("a path and a raw vector differ only in the default alt text", {
  a <- img_parts(image = xl_image(png_file(), at = "A1"))
  b <- img_parts(image = xl_image(PNG_1x1, at = "A1"))
  expect_equal(a$files, b$files)
  # Excel defaults an image's alt text to its file name, so the path form
  # carries a descr the raw form cannot have -- the rest must be identical
  expect_match(a$drawing, "descr=", fixed = TRUE)
  expect_no_match(b$drawing, "descr=", fixed = TRUE)
  strip <- function(x) sub(' descr="[^"]*"', "", x)
  expect_equal(strip(a$drawing), strip(b$drawing))
})

test_that("offsets and the object position are written", {
  r <- img_parts(image = xl_image(png_file(), at = "A1", offset = c(40, 20)))
  # Pixels become EMU at 9525 each.  The absolute offset is the stable thing to
  # assert: an offset larger than the row's height is normalised into the next
  # row, so a 20px y-offset on a 20px row comes out as row 1 with rowOff 0 --
  # the position is right, the spelling is not what a naive test expects.
  expect_match(r$drawing, sprintf('<a:off x="%d" y="%d"/>', 40 * 9525,
                                  20 * 9525), fixed = TRUE)
  expect_match(r$drawing, sprintf("<xdr:colOff>%d</xdr:colOff>", 40 * 9525),
               fixed = TRUE)
  # the position controls the anchor kind
  expect_match(img_parts(image = xl_image(png_file(), at = "A1",
                                          position = "dont_move_dont_size"))$drawing,
               "absoluteAnchor|editAs=\"absolute\"")
  expect_match(img_parts(image = xl_image(png_file(), at = "A1",
                                          position = "move_dont_size"))$drawing,
               "oneCellAnchor|editAs=\"oneCell\"")
})

test_that("a description is written, and decorative replaces it", {
  expect_match(img_parts(image = xl_image(png_file(), at = "A1",
                                          description = "A logo"))$drawing,
               "A logo", fixed = TRUE)
  # Excel does not write a description for a decorative image
  d <- img_parts(image = xl_image(png_file(), at = "A1",
                                  description = "A logo",
                                  decorative = TRUE))$drawing
  expect_no_match(d, "descr=\"A logo\"", fixed = TRUE)
})

test_that("a url becomes an external relationship", {
  r <- img_parts(image = xl_image(png_file(), at = "A1",
                                  url = "https://example.com", tip = "Home"))
  expect_match(r$rels, "https://example.com", fixed = TRUE)
  expect_match(r$rels, "TargetMode=\"External\"", fixed = TRUE)
})

test_that("several images share one drawing", {
  r <- img_parts(image = list(xl_image(png_file(), at = "A1"),
                              xl_image(png_file(), at = "A5"),
                              xl_image(PNG_1x1, at = "A9")))
  # the same picture three times is deduplicated to one media file
  expect_equal(sum(grepl("^xl/media/", r$files)), 1L)
  expect_equal(sum(grepl("^xl/drawings/drawing", r$files)), 1L)
  expect_equal(length(gregexpr("<xdr:pic>", r$drawing)[[1L]]), 3L)
})

test_that("an anchor must name a single cell", {
  expect_error(write_tmp(list(S = xl_sheet(data.frame(a = 1:4),
                                           image = xl_image(PNG_1x1, at = "A1:B2")))),
               "must name a single cell")
})

test_that("an embedded image writes a cell, not a drawing", {
  r <- img_parts(image = xl_image(PNG_1x1, at = "C2", embed = TRUE))
  expect_true("xl/media/image1.png" %in% r$files)
  expect_false("xl/drawings/drawing1.xml" %in% r$files)
  expect_true("xl/richData/richValueRel.xml" %in% r$files)
  # the cell carries the value-metadata index, and it must be in range
  expect_match(r$sheet, "vm=\"1\"", fixed = TRUE)
})

# ── Background ────────────────────────────────────────────────────────────────

test_that("a background image is stored and referenced", {
  r <- img_parts(background = png_file())
  expect_true("xl/media/image1.png" %in% r$files)
  expect_match(r$sheet, "<picture r:id=", fixed = TRUE)
  # a background is a backdrop, not a drawing
  expect_false("xl/drawings/drawing1.xml" %in% r$files)
})

test_that("a background may be a raw vector or an in-memory image", {
  expect_true("xl/media/image1.png" %in% img_parts(background = PNG_1x1)$files)
  skip_if_not(isTRUE(capabilities("png")))
  expect_true("xl/media/image1.png" %in%
                img_parts(background = as.raster(matrix("red", 2, 2)))$files)
})

test_that("a bad background is refused like any other image", {
  bad <- tempfile(fileext = ".png")
  writeLines("not an image", bad)
  expect_error(write_tmp(list(S = xl_sheet(data.frame(a = 1), background = bad))),
               "is not a PNG, JPEG, GIF or BMP")
})

# ── Guards against libxlsxwriter's counter bugs ───────────────────────────────

test_that("an embedded image may not share a workbook with any other image", {
  # _prepare_drawings() gives the embedded image the shared image ref id and
  # writes it as the cell's vm, which indexes the valueMetadata blocks instead
  df <- data.frame(a = 1:4)
  expect_error(write_tmp(list(
    A = xl_sheet(df, image = xl_image(PNG_1x1, at = "C2")),
    B = xl_sheet(df, image = xl_image(PNG_1x1, at = "C2", embed = TRUE)))),
    "points past the end of the metadata")
  # embedded images on their own are fine, in any number
  expect_silent(write_tmp(list(
    A = xl_sheet(df, image = xl_image(PNG_1x1, at = "C2", embed = TRUE)),
    B = xl_sheet(df, image = xl_image(PNG_1x1, at = "C4", embed = TRUE)))))
})

test_that("a drawing-less image sheet may not precede a floating image", {
  # header/footer and background images each consume a drawing id without
  # writing a drawing part, so a later floating image points at a missing file
  df <- data.frame(a = 1:4)
  logo <- png_file()
  hdr <- xl_sheet(df, page = xl_page_setup(header = "&L&G",
                                           header_image = list(left = logo)))
  expect_error(write_tmp(list(A = hdr,
                              B = xl_sheet(df, image = xl_image(logo, "C2")))),
               "earlier in the workbook has a header or footer image")
  expect_error(write_tmp(list(A = xl_sheet(df, background = logo),
                              B = xl_sheet(df, image = xl_image(logo, "C2")))),
               "earlier in the workbook has a background image")
  # the other order is fine, and is what the message recommends
  expect_silent(write_tmp(list(A = xl_sheet(df, image = xl_image(logo, "C2")),
                               B = hdr)))
  expect_silent(write_tmp(list(A = xl_sheet(df, image = xl_image(logo, "C2")),
                               B = xl_sheet(df, background = logo))))
})

test_that("only an embedded image costs row streaming", {
  df <- data.frame(a = 1:4)
  plan_float <- list(list(images = list(list(embed = 0L))))
  plan_embed <- list(list(images = list(list(embed = 1L))))
  expect_equal(.resolve_constant_memory(list(D = df), list(), plan_float)$on, 1L)
  cm <- .resolve_constant_memory(list(D = df), list(), plan_embed)
  expect_equal(cm$on, 0L)
  expect_match(cm$reasons, "embedded image")
})

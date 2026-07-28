# =============================================================================
# Images
# =============================================================================
#
# An image is either *inserted* -- floating over the sheet, anchored to a cell,
# which is what Excel has always done -- or *embedded* in a cell, which is the
# Excel 365 "place in cell" feature.  Older Excel shows #VALUE! where an
# embedded image sits, so the two are not interchangeable and `embed` is an
# explicit choice rather than a default.
#
# The image itself may be a path or a raw vector.  A raw vector is the common
# case in R, where a plot has just been written by a graphics device, and it
# maps to libxlsxwriter's _buffer variants.
#
# Format is detected from the leading bytes rather than the file extension: a
# ".png" that is really a JPEG would otherwise reach libxlsxwriter and be
# misfiled in the zip.  Excel supports PNG, JPEG, BMP and GIF; libxlsxwriter
# declines SVG deliberately, since Excel converts it to PNG on the way in.
# -----------------------------------------------------------------------------

.LXW_OBJECT_POSITION <- c(
  "default" = 0L, move_and_size = 1L, move_dont_size = 2L,
  dont_move_dont_size = 3L, move_and_size_after = 4L
)

# The formats Excel accepts, by their leading bytes.
.IMAGE_MAGIC <- list(
  png  = as.raw(c(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)),
  jpeg = as.raw(c(0xFF, 0xD8, 0xFF)),
  gif  = as.raw(c(0x47, 0x49, 0x46, 0x38)),
  bmp  = as.raw(c(0x42, 0x4D))
)

# Identify an image from its first bytes, or NULL if it is none of them.
.image_format <- function(bytes) {
  for (nm in names(.IMAGE_MAGIC)) {
    m <- .IMAGE_MAGIC[[nm]]
    if (length(bytes) >= length(m) &&
        identical(bytes[seq_along(m)], m))
      return(nm)
  }
  NULL
}

# The first bytes of an image given as a path or a raw vector.
.image_head <- function(x, n = 8L) {
  if (is.raw(x)) return(utils::head(x, n))
  con <- file(x, "rb")
  on.exit(close(con))
  readBin(con, "raw", n)
}

# Check that an image is one Excel can read, and say what it is otherwise.
.check_image <- function(x, arg = "file") {
  if (is.raw(x)) {
    if (!length(x))
      stop(sprintf("`%s` is an empty raw vector", arg), call. = FALSE)
  } else {
    if (!is.character(x) || length(x) != 1L || is.na(x) || !nzchar(x))
      stop(sprintf("`%s` must be a single file path or a raw vector", arg),
           call. = FALSE)
    if (!file.exists(x))
      stop(sprintf("`%s`: file does not exist: %s", arg, x), call. = FALSE)
    if (dir.exists(x))
      stop(sprintf("`%s` is a directory, not an image: %s", arg, x),
           call. = FALSE)
  }
  fmt <- .image_format(.image_head(x))
  if (is.null(fmt)) {
    what <- if (is.raw(x)) "the raw vector" else sprintf("\"%s\"", x)
    hint <- if (!is.raw(x) && grepl("\\.svgz?$", x, ignore.case = TRUE))
      paste0(" SVG is not supported: Excel converts it to PNG on the way in, ",
             "so save the plot as a PNG instead.")
    else ""
      stop(sprintf(paste0("`%s`: %s is not a PNG, JPEG, GIF or BMP -- those ",
                          "are the formats Excel reads.%s"), arg, what, hint),
           call. = FALSE)
  }
  fmt
}

#' Insert an image into a worksheet
#'
#' @description
#' `xl_image()` places an image on a sheet, either floating over the cells and
#' anchored to one of them (the default) or, with `embed = TRUE`, inside a cell.
#' Pass one or a list of them as `xl_sheet(image = )`.
#'
#' The image may be a file path or a raw vector, which is convenient when a plot
#' has just been written by a graphics device and never touched the disk.  PNG,
#' JPEG, GIF and BMP are supported --- the formats Excel reads --- and the
#' format is detected from the file's own bytes rather than its extension.
#'
#' @section Floating versus embedded:
#' An inserted image floats above the grid: it has a position but occupies no
#' cell, and `position` decides whether it moves and resizes as rows and columns
#' change.  An embedded image (`embed = TRUE`) lives *in* a cell and sizes with
#' it, which is the Excel 365 "place in cell" behaviour.  Excel versions without
#' that feature show `#VALUE!` in place of an embedded image, so it is worth
#' choosing deliberately.
#'
#' @section Caveats inherited from Excel:
#' An image's scaling can shift if it crosses a row whose height changed --- for
#' a taller font or wrapped text --- so set the height explicitly with
#' [xl_row_spec()] for rows an image spans.  BMP is supported only for backward
#' compatibility and must be 24-bit true colour; prefer PNG.  SVG is refused,
#' because Excel stores it converted to PNG anyway.
#'
#' @param file The image: a path to a PNG, JPEG, GIF or BMP, or a raw vector
#'   holding one.
#' @param at The cell the image is anchored to, such as `"B2"`, or a
#'   `list(rows = , cols = )` spec selecting a single cell.
#' @param scale Scale factor: one number for both axes, or `c(x, y)`.
#' @param offset Offset from the anchor cell's top-left corner in pixels, as
#'   `c(x, y)`.
#' @param position How the image behaves when rows and columns change size:
#'   `"move_and_size"` (the default), `"move_dont_size"`,
#'   `"dont_move_dont_size"`, `"move_and_size_after"`, or `"default"` for
#'   Excel's own default.  Ignored when `embed = TRUE`.
#' @param description Alt text, for screen readers.  Excel defaults it to the
#'   file name; `""` writes none.
#' @param decorative Mark the image as decorative, so screen readers skip it.
#'   Excel does not write a description for a decorative image.
#' @param url An optional hyperlink the image links to.
#' @param tip An optional mouseover tip for `url`.
#' @param embed Place the image inside the cell rather than floating above it.
#'   Requires a version of Excel that supports images in cells.
#' @param format An [xl_format] for the cell holding an embedded image.  Only
#'   meaningful with `embed = TRUE`.
#' @return An `xl_image` object.
#' @family writexl
#' @seealso [xl_sheet]
#' @export
#' @examples
#' logo <- system.file("help", "figures", "logo.png", package = "writexl")
#' if (nzchar(logo)) {
#'   xl_image(logo, at = "C2", scale = 0.5)
#'   xl_image(logo, at = "C2", url = "https://example.com", tip = "Home")
#' }
xl_image <- function(file, at = "A1", scale = 1, offset = NULL,
                     position = "move_and_size", description = NULL,
                     decorative = FALSE, url = NULL, tip = NULL,
                     embed = FALSE, format = NULL) {
  if (missing(file) || is.null(file))
    stop("`file` must be an image path or a raw vector", call. = FALSE)
  fmt <- .check_image(file, "file")

  str1 <- function(x, arg) {
    if (is.null(x)) return(NULL)
    if (!is.character(x) || length(x) != 1L || is.na(x))
      stop(sprintf("`%s` must be a single non-NA string", arg), call. = FALSE)
    x
  }
  pair <- function(x, arg, what) {
    if (is.null(x)) return(NULL)
    if (!is.numeric(x) || !length(x) %in% c(1L, 2L) || anyNA(x))
      stop(sprintf("`%s` must be one number or two, %s", arg, what),
           call. = FALSE)
    if (length(x) == 1L) x <- c(x, x)
    as.numeric(x)
  }

  sc <- pair(scale, "scale", "as c(x, y)")
  if (!is.null(sc) && any(sc <= 0))
    stop("`scale` must be positive", call. = FALSE)
  off <- pair(offset, "offset", "of pixels as c(x, y)")

  pos <- .val_enum(position, names(.LXW_OBJECT_POSITION), "position")
  dec <- .val_flag(decorative, "decorative")
  emb <- .val_flag(embed, "embed")
  if (!is.null(format) && !is_xl_format(format))
    stop("`format` must be an xl_format object", call. = FALSE)
  # cell_format reaches libxlsxwriter only through worksheet_embed_image_opt(),
  # so accepting it for a floating image would be accepting a no-op
  if (!is.null(format) && !isTRUE(emb))
    stop("`format` applies to the cell an embedded image sits in, so it needs ",
         "`embed = TRUE`", call. = FALSE)
  if (isTRUE(emb) && !is.null(url))
    stop("`url` is not available for an embedded image; Excel puts the image ",
         "in the cell, leaving nowhere for the link", call. = FALSE)

  structure(.drop_null(list(
    file = file, image_format = fmt, at = at, scale = sc, offset = off,
    position = pos, description = str1(description, "description"),
    decorative = dec, url = str1(url, "url"), tip = str1(tip, "tip"),
    embed = emb, format = format
  )), class = "xl_image")
}

#' @export
print.xl_image <- function(x, ...) {
  p <- unclass(x)
  src <- if (is.raw(p[["file"]]))
    sprintf("<%s, %d bytes>", p[["image_format"]], length(p[["file"]]))
  else basename(p[["file"]])
  cat(sprintf("<xl_image: %s at %s%s>\n", src,
              if (is.character(p[["at"]])) p[["at"]] else "a cell",
              if (isTRUE(p[["embed"]])) ", embedded" else ""))
  invisible(x)
}

# Normalise one image or a list of them to a checked list.
.image_list <- function(image, arg = "image") {
  if (is.null(image)) return(list())
  is_one <- inherits(image, "xl_image")
  ims <- if (is_one) list(image) else image
  if (!is.list(ims))
    stop(sprintf("`%s` must be an xl_image object or a list of them", arg),
         call. = FALSE)
  for (i in seq_along(ims))
    if (!inherits(ims[[i]], "xl_image"))
      stop(sprintf("`%s[[%d]]` must be an xl_image object", arg, i),
           call. = FALSE)
  ims
}

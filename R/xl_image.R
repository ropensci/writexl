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
# An image may arrive in four shapes: a file path, a raw vector of encoded
# bytes, or an in-memory picture as either a `raster` (or anything as.raster()
# accepts -- a colour matrix, a numeric matrix, an RGB/RGBA array) or a
# `nativeRaster`.  The last two are what rasterImage() draws, so anything that
# can be plotted can be written.
#
# xlsx stores encoded images, so an in-memory picture has to become PNG bytes
# first.  Base R has no PNG encoder outside its graphics devices, so the
# conversion renders the raster through grDevices::png() at exactly its own
# pixel dimensions.  That needs capabilities("png"), which is not guaranteed on
# a headless build -- when it is missing the error says how to do the
# conversion by hand rather than failing obscurely.
#
# Paths and raw vectors map to libxlsxwriter's file and _buffer variants
# respectively; a converted raster becomes a raw vector.
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

# Is this an in-memory picture rather than a path or encoded bytes?
.is_raster_like <- function(x) {
  inherits(x, "nativeRaster") || inherits(x, "raster") ||
    is.matrix(x) || (is.array(x) && length(dim(x)) == 3L)
}

# Render an in-memory picture to PNG bytes, at exactly its own pixel size.
#
# grDevices and graphics are Suggests rather than Imports: writexl needs them
# only to encode an in-memory image, and only this one path uses them, so the
# package keeps its zero-dependency footing and the requirement is checked
# where it arises.
.raster_to_png <- function(x, arg) {
  missing_ns <- !vapply(c("grDevices", "graphics"), requireNamespace,
                        logical(1), quietly = TRUE)
  if (any(missing_ns) || !isTRUE(capabilities("png")))
    stop(sprintf(paste0("`%s` is an in-memory image, which writexl encodes ",
                        "with R's PNG device -- unavailable here (%s).\n",
                        "  Encode it yourself and pass the result instead:\n",
                        "    xl_image(png::writePNG(img))            # raw ",
                        "vector\n",
                        "    png::writePNG(img, \"plot.png\")           # or a ",
                        "file\n",
                        "    xl_image(\"plot.png\")"),
                 arg,
                 if (any(missing_ns))
                   paste("missing package(s):",
                         paste(names(missing_ns)[missing_ns], collapse = ", "))
                 else "capabilities(\"png\") is FALSE"),
         call. = FALSE)
  d <- dim(x)
  if (length(d) < 2L || anyNA(d[1:2]) || any(d[1:2] < 1L))
    stop(sprintf("`%s` has no pixels to write", arg), call. = FALSE)
  height <- d[1L]
  width  <- d[2L]

  path <- tempfile(fileext = ".png")
  # png() warns that a small pixel size is "unlikely", meaning it suspects the
  # caller wanted inches.  Here the size is the image's own, so the warning is
  # noise -- muted by message, not blanket-suppressed, so anything else the
  # device has to say still comes through.
  withCallingHandlers(
    grDevices::png(path, width = width, height = height, units = "px",
                   bg = "transparent"),
    warning = function(w) {
      if (grepl("unlikely values in pixels", conditionMessage(w), fixed = TRUE))
        invokeRestart("muffleWarning")
    })
  # dev.off() must run even if drawing fails, or the device is left open
  on.exit({
    if (grDevices::dev.cur() > 1L) grDevices::dev.off()
    unlink(path)
  }, add = TRUE)
  graphics::par(mar = c(0, 0, 0, 0), xaxs = "i", yaxs = "i")
  graphics::plot.new()
  graphics::plot.window(c(0, 1), c(0, 1))
  graphics::rasterImage(x, 0, 0, 1, 1, interpolate = FALSE)
  grDevices::dev.off()
  readBin(path, "raw", file.size(path))
}

# Check that an image is one Excel can read, and say what it is otherwise.
.check_image <- function(x, arg = "image") {
  if (is.raw(x)) {
    if (!length(x))
      stop(sprintf("`%s` is an empty raw vector", arg), call. = FALSE)
  } else {
    if (!is.character(x) || length(x) != 1L || is.na(x) || !nzchar(x))
      stop(sprintf(paste0("`%s` must be a file path, a raw vector of encoded ",
                          "bytes, or an image rasterImage() can draw"), arg),
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
#' @param image The image, in any of four shapes: a path to a PNG, JPEG, GIF or
#'   BMP; a raw vector holding one of those encoded; a `raster` (or anything
#'   [grDevices::as.raster()] accepts, such as a colour matrix or an RGB/RGBA
#'   array); or a `nativeRaster`.  The last two are what
#'   [graphics::rasterImage()] draws, so anything you can plot can be written.
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
xl_image <- function(image, at = "A1", scale = 1, offset = NULL,
                     position = "move_and_size", description = NULL,
                     decorative = FALSE, url = NULL, tip = NULL,
                     embed = FALSE, format = NULL) {
  if (missing(image) || is.null(image))
    stop("`image` must be a file path, a raw vector, or an image ",
         "rasterImage() can draw", call. = FALSE)
  # an in-memory picture is encoded now, so a later failure cannot be blamed on
  # a file that was fine when it was named
  if (.is_raster_like(image)) image <- .raster_to_png(image, "image")
  fmt <- .check_image(image, "image")

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
    image = image, image_format = fmt, at = at, scale = sc, offset = off,
    position = pos, description = str1(description, "description"),
    decorative = dec, url = str1(url, "url"), tip = str1(tip, "tip"),
    embed = emb, format = format
  )), class = "xl_image")
}

#' @export
print.xl_image <- function(x, ...) {
  p <- unclass(x)
  src <- if (is.raw(p[["image"]]))
    sprintf("<%s, %d bytes>", p[["image_format"]], length(p[["image"]]))
  else basename(p[["image"]])
  cat(sprintf("<xl_image: %s at %s%s>\n", src,
              if (is.character(p[["at"]])) p[["at"]] else "a cell",
              if (isTRUE(p[["embed"]])) ", embedded" else ""))
  invisible(x)
}

# Resolve a sheet's images to the C payloads.  Applied after the row loop
# beside the merges and tables: an embedded image writes a cell, and a floating
# one anchors to a row whose height the row loop may still be setting.
.resolve_images <- function(el, df, reg, header_offset, props) {
  if (!inherits(el, "xl_sheet")) return(list())
  ims <- .image_list(el$image)
  lapply(seq_along(ims), function(i) {
    p <- unclass(ims[[i]])
    at <- .xl_resolve_range(p[["at"]], arg = sprintf("image[[%d]] at", i),
                            df = df, header_offset = header_offset,
                            allow_cell = TRUE)
    if (at[1L] != at[3L] || at[2L] != at[4L])
      stop(sprintf(paste0("`image[[%d]] at` must name a single cell, not a ",
                          "range: an image is anchored to one cell"), i),
           call. = FALSE)
    ent <- list(row = at[1L], col = at[2L],
                embed = as.integer(isTRUE(p[["embed"]])))
    if (is.raw(p[["image"]])) ent$buffer <- p[["image"]]
    else                      ent$filename <- p[["image"]]
    if (!is.null(p[["scale"]])) {
      ent$x_scale <- p[["scale"]][1L]
      ent$y_scale <- p[["scale"]][2L]
    }
    if (!is.null(p[["offset"]])) {
      ent$x_offset <- as.integer(p[["offset"]][1L])
      ent$y_offset <- as.integer(p[["offset"]][2L])
    }
    if (!is.null(p[["position"]]))
      ent$object_position <- unname(.LXW_OBJECT_POSITION[[p[["position"]]]])
    if (!is.null(p[["description"]])) ent$description <- p[["description"]]
    if (isTRUE(p[["decorative"]]))    ent$decorative <- 1L
    if (!is.null(p[["url"]]))         ent$url <- p[["url"]]
    if (!is.null(p[["tip"]]))         ent$tip <- p[["tip"]]
    if (!is.null(p[["format"]]))
      ent$cell_format_id <- .register_format(reg,
        merge_xl_format(props$default_format, p[["format"]]))
    ent
  })
}

# Resolve xl_sheet(background =) to the C payload.  A background takes the same
# shapes as an image but no options: Excel tiles it to fill the sheet, and it
# is a display-only backdrop that never prints.
.resolve_background <- function(el, arg = "background") {
  if (!inherits(el, "xl_sheet") || is.null(el$background)) return(NULL)
  x <- el$background
  if (.is_raster_like(x)) x <- .raster_to_png(x, arg)
  .check_image(x, arg)
  if (is.raw(x)) list(buffer = x) else list(filename = x)
}

# --- Guard against libxlsxwriter's drawing-id desync ---------------------------
#
# _prepare_drawings() in workbook.c increments drawing_id for any sheet holding
# images, embedded images, charts, header/footer VML or a background image, and
# uses it to build the sheet's "../drawings/drawingN.xml" relationship target.
# The packager numbers the drawing *files* separately, counting only sheets that
# actually have a drawing (_get_drawing_count() in packager.c).
#
# A sheet whose only image is embedded, or lives in a header or footer, is
# therefore counted without producing a drawing file, and every floating image
# on a LATER sheet gets a target one too high.  Excel finds the missing part and
# repairs the file, dropping the image.
#
# writexl cannot renumber the relationship without patching vendored code, so
# the combination is refused.  Ordering the sheets so the floating images come
# first is a real workaround, which is why the message says so.
.check_drawing_order <- function(sheets, sheet_names) {
  has_floating <- function(s)
    any(vapply(s$images, function(i) !isTRUE(as.logical(i$embed)), logical(1)))
  has_embedded <- function(s)
    any(vapply(s$images, function(i) isTRUE(as.logical(i$embed)), logical(1)))
  has_hf_image <- function(s)
    any(grepl("_image_(left|center|right)$", names(s$page)))
  has_background <- function(s) !is.null(s$background)

  # 1. An embedded image alongside any other kind.
  #
  # _prepare_drawings() gives an embedded image the shared image_ref_id and
  # passes it to worksheet_set_error_cell(), which writes it as the cell's `vm`
  # attribute.  But `vm` indexes the <valueMetadata> blocks, of which there is
  # one per *embedded* image.  With only embedded images in the workbook the two
  # counters move together and the file is right; any other image shifts the
  # ref id and `vm` then points past the end of the metadata, which Excel
  # repairs away.  The richValueRel target goes to the wrong image for the same
  # reason.
  emb <- vapply(sheets, has_embedded, logical(1))
  other <- vapply(sheets, function(s) has_floating(s) || has_hf_image(s),
                  logical(1))
  if (any(emb) && any(other))
    stop(sprintf(paste0("sheet \"%s\" has an embedded image and sheet \"%s\" ",
                        "has another image.\n  libxlsxwriter numbers the ",
                        "embedded image's metadata from a counter shared with ",
                        "every other image, so mixing them writes a cell ",
                        "reference that points past the end of the metadata ",
                        "and Excel repairs the sheet.\n  Use `embed = TRUE` ",
                        "only in a workbook whose images are all embedded, or ",
                        "make this one a floating image."),
                 sheet_names[which(emb)[1L]], sheet_names[which(other)[1L]]),
         call. = FALSE)

  # 2. A sheet counted for a drawing id that produces no drawing file, before a
  # sheet with a floating image.
  offender <- NULL
  for (i in seq_along(sheets)) {
    s <- sheets[[i]]
    if (!is.null(offender) && has_floating(s))
      stop(sprintf(paste0("sheet \"%s\" has a floating image, but sheet \"%s\" ",
                          "earlier in the workbook has %s.\n  libxlsxwriter ",
                          "numbers the drawing relationship wrongly in that ",
                          "order, and Excel repairs the file and drops the ",
                          "image.\n  Put the sheet(s) with floating images ",
                          "before \"%s\"."),
                   sheet_names[i], offender$name, offender$what,
                   offender$name), call. = FALSE)
    if (is.null(offender) && !has_floating(s)) {
      what <- if (has_hf_image(s)) "a header or footer image"
              else if (has_background(s)) "a background image"
              else NULL
      if (!is.null(what)) offender <- list(name = sheet_names[i], what = what)
    }
  }
  invisible(NULL)
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

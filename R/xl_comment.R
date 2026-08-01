# =============================================================================
# xl_comment: Excel cell comments (a.k.a. notes)
# =============================================================================
#
# A comment is a per-cell overlay, independent of the cell's value/formula/
# hyperlink.  Excel comment boxes support only a background color and a font
# name/size/family -- no bold, italic, text color, borders, number formats or
# alignment.  So a comment reuses `xl_format` as its *styling input* (via the
# `format` argument), extracts the four supported properties, and warns about
# anything else through `.check_comment_format()`.
# -----------------------------------------------------------------------------

#' Test whether an object is an `xl_comment`
#' @param x An object.
#' @return `TRUE` if `x` inherits from `"xl_comment"`.
#' @family cell content
#' @export
is_xl_comment <- function(x) inherits(x, "xl_comment")

#' Create a cell comment
#'
#' @description
#' `xl_comment()` builds a comment (note) that can be attached to a cell via
#' [xl_cell_general()]'s `comment` argument.  The simplest form is just text
#' (`xl_cell_general(value = x, comment = "see note")`); `xl_comment()` adds
#' options such as the author, initial visibility, box size and position, and
#' styling.
#'
#' Excel comment boxes support only a **background color** and a **font
#' name/size/family**.  These are supplied by reusing the formatting engine via
#' the `format` argument (e.g. `xl_font(name = "Arial", size = 10) +
#' xl_fill(background = "lightyellow")`); any other format property (bold,
#' borders, number formats, ...) is not supported by comments and triggers a
#' warning.
#'
#' @param value A single string: the comment's text.  Anything with an
#'   [as.character()] method is accepted, so a cell built for a sheet can be
#'   reused here.
#' @param format An optional [xl_format]; only its fill background color and
#'   font name/size/family are used (see [xl_fill()], [xl_font()]).  Other
#'   properties are unsupported and warned about.
#' @param author Comment author (shown in Excel's status bar).  Defaults to the
#'   sheet/workbook `comment_author` when unset (see [xl_sheet()],
#'   [xl_properties()]).
#' @param visible Initial visibility: `NA` follows the sheet default (comments
#'   are hidden unless `show_comments` is set), `TRUE` shows this comment,
#'   `FALSE` hides it.
#' @param width_pixels,height_pixels Comment box size in pixels (defaults
#'   128 x 74).
#' @param x_scale,y_scale Box scale factors.
#' @param start_row,start_col Zero-based anchor cell of the box (by default a
#'   comment sits one row up and one column right of its cell).
#' @param x_offset,y_offset Pixel offset of the box from its anchor.
#' @return An `xl_comment` object.
#' @family cell content
#' @seealso [xl_cell_general], [xl_format]
#' @export
#' @examples
#' # plain text
#' xl_comment("Double-check this figure")
#'
#' # with author and styling reused from the format engine
#' xl_comment("Estimate", author = "Finance",
#'            format = xl_font(name = "Arial", size = 10) +
#'                     xl_fill(background = "lightyellow"))
xl_comment <- function(value, format = NULL, author = NA, visible = NA,
                       width_pixels = NA, height_pixels = NA, x_scale = NA,
                       y_scale = NA, start_row = NA, start_col = NA,
                       x_offset = NA, y_offset = NA) {
  value <- .as_display_string(value, "value", null_ok = FALSE)

  # styling: validate now, but keep the xl_format object itself so it stays
  # easy to inspect and modify (e.g. cm$format <- cm$format + xl_font(size = 9)).
  # It is flattened to the fields libxlsxwriter wants only on the way to C, in
  # .comment_c_payload().
  if (!is.null(format)) {
    if (!is_xl_format(format))
      stop("`format` must be an xl_format object", call. = FALSE)
    .check_comment_format(format)
  }

  vis <- .val_flag(visible, "visible")
  visible_code <- if (is.null(vis)) NULL else if (isTRUE(vis)) 2L else 1L

  payload <- .drop_null(list(
    value       = value,
    format      = format,
    author      = .val_str(author, "author"),
    visible     = visible_code,
    width       = .val_int(width_pixels, "width_pixels", min = 0),
    height      = .val_int(height_pixels, "height_pixels", min = 0),
    x_scale     = .val_num(x_scale, "x_scale", min = 0),
    y_scale     = .val_num(y_scale, "y_scale", min = 0),
    start_row   = .val_int(start_row, "start_row", min = 0),
    start_col   = .val_int(start_col, "start_col", min = 0),
    x_offset    = .val_int(x_offset, "x_offset"),
    y_offset    = .val_int(y_offset, "y_offset")
  ))
  structure(payload, class = "xl_comment")
}

# Flatten a comment payload for C: replace the `format` object with the four
# fields Excel comments actually support.  Called on the final pass to C so the
# xl_format stays editable up to that point.
.comment_c_payload <- function(cm) {
  if (is.null(cm)) return(NULL)
  fmt <- cm$format
  cm$format <- NULL
  if (!is.null(fmt)) {
    f <- unclass(fmt)
    cm$color       <- f$fill$background
    cm$font_name   <- f$font$name
    cm$font_size   <- f$font$size
    cm$font_family <- f$font$family
  }
  .drop_null(cm)
}

#' @export
print.xl_comment <- function(x, ...) {
  p <- unclass(x)
  cat("<xl_comment>\n")
  cat("  value:", p$value, "\n")
  for (o in setdiff(names(p), c("value", "format")))
    cat(sprintf("  %s: %s\n", o, paste(format(p[[o]]), collapse = ",")))
  if (!is.null(p$format)) print(p$format)
  invisible(x)
}

# Warn about xl_format properties Excel comments cannot render.  Supported:
# fill$background, font$name/size/family.  fill$pattern is ignored only when
# "solid" (the artifact xl_fill() adds for a background); a non-solid pattern is
# a real, unsupported request and is warned.  Everything else is warned.
.check_comment_format <- function(format) {
  f <- unclass(format)
  bad <- character(0)
  if (!is.null(f$font)) {
    extra <- setdiff(names(f$font), c("name", "size", "family"))
    if (length(extra)) bad <- c(bad, paste0("font$", extra))
  }
  if (!is.null(f$fill)) {
    extra <- setdiff(names(f$fill), c("background", "pattern"))
    if (length(extra)) bad <- c(bad, paste0("fill$", extra))
    if (!is.null(f$fill$pattern) && !identical(f$fill$pattern, "solid"))
      bad <- c(bad, "fill$pattern")
  }
  for (g in c("border", "align", "num_format", "protection"))
    if (!is.null(f[[g]])) bad <- c(bad, g)
  if (isTRUE(f$quote_prefix)) bad <- c(bad, "quote_prefix")
  if (isTRUE(f$hyperlink))    bad <- c(bad, "hyperlink")
  if (length(bad))
    warning("xl_comment(): these format properties are not supported by Excel ",
            "comments and will be ignored: ", paste(bad, collapse = ", "),
            call. = FALSE)
  invisible(bad)
}

# Normalise one per-cell comment element to a C-ready payload (or NULL).
# Accepts NA (no comment), a single string (default-option comment), or an
# xl_comment.
.comment_payload <- function(x) {
  if (is.null(x)) return(NULL)
  if (is.character(x) && length(x) == 1L && is.na(x)) return(NULL)
  if (length(x) == 1L && !is.list(x) && is.na(x)) return(NULL)
  if (is_xl_comment(x)) return(unclass(x))
  if (is.character(x) && length(x) == 1L) return(unclass(xl_comment(x)))
  stop("each comment must be NA, a single string, or an xl_comment object",
       call. = FALSE)
}

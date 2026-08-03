# =============================================================================
# xl_format: the shared cell/row/column/workbook formatting engine
# =============================================================================
#
# An `xl_format` is a simple S3 object: a list with a class.  Each of its slots
# is either absent (unset) or a validated named list describing one group of
# formatting properties:
#
#   font        - font_name/size/color/bold/italic/underline/...
#   fill        - background/foreground/pattern
#   border      - per-side styles and colors, diagonal
#   align       - horizontal/vertical/wrap/rotation/indent/shrink/reading_order
#   num_format  - number-format string or built-in index
#   protection  - locked/hidden
#
# plus the scalar slots `quote_prefix` and `hyperlink`.
#
# Human-readable values (enum strings, R colors) are stored in the object so it
# prints legibly; conversion to the C-ready integer payload libxlsxwriter needs
# happens once, at write time, in `.xl_format_payload()`.
# -----------------------------------------------------------------------------

# Enum string -> libxlsxwriter integer maps.  These MUST match the values in
# src/libxlsxwriter/include/xlsxwriter/format.h.
.LXW <- list(
  underline = c("none" = 0L, "single" = 1L, "double" = 2L,
                "single-accounting" = 3L, "double-accounting" = 4L),
  script    = c("super" = 1L, "sub" = 2L),
  align_h   = c("none" = 0L, "left" = 1L, "center" = 2L, "right" = 3L,
                "fill" = 4L, "justify" = 5L, "center-across" = 6L,
                "distributed" = 7L),
  align_v   = c("top" = 8L, "bottom" = 9L, "center" = 10L, "justify" = 11L,
                "distributed" = 12L),
  pattern   = c("none" = 0L, "solid" = 1L, "medium-gray" = 2L, "dark-gray" = 3L,
                "light-gray" = 4L, "dark-horizontal" = 5L, "dark-vertical" = 6L,
                "dark-down" = 7L, "dark-up" = 8L, "dark-grid" = 9L,
                "dark-trellis" = 10L, "light-horizontal" = 11L,
                "light-vertical" = 12L, "light-down" = 13L, "light-up" = 14L,
                "light-grid" = 15L, "light-trellis" = 16L, "gray-125" = 17L,
                "gray-0625" = 18L),
  border    = c("none" = 0L, "thin" = 1L, "medium" = 2L, "dashed" = 3L,
                "dotted" = 4L, "thick" = 5L, "double" = 6L, "hair" = 7L,
                "medium-dashed" = 8L, "dash-dot" = 9L, "medium-dash-dot" = 10L,
                "dash-dot-dot" = 11L, "medium-dash-dot-dot" = 12L,
                "slant-dash-dot" = 13L),
  diagonal  = c("up" = 1L, "down" = 2L, "up-down" = 3L),
  reading_order = c("default" = 0L, "ltr" = 1L, "rtl" = 2L)
)

# --- validation helpers ------------------------------------------------------
# Each returns the coerced value, or NULL when the property is unset (NA).

.is_unset <- function(x) is.null(x) || (length(x) == 1L && is.na(x))

.val_flag <- function(x, arg) {
  if (.is_unset(x)) return(NULL)
  if (!is.logical(x) || length(x) != 1L)
    stop(sprintf("`%s` must be a single logical (TRUE/FALSE) or NA", arg), call. = FALSE)
  x
}

.val_enum <- function(x, choices, arg) {
  if (.is_unset(x)) return(NULL)
  if (!is.character(x) || length(x) != 1L || !x %in% choices)
    stop(sprintf("`%s` must be one of: %s (or NA)",
                 arg, paste(choices, collapse = ", ")), call. = FALSE)
  x
}

.val_str <- function(x, arg) {
  if (.is_unset(x)) return(NULL)
  if (!is.character(x) || length(x) != 1L)
    stop(sprintf("`%s` must be a single string or NA", arg), call. = FALSE)
  x
}

# What to write where a value has none.  Any single atomic value is kept with
# its own type, so `na = 0` stays a number and `na = "-"` a string -- the cell
# writer that handles ordinary values handles this one, which is why no type
# table is needed here.  NA returns NULL, meaning "nothing set at this level":
# a workbook then leaves the cell blank, and a column or a cell falls through
# to the level above it.
.val_na <- function(x, arg = "na") {
  if (is.null(x)) return(NULL)
  if (is.factor(x)) x <- as.character(x)
  if (!is.atomic(x) || length(x) != 1L)
    stop(sprintf(paste0("`%s` must be a single value of any type, or NA to ",
                        "leave the cell blank"), arg), call. = FALSE)
  # NaN is as unrepresentable as NA, so it cannot stand in for one either
  if (is.na(x)) return(NULL)
  x
}

.val_int <- function(x, arg, min = NA, max = NA) {
  if (.is_unset(x)) return(NULL)
  if (!is.numeric(x) || length(x) != 1L)
    stop(sprintf("`%s` must be a single number or NA", arg), call. = FALSE)
  x <- as.integer(round(x))
  if (!is.na(min) && x < min || !is.na(max) && x > max)
    stop(sprintf("`%s` must be between %s and %s", arg, min, max), call. = FALSE)
  x
}

.val_num <- function(x, arg, min = NA, max = NA) {
  if (.is_unset(x)) return(NULL)
  if (!is.numeric(x) || length(x) != 1L)
    stop(sprintf("`%s` must be a single number or NA", arg), call. = FALSE)
  x <- as.numeric(x)
  if (!is.na(min) && x < min || !is.na(max) && x > max)
    stop(sprintf("`%s` must be between %s and %s", arg, min, max), call. = FALSE)
  x
}

.val_color <- function(x, arg) {
  if (.is_unset(x)) return(NULL)
  tryCatch(xl_color(x),
           error = function(e)
             stop(sprintf("`%s`: %s", arg, conditionMessage(e)), call. = FALSE))
}

# Build a named list, dropping NULL entries.
.drop_null <- function(x) x[!vapply(x, is.null, logical(1))]

#' Normalize a color to a libxlsxwriter RGB integer
#'
#' Converts an R color name (e.g. `"navy"`), a hex string (`"#FF0000"` or
#' `"FF0000"`), or an integer in the range `0x000000`..`0xFFFFFF` into the
#' single `0xRRGGBB` integer that libxlsxwriter expects.  Used internally by
#' every `color`/`background`/`foreground` argument of the formatting
#' constructors, and exported so it can be used directly.
#'
#' @param x A single color: an R color name, a hex string, or an integer.
#'   `NA` returns `NA_integer_` (an unset color).
#' @return A single integer in `0x000000`..`0xFFFFFF`, or `NA_integer_`.
#' @family cell formatting
#' @export
#' @examples
#' xl_color("red")
#' xl_color("#0000FF")
#' xl_color(255L)
xl_color <- function(x) {
  if (length(x) != 1L)
    stop("xl_color() expects a single color, got length ", length(x), call. = FALSE)
  if (is.na(x)) return(NA_integer_)
  if (is.numeric(x)) {
    v <- as.integer(round(x))
    if (v < 0L || v > 0xFFFFFFL)
      stop("Numeric color must be in the range 0x000000..0xFFFFFF", call. = FALSE)
    return(v)
  }
  x <- as.character(x)
  if (grepl("^[0-9A-Fa-f]{6}$", x)) x <- paste0("#", x)
  rgb <- tryCatch(grDevices::col2rgb(x),
                  error = function(e) stop("Unknown color: '", x, "'", call. = FALSE))
  as.integer(rgb[1L] * 65536L + rgb[2L] * 256L + rgb[3L])
}

# --- internal constructor ----------------------------------------------------

new_xl_format <- function(font = NULL, fill = NULL, border = NULL, align = NULL,
                          num_format = NULL, protection = NULL,
                          quote_prefix = NULL, hyperlink = NULL,
                          class = character(0)) {
  slots <- .drop_null(list(
    font = font, fill = fill, border = border, align = align,
    num_format = num_format, protection = protection,
    quote_prefix = quote_prefix, hyperlink = hyperlink
  ))
  structure(slots, class = c(class, "xl_format"))
}

#' Test whether an object is an `xl_format`
#' @param x An object.
#' @return `TRUE` if `x` inherits from `"xl_format"`.
#' @family cell formatting
#' @export
is_xl_format <- function(x) inherits(x, "xl_format")

# =============================================================================
# Group constructors.  Each returns a complete `xl_format` with only its own
# group populated, so any one can be used directly as a cell/row/column format.
# Every property defaults to NA (= leave unset).
# =============================================================================

#' Cell formatting groups
#'
#' @description
#' These constructors build [xl_format] objects, each populating one group of
#' Excel cell-formatting properties.  Every property defaults to `NA`, meaning
#' "leave unset"; unset properties are simply not written.  Enum-like arguments
#' take lowercase strings and are validated against a fixed set of choices.
#'
#' Each constructor returns a full `xl_format`, so a single group can be used on
#' its own (`xl_cell_general(1, format = xl_font(bold = TRUE))`), and groups can
#' be combined with `+` (see [xl_format]).
#'
#' @param name Font name, e.g. `"Calibri"`.
#' @param size Font size in points (1--409).
#' @param color,background,foreground,left_color,right_color,top_color,bottom_color,diagonal_color
#'   A color: an R color name, a hex string, or an integer (see [xl_color]).
#' @param bold,italic,strikeout,outline,shadow,condense,extend,font_only Logical
#'   font flags.
#' @param underline One of `"none"`, `"single"`, `"double"`,
#'   `"single-accounting"`, `"double-accounting"`.
#' @param script One of `"super"`, `"sub"`.
#' @param family,charset,theme,color_indexed Advanced integer font properties
#'   (rarely needed).
#' @param scheme Font scheme string (advanced).
#' @return An [xl_format] object.
#' @name xl_format_groups
#' @family cell formatting
#' @seealso [xl_format], [xl_color]
#' @examples
#' xl_font(bold = TRUE, color = "navy", size = 12)
#' xl_fill(background = "#FFF2CC")
#' xl_border(all = "thin", color = "gray")
#' xl_align(horizontal = "center", vertical = "top", wrap = TRUE)
#' xl_num_format("#,##0.00")
#' xl_protection(locked = FALSE)
NULL

#' @rdname xl_format_groups
#' @export
xl_font <- function(bold = NA, italic = NA, color = NA, size = NA, name = NA,
                    underline = NA, strikeout = NA, script = NA, family = NA,
                    charset = NA, outline = NA, shadow = NA, condense = NA,
                    extend = NA, scheme = NA, theme = NA, color_indexed = NA,
                    font_only = NA) {
  font <- .drop_null(list(
    name          = .val_str(name, "name"),
    size          = .val_num(size, "size", min = 1, max = 409),
    color         = .val_color(color, "color"),
    bold          = .val_flag(bold, "bold"),
    italic        = .val_flag(italic, "italic"),
    underline     = .val_enum(underline, names(.LXW$underline), "underline"),
    strikeout     = .val_flag(strikeout, "strikeout"),
    script        = .val_enum(script, names(.LXW$script), "script"),
    family        = .val_int(family, "family"),
    charset       = .val_int(charset, "charset"),
    outline       = .val_flag(outline, "outline"),
    shadow        = .val_flag(shadow, "shadow"),
    condense      = .val_flag(condense, "condense"),
    extend        = .val_flag(extend, "extend"),
    scheme        = .val_str(scheme, "scheme"),
    theme         = .val_int(theme, "theme"),
    color_indexed = .val_int(color_indexed, "color_indexed"),
    font_only     = .val_flag(font_only, "font_only")
  ))
  new_xl_format(font = if (length(font)) font else NULL)
}

#' @rdname xl_format_groups
#' @param pattern Fill pattern, e.g. `"solid"`, `"light-gray"` (18 choices).
#'   When `background` is supplied and `pattern` is unset, a `"solid"` pattern
#'   is assumed.
#' @export
xl_fill <- function(background = NA, foreground = NA, pattern = NA,
                    transparency = NA) {
  bg  <- .val_color(background, "background")
  fg  <- .val_color(foreground, "foreground")
  pat <- .val_enum(pattern, names(.LXW$pattern), "pattern")
  if (is.null(pat) && !is.null(bg) && is.null(fg)) pat <- "solid"
  fill <- .drop_null(list(background = bg, foreground = fg, pattern = pat,
                          transparency = .val_int(transparency, "transparency",
                                                  min = 0, max = 100)))
  new_xl_format(fill = if (length(fill)) fill else NULL)
}

#' @rdname xl_format_groups
#' @param transparency Percentage transparency, 0--100.  **Charts only**: Excel
#'   has no transparency for a cell's fill or border, so this is ignored
#'   everywhere except a chart's line and fill (see [xl_chart()]).
#' @param all Border style applied to all four sides at once (one of `"none"`,
#'   `"thin"`, `"medium"`, `"dashed"`, `"dotted"`, `"thick"`, `"double"`,
#'   `"hair"`, `"medium-dashed"`, `"dash-dot"`, `"medium-dash-dot"`,
#'   `"dash-dot-dot"`, `"medium-dash-dot-dot"`, `"slant-dash-dot"`).
#' @param left,right,top,bottom Per-side border styles (override `all`).
#' @param diagonal Diagonal border direction: `"up"`, `"down"`, `"up-down"`.
#' @param diagonal_style Diagonal border style (same choices as `all`).
#' @export
xl_border <- function(all = NA, left = NA, right = NA, top = NA, bottom = NA,
                      color = NA, left_color = NA, right_color = NA,
                      top_color = NA, bottom_color = NA, diagonal = NA,
                      diagonal_style = NA, diagonal_color = NA,
                      transparency = NA) {
  bchoice <- names(.LXW$border)
  all_s   <- .val_enum(all, bchoice, "all")
  all_c   <- .val_color(color, "color")
  border <- .drop_null(list(
    left           = .val_enum(left, bchoice, "left")   %||% all_s,
    right          = .val_enum(right, bchoice, "right")  %||% all_s,
    top            = .val_enum(top, bchoice, "top")      %||% all_s,
    bottom         = .val_enum(bottom, bchoice, "bottom") %||% all_s,
    left_color     = .val_color(left_color, "left_color")     %||% all_c,
    right_color    = .val_color(right_color, "right_color")   %||% all_c,
    top_color      = .val_color(top_color, "top_color")       %||% all_c,
    bottom_color   = .val_color(bottom_color, "bottom_color") %||% all_c,
    diagonal       = .val_enum(diagonal, names(.LXW$diagonal), "diagonal"),
    diagonal_style = .val_enum(diagonal_style, bchoice, "diagonal_style"),
    diagonal_color = .val_color(diagonal_color, "diagonal_color"),
    transparency   = .val_int(transparency, "transparency", min = 0, max = 100)
  ))
  new_xl_format(border = if (length(border)) border else NULL)
}

#' @rdname xl_format_groups
#' @param horizontal Horizontal alignment: `"left"`, `"center"`, `"right"`,
#'   `"fill"`, `"justify"`, `"center-across"`, `"distributed"`.
#' @param vertical Vertical alignment: `"top"`, `"bottom"`, `"center"`,
#'   `"justify"`, `"distributed"`.
#' @param wrap Logical; wrap text in the cell.
#' @param rotation Text rotation in degrees (-90..90, or 270).
#' @param indent Integer indentation level.
#' @param shrink Logical; shrink text to fit.
#' @param reading_order One of `"default"`, `"ltr"`, `"rtl"`.
#' @export
xl_align <- function(horizontal = NA, vertical = NA, wrap = NA, rotation = NA,
                     indent = NA, shrink = NA, reading_order = NA) {
  align <- .drop_null(list(
    horizontal    = .val_enum(horizontal, names(.LXW$align_h), "horizontal"),
    vertical      = .val_enum(vertical, names(.LXW$align_v), "vertical"),
    wrap          = .val_flag(wrap, "wrap"),
    rotation      = .val_int(rotation, "rotation", min = -90, max = 270),
    indent        = .val_int(indent, "indent", min = 0),
    shrink        = .val_flag(shrink, "shrink"),
    reading_order = .val_enum(reading_order, names(.LXW$reading_order), "reading_order")
  ))
  if (!is.null(align$rotation) && align$rotation > 90 && align$rotation != 270)
    stop("`rotation` must be between -90 and 90, or exactly 270", call. = FALSE)
  new_xl_format(align = if (length(align)) align else NULL)
}

#' @rdname xl_format_groups
#' @param format A number-format string, e.g. `"#,##0.00"`, `"0%"`,
#'   `"yyyy-mm-dd"`.
#' @param index An Excel built-in number-format index (alternative to
#'   `format`).
#' @export
xl_num_format <- function(format = NA, index = NA) {
  num <- .drop_null(list(
    format = .val_str(format, "format"),
    index  = .val_int(index, "index", min = 0, max = 255)
  ))
  new_xl_format(num_format = if (length(num)) num else NULL)
}

#' @rdname xl_format_groups
#' @param locked Logical; whether the cell is locked (Excel's default is
#'   `TRUE`).  Only takes effect when the worksheet is protected.  `FALSE`
#'   unlocks the cell.
#' @param hidden Logical; hide the cell's formula.
#' @export
xl_protection <- function(locked = NA, hidden = NA) {
  prot <- .drop_null(list(
    locked = .val_flag(locked, "locked"),
    hidden = .val_flag(hidden, "hidden")
  ))
  new_xl_format(protection = if (length(prot)) prot else NULL)
}

`%||%` <- function(a, b) if (is.null(a)) b else a

# =============================================================================
# xl_format(): merge any number of group objects into one format.
# =============================================================================

#' Combine cell-formatting groups into a single format
#'
#' @description
#' `xl_format()` merges any number of [xl_format] objects (typically the
#' single-group objects returned by [xl_font], [xl_fill], [xl_border],
#' [xl_align], [xl_num_format] and [xl_protection]) into one combined format.
#' The `+` operator does the same for two formats.
#'
#' Merging is right-biased and works property-by-property: where two formats set
#' the *same* property the later one wins, but properties set by only one side
#' are all preserved.  So partial groups accumulate rather than overwrite.
#'
#' @param ... [xl_format] objects (and/or `NULL`s, which are ignored).  May also
#'   include the scalar flags `quote_prefix` and `hyperlink`.
#' @param quote_prefix Logical; treat the cell contents as literal text (as if
#'   prefixed with a single quote in Excel).
#' @param hyperlink Logical; apply the internal hyperlink style flag (advanced).
#' @return A combined [xl_format] object.
#' @family cell formatting
#' @seealso [xl_font], [xl_fill], [xl_border], [xl_align], [xl_num_format],
#'   [xl_protection], [xl_color]
#' @export
#' @examples
#' xl_format(xl_font(bold = TRUE), xl_border(bottom = "thin"),
#'           xl_num_format("#,##0.00"))
#'
#' # equivalent, using +
#' xl_font(bold = TRUE) + xl_border(bottom = "thin") + xl_num_format("#,##0.00")
xl_format <- function(..., quote_prefix = NA, hyperlink = NA) {
  parts <- list(...)
  parts <- parts[!vapply(parts, is.null, logical(1))]
  for (p in parts)
    if (!is_xl_format(p))
      stop("all `...` arguments to xl_format() must be xl_format objects ",
           "(e.g. from xl_font(), xl_fill(), ...)", call. = FALSE)
  out <- if (length(parts)) Reduce(merge_xl_format, parts) else new_xl_format()
  qp <- .val_flag(quote_prefix, "quote_prefix")
  hl <- .val_flag(hyperlink, "hyperlink")
  if (!is.null(qp)) out <- merge_xl_format(out, new_xl_format(quote_prefix = qp))
  if (!is.null(hl)) out <- merge_xl_format(out, new_xl_format(hyperlink = hl))
  out
}

# Merge b onto a: group-by-group then property-by-property, b wins on conflict.
# Subclass "extras" carried in the `xl_geometry`/`xl_target` attributes are
# merged too, and the result takes the more-specific class of the two.
merge_xl_format <- function(a, b) {
  if (is.null(a)) return(b)
  if (is.null(b)) return(a)
  out <- unclass(a)
  lb  <- unclass(b)
  for (slot in names(lb)) {
    # groups (named lists) merge property-by-property; scalar flags replace
    out[[slot]] <- if (is.list(lb[[slot]]) && is.list(out[[slot]]))
      utils::modifyList(out[[slot]], lb[[slot]]) else lb[[slot]]
  }
  # subclass extras (geometry/targeting) carried as attributes
  for (at in c("xl_geometry", "xl_target")) {
    av <- attr(a, at, exact = TRUE)
    bv <- attr(b, at, exact = TRUE)
    merged <- if (is.null(bv)) av else if (is.null(av)) bv else utils::modifyList(av, bv)
    if (!is.null(merged)) attr(out, at) <- merged
  }
  class(out) <- .result_format_class(a, b)
  out
}

# The more-specific of two format classes (a subclass such as xl_col_spec wins
# over a bare xl_format; if both are subclasses, the right operand wins).
.result_format_class <- function(a, b) {
  ca <- class(a); cb <- class(b)
  spec_b <- setdiff(cb, "xl_format")
  spec_a <- setdiff(ca, "xl_format")
  if (length(spec_b)) cb else if (length(spec_a)) ca else "xl_format"
}

#' @export
#' @rdname xl_format
#' @param e1,e2 [xl_format] objects to combine.
`+.xl_format` <- function(e1, e2) {
  if (is.null(e1)) return(e2)
  if (is.null(e2)) return(e1)
  if (!is_xl_format(e1) || !is_xl_format(e2))
    stop("both operands of `+` must be xl_format objects", call. = FALSE)
  merge_xl_format(e1, e2)
}

# =============================================================================
# Printing
# =============================================================================

#' @export
print.xl_format <- function(x, ...) {
  cls <- setdiff(class(x), "xl_format")
  label <- if (length(cls)) cls[1L] else "xl_format"
  cat(sprintf("<%s>\n", label))
  f <- unclass(x)
  tgt <- attr(x, "xl_target", exact = TRUE)
  geo <- attr(x, "xl_geometry", exact = TRUE)
  if (!is.null(tgt)) cat("  target:", .fmt_kv(tgt), "\n")
  if (!is.null(geo)) cat("  geometry:", .fmt_kv(geo), "\n")
  shown <- FALSE
  for (slot in names(f)) {
    v <- f[[slot]]
    if (is.list(v)) {
      cat(sprintf("  %s: %s\n", slot, .fmt_kv(v))); shown <- TRUE
    } else if (isTRUE(v)) {
      cat(sprintf("  %s: TRUE\n", slot)); shown <- TRUE
    }
  }
  if (!shown && is.null(tgt) && is.null(geo))
    cat("  (empty)\n")
  invisible(x)
}

.fmt_kv <- function(lst) {
  paste(vapply(names(lst), function(nm) {
    v <- lst[[nm]]
    if (is.null(v)) return(NA_character_)
    paste0(nm, "=", paste(format(v), collapse = ","))
  }, character(1)), collapse = ", ")
}

# =============================================================================
# Payload conversion + global dedup registry (used at write time)
# =============================================================================

# Flatten an xl_format to a named list of C-ready scalars.  Keys correspond
# 1:1 to format_set_* calls in src/write_xlsx.c.  Enum strings become the
# libxlsxwriter integer; colors are already integers; flags are TRUE.
.xl_format_payload <- function(fmt) {
  f <- unclass(fmt)
  p <- list()
  fo <- f$font
  if (!is.null(fo)) {
    if (!is.null(fo$name))          p$font_name     <- as.character(fo$name)
    if (!is.null(fo$size))          p$font_size     <- as.numeric(fo$size)
    if (!is.null(fo$color))         p$font_color    <- as.integer(fo$color)
    if (isTRUE(fo$bold))            p$bold          <- TRUE
    if (isTRUE(fo$italic))          p$italic        <- TRUE
    if (!is.null(fo$underline))     p$underline     <- .LXW$underline[[fo$underline]]
    if (isTRUE(fo$strikeout))       p$strikeout     <- TRUE
    if (!is.null(fo$script))        p$script        <- .LXW$script[[fo$script]]
    if (!is.null(fo$family))        p$font_family   <- as.integer(fo$family)
    if (!is.null(fo$charset))       p$font_charset  <- as.integer(fo$charset)
    if (isTRUE(fo$outline))         p$font_outline  <- TRUE
    if (isTRUE(fo$shadow))          p$font_shadow   <- TRUE
    if (isTRUE(fo$condense))        p$font_condense <- TRUE
    if (isTRUE(fo$extend))          p$font_extend   <- TRUE
    if (!is.null(fo$scheme))        p$font_scheme   <- as.character(fo$scheme)
    if (!is.null(fo$theme))         p$theme         <- as.integer(fo$theme)
    if (!is.null(fo$color_indexed)) p$color_indexed <- as.integer(fo$color_indexed)
    if (isTRUE(fo$font_only))       p$font_only     <- TRUE
  }
  fi <- f$fill
  if (!is.null(fi)) {
    if (!is.null(fi$background)) p$bg_color <- as.integer(fi$background)
    if (!is.null(fi$foreground)) p$fg_color <- as.integer(fi$foreground)
    if (!is.null(fi$pattern))    p$pattern  <- .LXW$pattern[[fi$pattern]]
  }
  bd <- f$border
  if (!is.null(bd)) {
    if (!is.null(bd$left))           p$border_left   <- .LXW$border[[bd$left]]
    if (!is.null(bd$right))          p$border_right  <- .LXW$border[[bd$right]]
    if (!is.null(bd$top))            p$border_top    <- .LXW$border[[bd$top]]
    if (!is.null(bd$bottom))         p$border_bottom <- .LXW$border[[bd$bottom]]
    if (!is.null(bd$left_color))     p$left_color    <- as.integer(bd$left_color)
    if (!is.null(bd$right_color))    p$right_color   <- as.integer(bd$right_color)
    if (!is.null(bd$top_color))      p$top_color     <- as.integer(bd$top_color)
    if (!is.null(bd$bottom_color))   p$bottom_color  <- as.integer(bd$bottom_color)
    if (!is.null(bd$diagonal))       p$diag_type     <- .LXW$diagonal[[bd$diagonal]]
    if (!is.null(bd$diagonal_style)) p$diag_border   <- .LXW$border[[bd$diagonal_style]]
    if (!is.null(bd$diagonal_color)) p$diag_color    <- as.integer(bd$diagonal_color)
  }
  al <- f$align
  if (!is.null(al)) {
    if (!is.null(al$horizontal))    p$align_h       <- .LXW$align_h[[al$horizontal]]
    if (!is.null(al$vertical))      p$align_v       <- .LXW$align_v[[al$vertical]]
    if (isTRUE(al$wrap))            p$text_wrap     <- TRUE
    if (!is.null(al$rotation))      p$rotation      <- as.integer(al$rotation)
    if (!is.null(al$indent))        p$indent        <- as.integer(al$indent)
    if (isTRUE(al$shrink))          p$shrink        <- TRUE
    if (!is.null(al$reading_order)) p$reading_order <- .LXW$reading_order[[al$reading_order]]
  }
  nf <- f$num_format
  if (!is.null(nf)) {
    if (!is.null(nf$format)) p$num_format       <- as.character(nf$format)
    if (!is.null(nf$index))  p$num_format_index <- as.integer(nf$index)
  }
  pr <- f$protection
  if (!is.null(pr)) {
    if (isFALSE(pr$locked)) p$unlocked <- TRUE
    if (isTRUE(pr$hidden))  p$hidden   <- TRUE
  }
  if (isTRUE(f$quote_prefix)) p$quote_prefix  <- TRUE
  if (isTRUE(f$hyperlink))    p$set_hyperlink <- TRUE
  p
}

# A stable string key for a payload, used to deduplicate identical formats.
.payload_key <- function(p) {
  if (length(p) == 0L) return("")
  p <- p[order(names(p))]
  paste0(names(p), ":",
         vapply(p, function(v) paste0(typeof(v), "=", as.character(v)), character(1)),
         collapse = "|")
}

# A registry accumulates unique format payloads across a whole workbook and
# hands back a 1-based integer id for each (0 means "no format").
.new_format_registry <- function() {
  reg <- new.env(parent = emptyenv())
  reg$keys  <- character(0)
  reg$table <- list()
  reg
}

.register_format <- function(reg, fmt) {
  if (is.null(fmt)) return(0L)
  p <- .xl_format_payload(fmt)
  if (length(p) == 0L) return(0L)
  k <- .payload_key(p)
  idx <- match(k, reg$keys)
  if (is.na(idx)) {
    reg$table[[length(reg$table) + 1L]] <- p
    reg$keys <- c(reg$keys, k)
    idx <- length(reg$table)
  }
  idx
}

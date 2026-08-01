# =============================================================================
# xl_format -> chart line / fill / pattern / font
# =============================================================================
#
# libxlsxwriter styles a chart through four small structs, and offers a
# _set_line / _set_fill / _set_pattern / _set_font function for each of fourteen
# targets -- series, data labels, markers, error bars, trendlines, the axis, its
# name and number fonts, major and minor gridlines, the legend, the title, the
# plot area, the chart area and the data table.  That is ~29 C functions
# expressing one idea: apply a format to a named part of the chart.
#
# So writexl exposes no chart-specific format type.  An ordinary xl_format is
# translated here, and the caller passes it to whichever chart argument it
# belongs to.
#
# The translation is lossy in known ways, and every loss is an error rather
# than a silent drop:
#
#   * A chart shape has no number format, no alignment and no cell protection,
#     so those groups are refused.  (An axis and data labels *do* take a number
#     format, but as their own argument -- it is not part of a shape's style.)
#   * lxw_chart_line has colour, dash and transparency but no *width*, so the
#     border styles that encode a width -- medium, thick, double, hair, and the
#     medium-* variants -- have no representation.  Refused, rather than
#     flattened to a plain line.
#   * A chart has one line, not four sides, so a border with differing sides is
#     refused rather than one side being picked.
# -----------------------------------------------------------------------------

# A colour on its way into a chart struct.  Every chart colour is tested for
# truth -- `if (line->color)` in _chart_write_a_ln(), and the same for fills and
# fonts -- so plain black, which xl_color() gives as 0, reads as "no colour set"
# and the part is drawn with Excel's default instead.  libxlsxwriter's own enum
# spells black LXW_COLOR_BLACK = 0x1000000 for exactly this reason, and
# LXW_COLOR_MASK strips the sentinel bit again on the way out.
.chart_color <- function(x) {
  if (is.null(x)) return(NULL)
  i <- as.integer(x)
  if (identical(i, 0L)) 0x1000000L else i
}

# Cell border style -> chart dash type.  Only exact equivalents are listed;
# .chart_dash() refuses the rest rather than approximating them.
.CHART_DASH <- c(
  thin       = 0L,   # LXW_CHART_LINE_DASH_SOLID
  dashed     = 3L,   # DASH
  "dash-dot" = 4L,   # DASH_DOT
  dotted     = 8L    # DOT
)

# Border styles a chart line cannot express, and why.
.CHART_DASH_UNSUPPORTED <- c(
  medium = "width", thick = "width", double = "width", hair = "width",
  "medium-dashed" = "width", "medium-dash-dot" = "width",
  "medium-dash-dot-dot" = "width", "dash-dot-dot" = "pattern",
  "slant-dash-dot" = "pattern"
)

.chart_dash <- function(style, arg) {
  if (is.null(style) || identical(style, "none")) return(NULL)
  if (style %in% names(.CHART_DASH)) return(unname(.CHART_DASH[[style]]))
  why <- .CHART_DASH_UNSUPPORTED[[style]]
  reason <- if (identical(why, "width"))
    paste0("a chart line has a colour, a dash pattern and a transparency, but ",
           "no width")
  else "Excel has no chart dash pattern matching it"
  stop(sprintf(paste0("`%s`: border style \"%s\" cannot be drawn on a chart -- ",
                      "%s.\n  Chart lines accept: %s (or \"none\" for no ",
                      "line)."),
               arg, style, reason,
               paste(sprintf("\"%s\"", names(.CHART_DASH)), collapse = ", ")),
       call. = FALSE)
}

# Cell fill pattern -> chart pattern.  A cell's patterns are ECMA-376's
# ST_PatternValue and a chart's are ST_PresetPatternVal: two vocabularies, so
# only the names that mean the same thing in both are listed, and
# .chart_pattern() refuses the rest rather than approximating them.  The grey
# screens map by the percentage their names state; 12.5% and 6.25% have no
# chart member, and the chart's grid and trellis sets have no dark/light pair.
.CHART_PATTERN <- c(
  "light-gray"       = 4L,   # LXW_CHART_PATTERN_PERCENT_25
  "medium-gray"      = 7L,   # PERCENT_50
  "dark-gray"        = 10L,  # PERCENT_75
  "light-down"       = 13L,  # LIGHT_DOWNWARD_DIAGONAL
  "light-up"         = 14L,  # LIGHT_UPWARD_DIAGONAL
  "dark-down"        = 15L,  # DARK_DOWNWARD_DIAGONAL
  "dark-up"          = 16L,  # DARK_UPWARD_DIAGONAL
  "light-vertical"   = 19L,  # LIGHT_VERTICAL
  "light-horizontal" = 20L,  # LIGHT_HORIZONTAL
  "dark-vertical"    = 23L,  # DARK_VERTICAL
  "dark-horizontal"  = 24L   # DARK_HORIZONTAL
)

.chart_pattern <- function(pattern, arg) {
  if (pattern %in% names(.CHART_PATTERN))
    return(unname(.CHART_PATTERN[[pattern]]))
  stop(sprintf(paste0("`%s`: fill pattern \"%s\" cannot be drawn on a chart -- ",
                      "Excel's chart patterns are a different set from its ",
                      "cell patterns, and none of them means the same thing.",
                      "\n  Chart patterns accept: %s."),
               arg, pattern,
               paste(sprintf("\"%s\"", names(.CHART_PATTERN)), collapse = ", ")),
       call. = FALSE)
}

# The groups a chart shape cannot express at all.  Alignment is the partial
# case: a chart shape has no horizontal or vertical placement for its text, but
# lxw_chart_font does carry a rotation, so xl_align(rotation =) turns a header
# cell and an axis label by the same angle.
.CHART_FORMAT_REFUSED <- c(
  num_format = paste0("a chart shape has no number format; set it on the axis ",
                      "or the data labels instead"),
  protection = "a chart shape cannot be locked or hidden"
)

# Not every chart part takes every piece.  A series is a shape and has no
# text, so it takes a line, a fill and a pattern but no font; a title is text
# and takes only a font.  A group the target cannot use has nowhere to go, so
# it is an error rather than a silent drop.
.CHART_PART_CANNOT <- c(
  font    = paste0("a font, but %s is a shape and has no text -- set the font ",
                   "on the chart title"),
  line    = "a border, but %s takes no line",
  fill    = "a fill, but %s takes no fill",
  pattern = "a pattern fill, but %s takes no fill"
)

.CHART_PART_GROUP <- c(font = "xl_font()", line = "xl_border()",
                       fill = "xl_fill()", pattern = "xl_fill()")

# Translate an xl_format into the pieces a chart understands.  Returns a list
# with any of `line`, `fill`, `pattern` and `font`, each ready for C.  `accept`
# names the pieces this particular chart part can use, and `part` names it for
# the error message.
.chart_format_payload <- function(fmt, arg,
                                  accept = c("line", "fill", "pattern", "font"),
                                  part = "a chart series") {
  if (is.null(fmt)) return(NULL)
  if (!is_xl_format(fmt))
    stop(sprintf("`%s` must be an xl_format object", arg), call. = FALSE)
  f <- unclass(fmt)

  for (g in names(.CHART_FORMAT_REFUSED))
    if (!is.null(f[[g]]))
      stop(sprintf("`%s` sets %s, which a chart cannot use: %s.", arg,
                   switch(g, num_format = "a number format",
                          protection = "protection"),
                   .CHART_FORMAT_REFUSED[[g]]), call. = FALSE)

  rotation <- NULL
  if (!is.null(f$align)) {
    placement <- setdiff(names(f$align), "rotation")
    if (length(placement))
      stop(sprintf(paste0("`%s` sets %s, which a chart cannot use: a chart ",
                          "shape places its own text. Of the alignment group ",
                          "it takes only `rotation`."),
                   arg, paste(sprintf("`%s`", placement), collapse = ", ")),
           call. = FALSE)
    # 0 has to travel as 360: lxw_chart_font treats a rotation of 0 as unset,
    # which is why libxlsxwriter spells "explicitly upright" 360
    rotation <- if (identical(as.integer(f$align$rotation), 0L)) 360L
                else as.integer(f$align$rotation)
  }

  out <- list()

  if (!is.null(f$font) || !is.null(rotation)) {
    fo <- f$font
    ent <- list()
    if (!is.null(fo$name))  ent$name <- fo$name
    if (!is.null(fo$size))  ent$size <- as.numeric(fo$size)
    if (!is.null(fo$color)) ent$color <- .chart_color(fo$color)
    if (isTRUE(fo$bold))    ent$bold <- 1L
    if (isTRUE(fo$italic))  ent$italic <- 1L
    # lxw_chart_font.underline is a flag, not Excel's five-way choice, so any
    # underline at all becomes underlined; "none" means not underlined
    if (!is.null(fo$underline) && !identical(fo$underline, "none"))
      ent$underline <- 1L
    if (!is.null(rotation)) ent$rotation <- rotation
    if (length(ent)) out$font <- ent
  }

  if (!is.null(f$fill)) {
    fi <- f$fill
    # a real pattern needs two colours and becomes lxw_chart_pattern; a plain
    # or solid fill becomes lxw_chart_fill
    if (!is.null(fi$pattern) && !identical(fi$pattern, "solid") &&
        !identical(fi$pattern, "none")) {
      if (is.null(fi$foreground) || is.null(fi$background))
        stop(sprintf(paste0("`%s` sets a pattern, which a chart draws from a ",
                            "foreground and a background colour -- give both."),
                     arg), call. = FALSE)
      out$pattern <- list(fg_color = .chart_color(fi$foreground),
                          bg_color = .chart_color(fi$background),
                          type = .chart_pattern(fi$pattern, arg))
    } else {
      ent <- list()
      if (!is.null(fi$background)) ent$color <- .chart_color(fi$background)
      if (identical(fi$pattern, "none")) ent$none <- 1L
      if (!is.null(fi$transparency))
        ent$transparency <- as.integer(fi$transparency)
      if (length(ent)) out$fill <- ent
    }
  }

  if (!is.null(f$border)) {
    b <- f$border
    sides <- c(b$left, b$right, b$top, b$bottom)
    cols  <- c(b$left_color, b$right_color, b$top_color, b$bottom_color)
    if (length(unique(sides)) > 1L || length(unique(cols)) > 1L)
      stop(sprintf(paste0("`%s` gives different borders on different sides; a ",
                          "chart has one line. Use xl_border(all = ) with a ",
                          "single style and colour."), arg), call. = FALSE)
    style <- if (length(sides)) sides[[1L]] else NULL
    ent <- list()
    if (identical(style, "none")) ent$none <- 1L
    else {
      d <- .chart_dash(style, arg)
      if (!is.null(d)) ent$dash_type <- d
    }
    if (length(cols)) ent$color <- .chart_color(cols[[1L]])
    if (!is.null(b$transparency))
      ent$transparency <- as.integer(b$transparency)
    if (length(ent)) out$line <- ent
  }

  extra <- setdiff(names(out), accept)
  if (length(extra))
    stop(sprintf("`%s` sets %s. Use %s instead.", arg,
                 sprintf(.CHART_PART_CANNOT[[extra[[1L]]]], part),
                 paste(unique(.CHART_PART_GROUP[accept]), collapse = " or ")),
         call. = FALSE)

  if (!length(out))
    stop(sprintf("`%s` is an empty format; give %s for the chart to use.", arg,
                 paste(unique(.CHART_PART_GROUP[accept]), collapse = " or ")),
         call. = FALSE)
  out
}

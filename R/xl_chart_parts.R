# =============================================================================
# The parts of a chart series: markers, data labels, trendlines, error bars
# and individual points
# =============================================================================
#
# Five constructors over the chart_series_*() functions.  Each takes a
# `format` translated by .chart_format_payload() into whichever of line / fill
# / pattern / font that part can use: a trendline is a line and takes nothing
# else, a data label is a shape with text and takes all four.
#
# Four restrictions are documented in chart.h and enforced nowhere:
#
#   * the positions a data label may take depend on the chart type, and the
#     header carries the whole table (Line/Scatter, Bar/Column, Pie/Doughnut,
#     Area/Radar);
#   * a moving-average trendline has no forecast, equation or R-squared;
#   * an intercept applies only to exponential, linear and polynomial
#     trendlines;
#   * an automatic marker cannot also be given a size, a line or a fill.
#
# Excel drops all of them without a word, so they are refused here.
# -----------------------------------------------------------------------------

.LXW_MARKER_TYPE <- c(
  automatic = 0L, none = 1L, square = 2L, diamond = 3L, triangle = 4L,
  x = 5L, star = 6L, short_dash = 7L, long_dash = 8L, circle = 9L, plus = 10L
)

.LXW_LABEL_POSITION <- c(
  default = 0L, center = 1L, right = 2L, left = 3L, above = 4L, below = 5L,
  inside_base = 6L, inside_end = 7L, outside_end = 8L, best_fit = 9L
)

# Which families accept each position, from the table in chart.h.  "default"
# means "leave it to Excel" and is legal everywhere.
.LABEL_POSITION_FAMILIES <- list(
  default     = c("line", "scatter", "bar", "pie", "area", "radar", "other"),
  center      = c("line", "scatter", "bar", "pie", "area", "radar", "other"),
  right       = c("line", "scatter"),
  left        = c("line", "scatter"),
  above       = c("line", "scatter"),
  below       = c("line", "scatter"),
  inside_base = c("bar", "other"),
  inside_end  = c("bar", "pie", "other"),
  outside_end = c("bar", "pie", "other"),
  best_fit    = "pie"
)

.LXW_LABEL_SEPARATOR <- c(comma = 0L, semicolon = 1L, period = 2L,
                          newline = 3L, space = 4L)

.LXW_TRENDLINE_TYPE <- c(linear = 0L, log = 1L, poly = 2L, power = 3L,
                         exp = 4L, average = 5L)

.LXW_ERROR_BAR_TYPE <- c(std_error = 0L, fixed = 1L, percentage = 2L,
                         std_dev = 3L)

.LXW_ERROR_BAR_DIRECTION <- c(both = 0L, plus = 1L, minus = 2L)

# --- Markers ------------------------------------------------------------------

#' A marker on a chart series
#'
#' @description
#' `xl_chart_marker()` draws a symbol at each point of a series.  Line, scatter
#' and radar charts are where they show; on other types Excel ignores them.
#'
#' `type = "automatic"` asks Excel for the default marker of that series, and
#' is the one type that cannot be given a size or a format --- libxlsxwriter
#' documents that, and Excel drops them, so both are refused here.
#'
#' @param type The symbol: `"automatic"`, `"none"`, `"square"`, `"diamond"`,
#'   `"triangle"`, `"x"`, `"star"`, `"short_dash"`, `"long_dash"`, `"circle"`
#'   or `"plus"`.
#' @param size The symbol's size in points, 2 to 72.
#' @param format An [xl_format()] styling the symbol: [xl_border()] for its
#'   outline, [xl_fill()] for its fill or pattern.
#' @return An `xl_chart_marker` object.
#' @family images and charts
#' @seealso [xl_chart_series]
#' @export
#' @examples
#' xl_chart_marker(type = "circle", size = 8)
#' xl_chart_marker(type = "none")
xl_chart_marker <- function(type = NULL, size = NA, format = NULL) {
  ty <- .val_enum(type, names(.LXW_MARKER_TYPE), "type")
  sz <- .val_int(size, "size", min = 2, max = 72)
  fmt <- .chart_format_payload(format, "format",
                               accept = c("line", "fill", "pattern"),
                               part = "a chart marker")
  if (identical(ty, "automatic") && (!is.null(sz) || !is.null(fmt)))
    stop(paste0("`type = \"automatic\"` asks Excel for the series' own default ",
                "marker, which cannot then be given a size or a format. Name ",
                "the symbol you want instead."), call. = FALSE)
  structure(.drop_null(list(type = ty, size = sz, format = fmt)),
            class = "xl_chart_marker")
}

# --- Data labels --------------------------------------------------------------

#' Data labels on a chart series
#'
#' @description
#' `xl_chart_labels()` prints the numbers next to the points that carry them.
#' With no arguments it shows the value of each point, which is Excel's own
#' default.
#'
#' `xl_chart_label()` describes one label, for `custom`: a label can be given
#' its own text or hidden, one point at a time.
#'
#' Excel allows different label positions for different chart types, and drops
#' one that does not apply.  The table is in libxlsxwriter's header and is
#' enforced here: `"center"` is allowed everywhere, `"right"`, `"left"`,
#' `"above"` and `"below"` on line and scatter charts, `"inside_base"` on bar
#' and column, `"inside_end"` and `"outside_end"` on bar, column, pie and
#' doughnut, and `"best_fit"` on pie and doughnut.
#'
#' @param show_value,show_name,show_category,show_percentage What each label
#'   holds: the point's value, the series name, the category, and the value as
#'   a percentage of the series.  Naming any of them means the label holds
#'   exactly those, so `show_percentage = TRUE` alone gives a percentage and
#'   nothing else; naming none of them leaves Excel's default, the value.
#'   Percentages are meaningful on pie and doughnut charts.
#' @param show_legend_key Print the series' legend swatch in each label.
#' @param num_format A number format for the labels, as an Excel format string
#'   or an [xl_num_format()].
#' @param position Where the label sits relative to its point; see the
#'   description for which chart types allow which.
#' @param separator What joins the parts of a label when it holds more than
#'   one: `"comma"`, `"semicolon"`, `"period"`, `"newline"` or `"space"`.
#' @param format An [xl_format()] styling the labels.  A label is a shape with
#'   text in it, so all of [xl_font()], [xl_border()] and [xl_fill()] apply.
#' @param leader_lines Draw a line from a label back to its point.  Excel only
#'   shows one once the label has been dragged away from the point.
#' @param custom A list of [xl_chart_label()]s, one per point in order, giving
#'   individual labels their own text, styling, or `hide = TRUE`.  `NULL` in
#'   the list leaves that point's label alone.
#' @return An `xl_chart_labels` object.
#' @family images and charts
#' @seealso [xl_chart_series]
#' @export
#' @examples
#' xl_chart_labels()
#' xl_chart_labels(show_category = TRUE, show_percentage = TRUE,
#'                 separator = "newline", position = "outside_end")
xl_chart_labels <- function(show_value = NA, show_name = NA,
                            show_category = NA, show_percentage = NA,
                            show_legend_key = NA, num_format = NULL,
                            position = NULL, separator = NULL,
                            format = NULL, leader_lines = NA, custom = NULL) {
  cu <- custom
  if (!is.null(cu)) {
    if (inherits(cu, "xl_chart_label")) cu <- list(cu)
    if (!is.list(cu))
      stop("`custom` must be a list of xl_chart_label() objects", call. = FALSE)
    for (i in seq_along(cu))
      if (!is.null(cu[[i]]) && !inherits(cu[[i]], "xl_chart_label"))
        stop(sprintf("`custom[[%d]]` must be an xl_chart_label() or NULL", i),
             call. = FALSE)
  }
  structure(.drop_null(list(
    show_value = .val_flag(show_value, "show_value"),
    show_name = .val_flag(show_name, "show_name"),
    show_category = .val_flag(show_category, "show_category"),
    show_percentage = .val_flag(show_percentage, "show_percentage"),
    show_legend_key = .val_flag(show_legend_key, "show_legend_key"),
    num_format = .chart_num_format(num_format, "num_format"),
    position = .val_enum(position, names(.LXW_LABEL_POSITION), "position"),
    separator = .val_enum(separator, names(.LXW_LABEL_SEPARATOR), "separator"),
    format = .chart_format_payload(format, "format", part = "a data label"),
    leader_lines = .val_flag(leader_lines, "leader_lines"),
    custom = cu
  )), class = "xl_chart_labels")
}

#' @param value The label's text.  A string beginning with `"="` is a formula,
#'   so `"=Sheet1!$A$1"` takes the text from a cell.
#' @param hide Remove this point's label, leaving the others.
#' @rdname xl_chart_labels
#' @export
xl_chart_label <- function(value = NULL, hide = NA, format = NULL) {
  hd <- .val_flag(hide, "hide")
  val <- .val_str(value, "value")
  if (isTRUE(hd) && (!is.null(val) || !is.null(format)))
    stop(paste0("`hide = TRUE` removes the label, so it cannot also carry a ",
                "value or a format"), call. = FALSE)
  structure(.drop_null(list(
    value = val, hide = if (isTRUE(hd)) 1L else NULL,
    format = .chart_format_payload(format, "format", part = "a data label")
  )), class = "xl_chart_label")
}

# --- Trendlines ---------------------------------------------------------------

#' A trendline on a chart series
#'
#' @description
#' `xl_chart_trendline()` fits a line through a series.
#'
#' Two of Excel's own restrictions are enforced, because it discards these
#' rather than complain: a **moving average** has no forecast, no equation and
#' no R-squared, and an **intercept** applies only to exponential, linear and
#' polynomial fits.
#'
#' @param type `"linear"`, `"log"`, `"poly"`, `"power"`, `"exp"` or
#'   `"average"` for a moving average.
#' @param order The order of a polynomial fit, 2 or more.  `"poly"` only.
#' @param period The number of points a moving average covers, 2 or more.
#'   `"average"` only.
#' @param forward,backward How far to project the line beyond the data, in
#'   categories.
#' @param intercept Force the line through this value on the y axis.
#'   Exponential, linear and polynomial fits only.
#' @param equation Print the fitted equation on the chart.
#' @param r_squared Print the R-squared value on the chart.
#' @param name The trendline's name in the legend.  Excel generates one
#'   otherwise.
#' @param format An [xl_format()] styling the line --- [xl_border()] only,
#'   since a trendline is a line.
#' @return An `xl_chart_trendline` object.
#' @family images and charts
#' @seealso [xl_chart_series]
#' @export
#' @examples
#' xl_chart_trendline("linear", equation = TRUE, r_squared = TRUE)
#' xl_chart_trendline("poly", order = 3)
#' xl_chart_trendline("average", period = 2)
xl_chart_trendline <- function(type, order = NA, period = NA, forward = NA,
                               backward = NA, intercept = NA, equation = NA,
                               r_squared = NA, name = NULL, format = NULL) {
  if (missing(type))
    stop("`type` must name the fit, e.g. \"linear\"", call. = FALSE)
  ty <- .val_enum(type, names(.LXW_TRENDLINE_TYPE), "type")
  if (is.null(ty))
    stop("`type` must name the fit, e.g. \"linear\"", call. = FALSE)
  or <- .val_int(order, "order", min = 2)
  pe <- .val_int(period, "period", min = 2)
  if (!is.null(or) && !identical(ty, "poly"))
    stop(sprintf("`order` is the degree of a polynomial fit, not of a \"%s\" one",
                 ty), call. = FALSE)
  if (!is.null(pe) && !identical(ty, "average"))
    stop(sprintf(paste0("`period` is how many points a moving average covers; ",
                        "a \"%s\" fit has none"), ty), call. = FALSE)
  if (identical(ty, "poly") && is.null(or))
    stop("a \"poly\" trendline needs an `order`", call. = FALSE)
  if (identical(ty, "average") && is.null(pe))
    stop("an \"average\" trendline needs a `period`", call. = FALSE)

  fw <- .val_num(forward, "forward", min = 0)
  bw <- .val_num(backward, "backward", min = 0)
  ic <- .val_num(intercept, "intercept")
  eq <- .val_flag(equation, "equation")
  rs <- .val_flag(r_squared, "r_squared")
  if (identical(ty, "average")) {
    given <- c("forward", "backward", "equation", "r_squared")[
      c(!is.null(fw), !is.null(bw), isTRUE(eq), isTRUE(rs))]
    if (length(given))
      stop(sprintf(paste0("`%s` does not apply to a moving average; Excel ",
                          "offers a forecast, an equation and an R-squared ",
                          "for the fitted types only."), given[[1L]]),
           call. = FALSE)
  }
  if (!is.null(ic) && !ty %in% c("exp", "linear", "poly"))
    stop(sprintf(paste0("`intercept` applies to an exponential, linear or ",
                        "polynomial fit, not to a \"%s\" one."), ty),
         call. = FALSE)

  structure(.drop_null(list(
    type = ty, order = or, period = pe, forward = fw, backward = bw,
    intercept = ic, equation = if (isTRUE(eq)) 1L else NULL,
    r_squared = if (isTRUE(rs)) 1L else NULL,
    name = .val_str(name, "name"),
    format = .chart_format_payload(format, "format", accept = "line",
                                   part = "a trendline")
  )), class = "xl_chart_trendline")
}

# --- Error bars ---------------------------------------------------------------

#' Error bars on a chart series
#'
#' @description
#' `xl_chart_error_bars()` draws an error bar at each point, and is given to
#' [xl_chart_series()] as `x_error_bars` or `y_error_bars`.
#'
#' @param type How the size of each bar is worked out: `"std_error"` for the
#'   standard error, `"fixed"` for a constant, `"percentage"` of the point's
#'   own value, or `"std_dev"` for that many standard deviations.
#' @param value The constant, the percentage, or the number of standard
#'   deviations.  `"std_error"` needs none.
#' @param direction `"both"`, `"plus"` or `"minus"`.
#' @param endcap Draw the cap at the end of each bar.  On by default.
#' @param format An [xl_format()] styling the bars --- [xl_border()] only,
#'   since an error bar is a line.
#' @return An `xl_chart_error_bars` object.
#' @family images and charts
#' @seealso [xl_chart_series]
#' @export
#' @examples
#' xl_chart_error_bars("percentage", 5)
#' xl_chart_error_bars("std_dev", 1, direction = "plus", endcap = FALSE)
xl_chart_error_bars <- function(type, value = NA, direction = NULL,
                                endcap = NA, format = NULL) {
  if (missing(type))
    stop("`type` must say how the bars are sized, e.g. \"percentage\"",
         call. = FALSE)
  ty <- .val_enum(type, names(.LXW_ERROR_BAR_TYPE), "type")
  if (is.null(ty))
    stop("`type` must say how the bars are sized, e.g. \"percentage\"",
         call. = FALSE)
  val <- .val_num(value, "value", min = 0)
  if (is.null(val) && !identical(ty, "std_error"))
    stop(sprintf(paste0("a \"%s\" error bar needs a `value` -- the %s it is ",
                        "measured in."), ty,
                 switch(ty, fixed = "constant amount",
                        percentage = "percentage",
                        std_dev = "number of standard deviations")),
         call. = FALSE)
  structure(.drop_null(list(
    type = ty, value = val,
    direction = .val_enum(direction, names(.LXW_ERROR_BAR_DIRECTION),
                          "direction"),
    endcap = .val_flag(endcap, "endcap"),
    format = .chart_format_payload(format, "format", accept = "line",
                                   part = "an error bar")
  )), class = "xl_chart_error_bars")
}

# --- Shared helpers -----------------------------------------------------------

# A number format may be written as the Excel string or as xl_num_format().
.chart_num_format <- function(x, arg) {
  if (is.null(x)) return(NULL)
  if (is_xl_format(x)) {
    f <- unclass(x)$num_format$format
    if (is.null(f))
      stop(sprintf("`%s` is a format with no number format in it", arg),
           call. = FALSE)
    return(f)
  }
  if (!is.character(x) || length(x) != 1L || is.na(x))
    stop(sprintf("`%s` must be an Excel format string, e.g. \"#,##0.0\"", arg),
         call. = FALSE)
  x
}

# --- Checks that need the chart's type ----------------------------------------

.check_label_position <- function(pos, type, arg) {
  if (is.null(pos)) return(invisible(NULL))
  fams <- .LABEL_POSITION_FAMILIES[[pos]]
  if (.CHART_FAMILY(type) %in% fams) return(invisible(NULL))
  ok <- Filter(function(t) .CHART_FAMILY(t) %in% fams, names(.LXW_CHART_TYPE))
  allowed <- Filter(function(p) .CHART_FAMILY(type) %in%
                      .LABEL_POSITION_FAMILIES[[p]],
                    names(.LXW_LABEL_POSITION))
  stop(sprintf(paste0("`%s$position = \"%s\"` does not apply to a \"%s\" ",
                      "chart; Excel drops it.\n  On this chart the positions ",
                      "are: %s.\n  \"%s\" applies to: %s."),
               arg, pos, type, paste(allowed, collapse = ", "), pos,
               paste(ok, collapse = ", ")), call. = FALSE)
}

# Which constructor each xl_chart_series() part argument expects.
.PART_CLASS <- c(marker = "xl_chart_marker", labels = "xl_chart_labels",
                 trendline = "xl_chart_trendline",
                 x_error_bars = "xl_chart_error_bars",
                 y_error_bars = "xl_chart_error_bars")

# What needs the chart's type, and so cannot be checked until the series is
# put into a chart.  The parts validate themselves otherwise.
.check_series_parts <- function(se, type, arg) {
  p <- unclass(se)
  if (!is.null(p[["labels"]]))
    .check_label_position(unclass(p[["labels"]])[["position"]], type,
                          sprintf("%s$labels", arg))
  invisible(NULL)
}

# --- Payloads -----------------------------------------------------------------

.marker_payload <- function(m) {
  if (is.null(m)) return(NULL)
  p <- unclass(m)
  ent <- list()
  if (!is.null(p[["type"]]))
    ent$type <- unname(.LXW_MARKER_TYPE[[p[["type"]]]])
  ent$size <- p[["size"]]
  ent$line <- p[["format"]][["line"]]
  ent$fill <- p[["format"]][["fill"]]
  ent$pattern <- p[["format"]][["pattern"]]
  .drop_null(ent)
}

.label_payload <- function(l) {
  p <- unclass(l)
  .drop_null(list(value = p[["value"]], hide = p[["hide"]],
                  font = p[["format"]][["font"]],
                  line = p[["format"]][["line"]],
                  fill = p[["format"]][["fill"]],
                  pattern = p[["format"]][["pattern"]]))
}

.labels_payload <- function(l) {
  if (is.null(l)) return(NULL)
  p <- unclass(l)
  ent <- list(on = 1L)
  # Excel puts the value in a label unless told otherwise, so asking for the
  # percentage alone has to switch the value off -- a label would otherwise
  # read "10, 11.8%".  Naming any part means the label holds exactly those.
  shown <- c("show_name", "show_category", "show_value")
  asked <- c(shown, "show_percentage")
  if (any(vapply(asked, function(k) !is.null(p[[k]]), logical(1))))
    ent$options <- as.integer(vapply(shown, function(k) isTRUE(p[[k]]),
                                     logical(1)))
  if (isTRUE(p[["show_percentage"]])) ent$percentage <- 1L
  if (isTRUE(p[["show_legend_key"]])) ent$legend <- 1L
  if (isTRUE(p[["leader_lines"]]))    ent$leader_lines <- 1L
  if (!is.null(p[["position"]]))
    ent$position <- unname(.LXW_LABEL_POSITION[[p[["position"]]]])
  if (!is.null(p[["separator"]]))
    ent$separator <- unname(.LXW_LABEL_SEPARATOR[[p[["separator"]]]])
  ent$num_format <- p[["num_format"]]
  ent$font    <- p[["format"]][["font"]]
  ent$line    <- p[["format"]][["line"]]
  ent$fill    <- p[["format"]][["fill"]]
  ent$pattern <- p[["format"]][["pattern"]]
  if (!is.null(p[["custom"]]))
    # a NULL entry leaves that point alone, and reaches C as an empty label
    ent$custom <- lapply(p[["custom"]], function(x)
      if (is.null(x)) list() else .label_payload(x))
  .drop_null(ent)
}

.trendline_payload <- function(t) {
  if (is.null(t)) return(NULL)
  p <- unclass(t)
  # chart_series_set_trendline() takes one `value`: the polynomial order or
  # the moving-average period, depending on the type
  val <- if (!is.null(p[["order"]])) p[["order"]] else p[["period"]]
  .drop_null(list(
    type = unname(.LXW_TRENDLINE_TYPE[[p[["type"]]]]),
    value = if (is.null(val)) NULL else as.integer(val),
    forward = p[["forward"]], backward = p[["backward"]],
    intercept = p[["intercept"]], equation = p[["equation"]],
    r_squared = p[["r_squared"]], name = p[["name"]],
    line = p[["format"]][["line"]]))
}

.error_bars_payload <- function(e) {
  if (is.null(e)) return(NULL)
  p <- unclass(e)
  .drop_null(list(
    type = unname(.LXW_ERROR_BAR_TYPE[[p[["type"]]]]),
    value = if (is.null(p[["value"]])) 0 else p[["value"]],
    direction = if (is.null(p[["direction"]])) NULL else
      unname(.LXW_ERROR_BAR_DIRECTION[[p[["direction"]]]]),
    endcap = if (is.null(p[["endcap"]])) NULL else
      as.integer(p[["endcap"]]),
    line = p[["format"]][["line"]]))
}

.points_payload <- function(pts, arg) {
  if (is.null(pts)) return(NULL)
  if (is_xl_format(pts)) pts <- list(pts)
  if (!is.list(pts))
    stop(sprintf(paste0("`%s` must be a list of xl_format objects, one per ",
                        "point, with NULL for a point to leave alone"), arg),
         call. = FALSE)
  out <- lapply(seq_along(pts), function(i) {
    x <- pts[[i]]
    if (is.null(x)) return(list())
    f <- .chart_format_payload(x, sprintf("%s[[%d]]", arg, i),
                               accept = c("line", "fill", "pattern"),
                               part = "a chart point")
    .drop_null(list(line = f[["line"]], fill = f[["fill"]],
                    pattern = f[["pattern"]]))
  })
  if (!length(out)) return(NULL)
  out
}

#' @export
print.xl_chart_marker <- function(x, ...) .print_part(x, "xl_chart_marker")

#' @export
print.xl_chart_labels <- function(x, ...) .print_part(x, "xl_chart_labels")

#' @export
print.xl_chart_label <- function(x, ...) .print_part(x, "xl_chart_label")

#' @export
print.xl_chart_trendline <- function(x, ...) .print_part(x, "xl_chart_trendline")

#' @export
print.xl_chart_error_bars <- function(x, ...)
  .print_part(x, "xl_chart_error_bars")

.print_part <- function(x, cls) {
  p <- unclass(x)
  cat("<", cls, ">\n", sep = "")
  if (length(p)) cat("  set:", paste(names(p), collapse = ", "), "\n")
  invisible(x)
}

# =============================================================================
# Chart axes
# =============================================================================
#
# libxlsxwriter offers 32 chart_axis_*() functions.  Most take one value, and
# several of them are only meaningful on one *kind* of axis -- the header
# documents that per function ("Axis types: value axes only") and then does not
# enforce it: setting a minimum on a category axis is accepted, written, and
# ignored by Excel.
#
# There are two kinds of axis in reach.  libxlsxwriter also mentions date axes,
# but nothing in its API turns an axis into one, so every axis writexl can
# produce is either a category axis or a value axis:
#
#   * a scatter chart plots numbers against numbers, so both of its axes are
#     value axes;
#   * every other type plots values against labels, so x is the category axis
#     and y the value axis.  This holds for bar charts too, which Excel draws
#     with the categories running up the side -- the drawing is rotated, the
#     axes are not renamed.
#
# .check_axis_option() refuses an option the axis cannot use, naming the axis
# that can.
# -----------------------------------------------------------------------------

.LXW_AXIS_POSITION <- c(on_tick = 1L, between = 2L)

.LXW_AXIS_LABEL_POSITION <- c(next_to = 0L, high = 1L, low = 2L, none = 3L)

.LXW_AXIS_LABEL_ALIGN <- c(center = 0L, left = 1L, right = 2L)

.LXW_AXIS_TICK_MARK <- c(default = 0L, none = 1L, inside = 2L, outside = 3L,
                         crossing = 4L)

# The names are Excel's, not the enum's: "ten_thousands" rather than
# TEN_THOUSANDS, and the order is the enum's own.
.LXW_AXIS_UNITS <- c(
  none = 0L, hundreds = 1L, thousands = 2L, ten_thousands = 3L,
  hundred_thousands = 4L, millions = 5L, ten_millions = 6L,
  hundred_millions = 7L, billions = 8L, trillions = 9L
)

# Options only one kind of axis can use, and which kind that is.  Taken from
# the "Axis types" line of each function's documentation in chart.h; everything
# not listed applies to both.
.AXIS_OPTION_KIND <- c(
  position      = "category",   # chart_axis_set_position()
  label_align   = "category",   # chart_axis_set_label_align()
  interval_unit = "category",   # chart_axis_set_interval_unit()
  interval_tick = "category",   # chart_axis_set_interval_tick()
  min           = "value",      # chart_axis_set_min()
  max           = "value",      # chart_axis_set_max()
  log_base      = "value",      # chart_axis_set_log_base()
  major_unit    = "value",      # chart_axis_set_major_unit()
  minor_unit    = "value",      # chart_axis_set_minor_unit()
  display_units = "value",      # chart_axis_set_display_units()
  display_units_visible = "value"
)

# Which kind an axis is, given the chart type it belongs to.
.axis_kind <- function(type, which) {
  if (identical(.CHART_FAMILY(type), "scatter")) "value"
  else if (identical(which, "x")) "category" else "value"
}

.check_axis_option <- function(opt, type, which) {
  kind <- .AXIS_OPTION_KIND[[opt]]
  if (is.null(kind) || is.na(kind)) return(invisible(NULL))
  have <- .axis_kind(type, which)
  if (identical(kind, have)) return(invisible(NULL))
  other <- if (identical(which, "x")) "y_axis" else "x_axis"
  stop(sprintf(paste0("`%s_axis$%s` applies to a %s axis, and the %s axis of a ",
                      "\"%s\" chart is a %s axis; Excel ignores it there.\n  ",
                      "%s"),
               which, opt, kind, which, type, have,
               if (identical(.CHART_FAMILY(type), "scatter"))
                 "Both axes of a scatter chart are value axes."
               else sprintf("Use `%s`, or a chart type whose %s axis is a %s axis.",
                            other, which, kind)),
       call. = FALSE)
}

#' An axis of a chart
#'
#' @description
#' `xl_chart_axis()` describes one axis, and is given to [xl_chart()] as
#' `x_axis` or `y_axis`.
#'
#' Several options apply to one kind of axis only, and Excel discards the rest
#' without a word, so writexl refuses them instead.  A scatter chart plots
#' numbers against numbers, so both of its axes are **value** axes; every other
#' type has a **category** x axis and a value y axis.  (A bar chart is drawn
#' with its categories up the side, but the axes keep their names.)  Pie and
#' doughnut charts have no axes at all.
#'
#' * value axes only --- `min`, `max`, `log_base`, `major_unit`, `minor_unit`,
#'   `display_units`, `display_units_visible`;
#' * category axes only --- `position`, `label_align`, `interval_unit`,
#'   `interval_tick`.
#'
#' @param title The axis title: a string, or a range spec holding one --- see
#'   [xl_chart_series()] for the spellings, including
#'   `list(header = "revenue")`.
#' @param title_format An [xl_format()] styling the axis title.  A title is
#'   text, so only the [xl_font()] group applies.
#' @param title_layout Where to put the axis title by hand, as `c(x, y)`
#'   fractions of the chart, each above 0 and at most 1.  Excel places it for
#'   you otherwise.
#' @param label_format An [xl_format()] styling the tick labels --- the
#'   [xl_font()] group only.
#' @param num_format A number format for the tick labels, as an Excel format
#'   string (`"#,##0"`) or an [xl_num_format()].
#' @param line_format An [xl_format()] styling the axis line itself:
#'   [xl_border()] for the line, [xl_fill()] for the fill behind it.  An axis
#'   has four parts that can be styled, so none of them is just `format`.
#' @param visible `FALSE` hides the axis.
#' @param reverse Draw the axis in the opposite direction.
#' @param min,max The axis bounds.  Value axes only.
#' @param log_base Use a logarithmic scale with this base, 2 or more.  Value
#'   axes only.
#' @param major_unit,minor_unit The spacing between major and minor tick marks.
#'   Value axes only.
#' @param display_units Scale the labels by `"thousands"`, `"millions"`,
#'   `"billions"` and so on; see Details for the full set.  Value axes only.
#' @param display_units_visible Whether the caption naming the units --- the
#'   small rotated "Millions" beside the axis --- is drawn.  Setting
#'   `display_units` turns it **on**, as Excel does, so this is really for
#'   `FALSE`: rescaled labels with no caption.  Value axes only.
#' @param interval_unit Label one category in every `n`.  Category axes only.
#' @param interval_tick Put a tick mark on one category in every `n`.
#'   Category axes only.
#' @param position Whether the data sits `"on_tick"` or `"between"` the tick
#'   marks.  Category axes only.
#' @param label_position Where the tick labels go: `"next_to"`, `"high"`,
#'   `"low"`, or `"none"` for no labels.
#' @param label_align Tick-label alignment: `"center"`, `"left"` or `"right"`.
#'   Category axes only.
#' @param major_tick,minor_tick The tick marks: `"default"`, `"none"`,
#'   `"inside"`, `"outside"` or `"crossing"`.
#' @param crossing Where the other axis crosses this one: a number, or `"min"`
#'   or `"max"` for either end.
#' @param major_gridlines,minor_gridlines Show the gridlines.  A chart's major
#'   y gridlines are on by default and everything else is off.
#' @param major_gridlines_format,minor_gridlines_format An [xl_format()]
#'   styling the gridlines --- [xl_border()] only, since a gridline is a line.
#' @return An `xl_chart_axis` object.
#' @details
#' `display_units` is one of `"none"`, `"hundreds"`, `"thousands"`,
#' `"ten_thousands"`, `"hundred_thousands"`, `"millions"`, `"ten_millions"`,
#' `"hundred_millions"`, `"billions"` or `"trillions"`.
#' @family images and charts
#' @seealso [xl_chart], [xl_chart_series]
#' @export
#' @examples
#' xl_chart_axis(title = "Quarter")
#' xl_chart_axis(title = "Revenue", min = 0, num_format = "$#,##0",
#'               major_gridlines = FALSE)
xl_chart_axis <- function(title = NULL, title_format = NULL, title_layout = NULL,
                          label_format = NULL, num_format = NULL,
                          line_format = NULL, visible = NA, reverse = NA,
                          min = NA, max = NA, log_base = NA,
                          major_unit = NA, minor_unit = NA,
                          display_units = NULL, display_units_visible = NA,
                          interval_unit = NA, interval_tick = NA,
                          position = NULL, label_position = NULL,
                          label_align = NULL,
                          major_tick = NULL, minor_tick = NULL,
                          crossing = NULL,
                          major_gridlines = NA, minor_gridlines = NA,
                          major_gridlines_format = NULL,
                          minor_gridlines_format = NULL) {
  nm <- if (is.null(title)) NULL
        else if (is.character(title) && length(title) == 1L && !is.na(title))
          list(text = title)
        else .chart_range(title, "title")
  if (!is.null(title_format) && is.null(nm))
    stop(paste0("`title_format` styles the axis title, so give a `title` too. ",
                "An axis with no title has nothing to style."), call. = FALSE)

  # chart_axis_set_num_format() takes the format string itself rather than an
  # xl_format; the constructor is accepted too so either spelling works.
  nf <- num_format
  if (is_xl_format(nf)) {
    grp <- unclass(nf)
    nf <- grp$num_format$format
    if (is.null(nf))
      stop("`num_format` is a format with no number format in it",
           call. = FALSE)
  }
  if (!is.null(nf) && (!is.character(nf) || length(nf) != 1L || is.na(nf)))
    stop("`num_format` must be an Excel format string, e.g. \"#,##0.0\"",
         call. = FALSE)

  cr <- crossing
  if (!is.null(cr)) {
    if (is.character(cr)) {
      if (!identical(cr, "min") && !identical(cr, "max"))
        stop("`crossing` must be a number, \"min\" or \"max\"", call. = FALSE)
    } else if (!is.numeric(cr) || length(cr) != 1L || is.na(cr))
      stop("`crossing` must be a number, \"min\" or \"max\"", call. = FALSE)
  }

  at <- title_layout
  if (!is.null(at)) {
    if (!is.numeric(at) || length(at) != 2L || anyNA(at) ||
        any(at <= 0) || any(at > 1))
      stop(paste0("`title_layout` must be c(x, y), each above 0 and at most 1 ",
                  "-- they are fractions of the chart's width and height"),
           call. = FALSE)
    at <- as.numeric(at)
  }

  lb <- .val_int(log_base, "log_base", min = 2, max = 1000)

  structure(.drop_null(list(
    title = nm,
    title_format = .chart_format_payload(title_format, "title_format",
                                        accept = "font",
                                        part = "an axis title"),
    title_layout = at,
    label_format = .chart_format_payload(label_format, "label_format",
                                         accept = "font",
                                         part = "an axis tick label"),
    num_format = nf,
    line_format = .chart_format_payload(line_format, "line_format",
                                        accept = c("line", "fill", "pattern"),
                                        part = "an axis"),
    visible = .val_flag(visible, "visible"),
    reverse = .val_flag(reverse, "reverse"),
    min = .val_num(min, "min"), max = .val_num(max, "max"),
    log_base = lb,
    major_unit = .val_num(major_unit, "major_unit", min = 0),
    minor_unit = .val_num(minor_unit, "minor_unit", min = 0),
    display_units = .val_enum(display_units, names(.LXW_AXIS_UNITS),
                              "display_units"),
    display_units_visible = .val_flag(display_units_visible,
                                      "display_units_visible"),
    interval_unit = .val_int(interval_unit, "interval_unit", min = 1),
    interval_tick = .val_int(interval_tick, "interval_tick", min = 1),
    position = .val_enum(position, names(.LXW_AXIS_POSITION), "position"),
    label_position = .val_enum(label_position,
                               names(.LXW_AXIS_LABEL_POSITION),
                               "label_position"),
    label_align = .val_enum(label_align, names(.LXW_AXIS_LABEL_ALIGN),
                            "label_align"),
    major_tick = .val_enum(major_tick, names(.LXW_AXIS_TICK_MARK),
                           "major_tick"),
    minor_tick = .val_enum(minor_tick, names(.LXW_AXIS_TICK_MARK),
                           "minor_tick"),
    crossing = cr,
    major_gridlines = .val_flag(major_gridlines, "major_gridlines"),
    minor_gridlines = .val_flag(minor_gridlines, "minor_gridlines"),
    major_gridlines_format = .chart_format_payload(
      major_gridlines_format, "major_gridlines_format", accept = "line",
      part = "a gridline"),
    minor_gridlines_format = .chart_format_payload(
      minor_gridlines_format, "minor_gridlines_format", accept = "line",
      part = "a gridline")
  )), class = "xl_chart_axis")
}

#' @export
print.xl_chart_axis <- function(x, ...) {
  p <- unclass(x)
  cat("<xl_chart_axis>",
      if (!is.null(p[["name"]][["text"]]))
        paste0(" \"", p[["name"]][["text"]], "\"") else "",
      "\n", sep = "")
  set <- setdiff(names(p), "name")
  if (length(set)) cat("  set:", paste(set, collapse = ", "), "\n")
  invisible(x)
}

# --- Resolution --------------------------------------------------------------

# Everything about an axis that the chart's type decides.  Checked here so it
# is refused where the caller wrote it rather than at write time.
.check_axis <- function(ax, which, type, arg) {
  if (is.null(ax)) return(invisible(NULL))
  if (!inherits(ax, "xl_chart_axis"))
    stop(sprintf("`%s` must be an xl_chart_axis() object", arg), call. = FALSE)
  # pie and doughnut charts have no axes to put this on
  .check_chart_feature(type, "axes", arg)
  for (opt in intersect(names(unclass(ax)), names(.AXIS_OPTION_KIND)))
    .check_axis_option(opt, type, which)
  invisible(NULL)
}

# One axis -> the payload C reads.  Anything type-dependent has already been
# checked by .check_axis().
.axis_payload <- function(ax, arg, sheets, own, header_offset) {
  if (is.null(ax)) return(NULL)
  p <- unclass(ax)

  ent <- list()
  nm <- p[["title"]]
  if (!is.null(nm)) {
    if (!is.null(nm[["text"]])) ent$name <- nm[["text"]]
    else {
      r <- .resolve_chart_range(nm, sprintf("%s title", arg), sheets, own,
                                header_offset)
      ent$name_sheet <- r$sheet
      ent$name_range <- r$range
    }
  }
  # the payload itself, not its $font: chart_font_of() looks the font up by
  # key, the same way it reads a series format
  ent$name_font  <- p[["title_format"]]
  ent$num_font   <- p[["label_format"]]
  ent$num_format <- p[["num_format"]]
  if (!is.null(p[["title_layout"]])) {
    ent$name_x <- p[["title_layout"]][[1L]]
    ent$name_y <- p[["title_layout"]][[2L]]
  }
  ent$line    <- p[["line_format"]][["line"]]
  ent$fill    <- p[["line_format"]][["fill"]]
  ent$pattern <- p[["line_format"]][["pattern"]]

  if (identical(p[["visible"]], FALSE)) ent$off <- 1L
  if (isTRUE(p[["reverse"]]))           ent$reverse <- 1L
  ent$min <- p[["min"]]
  ent$max <- p[["max"]]
  ent$log_base <- p[["log_base"]]
  ent$major_unit <- p[["major_unit"]]
  ent$minor_unit <- p[["minor_unit"]]
  ent$interval_unit <- p[["interval_unit"]]
  ent$interval_tick <- p[["interval_tick"]]
  if (!is.null(p[["display_units"]]))
    ent$display_units <- unname(.LXW_AXIS_UNITS[[p[["display_units"]]]])
  if (!is.null(p[["display_units_visible"]]))
    ent$display_units_visible <- as.integer(p[["display_units_visible"]])
  if (!is.null(p[["position"]]))
    ent$position <- unname(.LXW_AXIS_POSITION[[p[["position"]]]])
  if (!is.null(p[["label_position"]]))
    ent$label_position <-
      unname(.LXW_AXIS_LABEL_POSITION[[p[["label_position"]]]])
  if (!is.null(p[["label_align"]]))
    ent$label_align <- unname(.LXW_AXIS_LABEL_ALIGN[[p[["label_align"]]]])
  if (!is.null(p[["major_tick"]]))
    ent$major_tick <- unname(.LXW_AXIS_TICK_MARK[[p[["major_tick"]]]])
  if (!is.null(p[["minor_tick"]]))
    ent$minor_tick <- unname(.LXW_AXIS_TICK_MARK[[p[["minor_tick"]]]])

  cr <- p[["crossing"]]
  if (!is.null(cr)) {
    if (identical(cr, "min")) ent$crossing_min <- 1L
    else if (identical(cr, "max")) ent$crossing_max <- 1L
    else ent$crossing <- as.numeric(cr)
  }

  if (!is.null(p[["major_gridlines"]]))
    ent$major_gridlines <- as.integer(p[["major_gridlines"]])
  if (!is.null(p[["minor_gridlines"]]))
    ent$minor_gridlines <- as.integer(p[["minor_gridlines"]])
  # a gridline styled but never switched on would be invisible
  for (g in c("major", "minor")) {
    ln <- p[[paste0(g, "_gridlines_format")]][["line"]]
    if (is.null(ln)) next
    ent[[paste0(g, "_gridlines_line")]] <- ln
    if (is.null(ent[[paste0(g, "_gridlines")]]))
      ent[[paste0(g, "_gridlines")]] <- 1L
  }

  .drop_null(ent)
}

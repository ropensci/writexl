# =============================================================================
# Charts
# =============================================================================
#
# A chart is a drawing placed on a worksheet, so its placement arguments are
# xl_image()'s: lxw_chart_options is lxw_image_options minus url, tip and
# cell_format, with the same field names.  Keeping the R names identical is not
# cosmetic -- the API-consistency gate enforces it.
#
# A series names two ranges: the values, and optionally the categories to plot
# them against.  chart_series_set_values() and _set_categories() take a sheet
# name plus a 0-based quad rather than an A1 string, so the shared range
# resolver's output plugs straight in.  A series may point at a different sheet
# from the one the chart sits on, which is why ranges carry an optional `sheet`
# and are resolved once the whole workbook is known.
#
# Charts themselves produce a drawing part, so they do not desync
# libxlsxwriter's drawing-id counter -- but they are *victims* of it in the
# same way a floating image is, which .check_drawing_order() accounts for.
# -----------------------------------------------------------------------------

# The chart types, named as Excel names them rather than as the enum spells it.
.LXW_CHART_TYPE <- c(
  area = 1L, area_stacked = 2L, area_stacked_percent = 3L,
  bar = 4L, bar_stacked = 5L, bar_stacked_percent = 6L,
  column = 7L, column_stacked = 8L, column_stacked_percent = 9L,
  doughnut = 10L,
  line = 11L, line_stacked = 12L, line_stacked_percent = 13L,
  pie = 14L,
  scatter = 15L, scatter_straight = 16L, scatter_straight_markers = 17L,
  scatter_smooth = 18L, scatter_smooth_markers = 19L,
  radar = 20L, radar_markers = 21L, radar_filled = 22L
)

# In the enum's own order, so it can be read against the header.  NONE is 0
# rather than one past the end, and chart_legend_set_position() rejects
# anything above OVERLAY_TOP_RIGHT.
.LXW_CHART_LEGEND <- c(
  none = 0L, right = 1L, left = 2L, top = 3L, bottom = 4L, top_right = 5L,
  overlay_right = 6L, overlay_left = 7L, overlay_top_right = 8L
)

# --- Which features a chart type supports ------------------------------------
#
# libxlsxwriter documents these restrictions in prose and then ignores them:
# chart_set_hole_size() on a pie chart, or up-down bars on a column chart, are
# accepted and silently dropped by Excel.  The matrix is here so the refusal
# happens in R, where it can name the type.
.CHART_FAMILY <- function(type) {
  if (type %in% c("pie", "doughnut")) return("pie")
  if (startsWith(type, "scatter")) return("scatter")
  if (startsWith(type, "radar")) return("radar")
  if (startsWith(type, "bar") || startsWith(type, "column")) return("bar")
  if (startsWith(type, "line")) return("line")
  if (startsWith(type, "area")) return("area")
  "other"
}

# feature -> the families that support it
.CHART_FEATURE_FAMILIES <- list(
  axes            = c("scatter", "radar", "bar", "line", "area", "other"),
  hole_size       = "pie",     # doughnut only, narrowed below
  rotation        = "pie",
  up_down_bars    = "line",
  high_low_lines  = "line",
  drop_lines      = c("line", "area"),
  series_gap      = "bar",
  series_overlap  = "bar",
  smooth          = c("line", "scatter")
)

# Is `feature` legal for `type`?  Two features are narrower than their family.
.chart_supports <- function(type, feature) {
  if (feature == "hole_size") return(identical(type, "doughnut"))
  fam <- .CHART_FAMILY(type)
  fam %in% .CHART_FEATURE_FAMILIES[[feature]]
}

.check_chart_feature <- function(type, feature, arg) {
  if (.chart_supports(type, feature)) return(invisible(NULL))
  ok <- Filter(function(t) .chart_supports(t, feature), names(.LXW_CHART_TYPE))
  stop(sprintf(paste0("`%s` does not apply to a \"%s\" chart; Excel drops it ",
                      "silently.\n  It applies to: %s."),
               arg, type, paste(ok, collapse = ", ")), call. = FALSE)
}

# --- Ranges a series points at ------------------------------------------------
#
# A range may name a different sheet, so it cannot be resolved until every
# sheet is known.  Until then it is carried as given.
.chart_range <- function(x, arg) {
  if (is.null(x)) return(NULL)
  if (is.character(x) && length(x) == 1L && !is.na(x)) return(list(spec = x))
  if (is.list(x)) {
    nms <- names(x)
    if (is.null(nms) || any(!nzchar(nms)))
      stop(sprintf(paste0("`%s` list must be named, e.g. ",
                          "list(sheet = \"Data\", cols = \"revenue\")"), arg),
           call. = FALSE)
    unknown <- setdiff(nms, c("sheet", "rows", "cols", "header"))
    if (length(unknown))
      stop(sprintf("unknown `%s` element(s): %s", arg,
                   paste(unknown, collapse = ", ")), call. = FALSE)
    # `header` names one column's header cell.  Its own element rather than
    # `rows = 0`, because `rows` counts data rows everywhere else in writexl.
    if (!is.null(x[["header"]])) {
      if (!is.null(x[["rows"]]) || !is.null(x[["cols"]]))
        stop(sprintf(paste0("`%s` gives `header` as well as rows/cols; ",
                            "`header` is the header cell of one column, so it ",
                            "stands alone."), arg), call. = FALSE)
      if (length(x[["header"]]) != 1L || anyNA(x[["header"]]))
        stop(sprintf("`%s$header` must name a single column", arg),
             call. = FALSE)
    }
    sheet <- x[["sheet"]]
    if (!is.null(sheet) &&
        (!is.character(sheet) || length(sheet) != 1L || is.na(sheet)))
      stop(sprintf("`%s$sheet` must be a single sheet name", arg),
           call. = FALSE)
    return(list(spec = x[setdiff(nms, "sheet")], sheet = sheet))
  }
  stop(sprintf(paste0("`%s` must be a range string or a list(sheet = , rows = ",
                      ", cols = ) spec"), arg), call. = FALSE)
}

#' A data series within a chart
#'
#' @description
#' `xl_chart_series()` names the values a chart plots, and optionally the
#' categories to plot them against and a name for the legend.  A series that
#' plots a column is named after that column's header unless told otherwise.
#'
#' Each range may live on a different sheet from the chart, so it takes an
#' optional `sheet`:
#'
#' * `"Data!B2:B10"` --- an A1 range, sheet-qualified;
#' * `list(cols = "revenue")` --- resolved against the chart's own sheet;
#' * `list(sheet = "Data", cols = "revenue")` --- against another sheet;
#' * `list(header = "revenue")` --- that column's header cell, which is where
#'   a series name usually lives.
#'
#' A range that selects no data is an error rather than an empty chart.
#'
#' @param values The range holding the numbers to plot.
#' @param categories The range holding the labels to plot them against.  Omit
#'   for a chart that numbers its points.
#' @param name The series name, shown in the legend.  Left unset, a series that
#'   plots a column takes its name from that column's header cell, which is
#'   what Excel does when you chart a column along with its header; `FALSE`
#'   leaves it unnamed.  A string is always taken literally --- a series may
#'   legitimately be called `"Q1!"` --- so to take the name from another cell,
#'   give a range spec: `name = list(header = "cost")` for a different column's
#'   header, or `name = list(rows = 1, cols = 1)` for a data cell.
#' @param format An [xl_format] styling the series --- its line and fill.  See
#'   [xl_chart()] for which format properties a chart can express.
#' @param marker An [xl_chart_marker()] drawn at each point.
#' @param labels An [xl_chart_labels()] printing the numbers beside the points.
#' @param trendline An [xl_chart_trendline()] fitted through the series.
#' @param x_error_bars,y_error_bars An [xl_chart_error_bars()] on each point.
#' @param points An [xl_format()] per point, as a list, styling individual
#'   points --- one slice of a pie, one bar of a column chart.  `NULL` in the
#'   list leaves that point as it is.
#' @param smooth Draw the line smoothed.  Line and scatter charts only.
#' @param invert_if_negative Fill negative values with the inverse colour.
#' @return An `xl_chart_series` object.
#' @family images and charts
#' @seealso [xl_chart]
#' @export
#' @examples
#' xl_chart_series(values = list(cols = "revenue"))
#' xl_chart_series(values = "Data!B2:B10", categories = "Data!A2:A10",
#'                 name = "2024")
xl_chart_series <- function(values, categories = NULL, name = NULL,
                            format = NULL, marker = NULL, labels = NULL,
                            trendline = NULL, x_error_bars = NULL,
                            y_error_bars = NULL, points = NULL, smooth = NA,
                            invert_if_negative = NA) {
  if (missing(values) || is.null(values))
    stop("`values` must name the range holding the numbers to plot",
         call. = FALSE)
  if (!is.null(format) && !is_xl_format(format))
    stop("`format` must be an xl_format object", call. = FALSE)
  # Translated again at write time, when the payload is needed; here too so
  # that a format a series cannot use is refused where it was written.
  .chart_format_payload(format, "format",
                        accept = c("line", "fill", "pattern"),
                        part = "a chart series")
  parts <- list(marker = marker, labels = labels, trendline = trendline,
                x_error_bars = x_error_bars, y_error_bars = y_error_bars)
  for (k in names(parts))
    if (!is.null(parts[[k]]) && !inherits(parts[[k]], .PART_CLASS[[k]]))
      stop(sprintf("`%s` must be an %s() object", k,
                   sub("^xl", "xl", .PART_CLASS[[k]])), call. = FALSE)
  nm <- if (is.null(name)) NULL
        else if (identical(name, FALSE)) list(off = TRUE)
        else if (is.character(name) && length(name) == 1L && !is.na(name))
          list(text = name)
        else .chart_range(name, "name")
  structure(.drop_null(list(
    values = .chart_range(values, "values"),
    categories = .chart_range(categories, "categories"),
    name = nm, format = format,
    marker = marker, labels = labels, trendline = trendline,
    x_error_bars = x_error_bars, y_error_bars = y_error_bars,
    points = .points_payload(points, "points"),
    smooth = .val_flag(smooth, "smooth"),
    invert_if_negative = .val_flag(invert_if_negative, "invert_if_negative")
  )), class = "xl_chart_series")
}

#' @export
print.xl_chart_series <- function(x, ...) {
  p <- unclass(x)
  cat(sprintf("<xl_chart_series%s>\n",
              if (!is.null(p[["name"]][["text"]]))
                paste0(": ", p[["name"]][["text"]]) else ""))
  invisible(x)
}

#' Add a chart to a worksheet
#'
#' @description
#' `xl_chart()` builds a chart from one or more [xl_chart_series()] and places
#' it on a sheet, anchored to a cell.  Pass one or a list of them as
#' `xl_sheet(chart = )`.
#'
#' Placement works exactly as it does for [xl_image()] --- `at`, `scale`,
#' `offset`, `position`, `description` and `decorative` mean the same things,
#' because libxlsxwriter describes both with the same fields.
#'
#' @section What a chart type supports:
#' Excel silently drops options a chart type cannot use, so writexl refuses them
#' instead, naming the types that would work.  Pie and doughnut charts have no
#' axes; only a doughnut has a hole; only pie and doughnut rotate; up-down bars
#' and high-low lines are line-only; the series gap and overlap are bar and
#' column only.
#'
#' @param type The chart type: `"column"`, `"bar"`, `"line"`, `"pie"`,
#'   `"doughnut"`, `"area"`, `"scatter"`, `"radar"`, and the stacked,
#'   percent-stacked, smoothed and marker variants.
#' @param series One [xl_chart_series()], or a list of them.  Every series of a
#'   scatter chart must have `categories`, which are its x axis.
#' @param title The chart title.  A string is always taken literally, so to
#'   take the title from a cell give a range spec ---
#'   `list(header = "revenue")` for a column's header cell, or
#'   `list(rows = 1, cols = 1)` for a data cell.  `FALSE` removes the title
#'   Excel would otherwise generate.
#' @param title_format An [xl_format()] styling the title text.  A title is
#'   text, so only the [xl_font()] group applies.
#' @param title_layout Where to put the title by hand, as `c(x, y)` fractions
#'   of the chart.  Excel places it for you otherwise.
#' @param title_overlay Let the title sit over the plot rather than above it.
#' @param legend An [xl_chart_legend()] moving, styling or removing the legend.
#' @param data_table An [xl_chart_table()] printing the plotted numbers in a
#'   grid beneath the chart.
#' @param plot_area_format,chart_area_format An [xl_format()] styling the plot
#'   area --- the panel the data is drawn in --- and the chart area around it:
#'   [xl_border()] for the line, [xl_fill()] for the fill or pattern.
#' @param plot_area_layout Where to put the plot area by hand, as `c(x, y)`
#'   or `c(x, y, width, height)` fractions of the chart.
#' @param drop_lines Drop lines from each point to the category axis: `TRUE`,
#'   or an [xl_format()] giving the line to draw them with.  Line and area
#'   charts.
#' @param high_low_lines A line joining the highest and lowest series at each
#'   category, the same way.  Line charts.
#' @param up_down_bars Bars between the first and last series at each category:
#'   `TRUE`, or `list(up = , down = )` with an [xl_format()] for either bar.
#'   Line charts.
#' @param hole_size The size of a doughnut's hole, 10 to 90 percent.
#' @param rotation Where a pie or doughnut starts, 0 to 360 degrees clockwise
#'   from the top.
#' @param series_gap The gap between category groups on a bar or column chart,
#'   0 to 500 percent of a bar's width.
#' @param series_overlap How far bars of one category overlap, -100 to 100
#'   percent.  100 stacks them, -100 pushes them apart.
#' @param show_blanks What an empty cell does to the plot: leave a `"gap"`,
#'   plot it as `"zero"`, or join across it with `"connected"`.
#' @param show_hidden_data Plot data from rows and columns that are hidden.
#'   Excel leaves them out otherwise.
#' @param x_axis,y_axis An [xl_chart_axis()] describing that axis.  Pie and
#'   doughnut charts have none, and several axis options apply to a value or a
#'   category axis only --- see [xl_chart_axis()].
#' @param at The cell the chart's top-left corner is anchored to.
#' @param scale Scale factor: one number for both axes, or `c(x, y)`.
#' @param offset Offset from the anchor cell's corner in pixels, as `c(x, y)`.
#' @param position How the chart behaves when rows and columns change size; see
#'   [xl_image()].
#' @param description Alt text, for screen readers.
#' @param decorative Mark the chart as decorative, so screen readers skip it.
#' @param style Excel's built-in chart style, 1--48.
#' @return An `xl_chart` object.
#' @family images and charts
#' @seealso [xl_chart_series], [xl_sheet]
#' @export
#' @examples
#' xl_chart("column", xl_chart_series(values = list(cols = "revenue")))
#' xl_chart("pie", xl_chart_series(values = "Data!B2:B5"), title = "Share")
xl_chart <- function(type, series, title = NULL, title_format = NULL,
                     title_layout = NULL, title_overlay = NA,
                     x_axis = NULL, y_axis = NULL,
                     legend = NULL, data_table = NULL,
                     plot_area_format = NULL, plot_area_layout = NULL,
                     chart_area_format = NULL,
                     drop_lines = NA, high_low_lines = NA, up_down_bars = NA,
                     hole_size = NA, rotation = NA,
                     series_gap = NA, series_overlap = NA,
                     show_blanks = NULL, show_hidden_data = NA,
                     at = "A1", scale = 1,
                     offset = NULL, position = "move_and_size",
                     description = NULL, decorative = FALSE, style = NA) {
  ty <- .val_enum(type, names(.LXW_CHART_TYPE), "type")
  if (is.null(ty))
    stop("`type` must name a chart type, e.g. \"column\"", call. = FALSE)
  if (missing(series))
    stop("`series` must give at least one xl_chart_series() to plot",
         call. = FALSE)
  ss <- .chart_series_list(series)
  if (!length(ss))
    stop("a chart needs at least one series", call. = FALSE)
  .check_axis(x_axis, "x", ty, "x_axis")
  .check_axis(y_axis, "y", ty, "y_axis")
  for (nm in c("legend", "data_table"))
    if (!is.null(get(nm)) && !inherits(get(nm), paste0("xl_chart_",
                                                       sub("data_", "", nm))))
      stop(sprintf("`%s` must be an xl_chart_%s() object", nm,
                   sub("data_", "", nm)), call. = FALSE)

  # the six that belong to one family, checked against the matrix below
  drop_l <- .chart_line_switch(drop_lines, "drop_lines")
  high_l <- .chart_line_switch(high_low_lines, "high_low_lines")
  updown <- .chart_up_down(up_down_bars, "up_down_bars")
  hole <- .val_int(hole_size, "hole_size", min = 10, max = 90)
  rot <- .val_int(rotation, "rotation", min = 0, max = 360)
  gap <- .val_int(series_gap, "series_gap", min = 0, max = 500)
  ovl <- .val_int(series_overlap, "series_overlap", min = -100, max = 100)
  for (f in list(list(drop_l, "drop_lines"), list(high_l, "high_low_lines"),
                 list(updown, "up_down_bars"), list(hole, "hole_size"),
                 list(rot, "rotation"), list(gap, "series_gap"),
                 list(ovl, "series_overlap")))
    if (!is.null(f[[1L]])) .check_chart_feature(ty, f[[2L]], f[[2L]])
  for (i in seq_along(ss))
    .check_series_parts(ss[[i]], ty, sprintf("series[[%d]]", i))

  for (i in seq_along(ss)) {
    p <- unclass(ss[[i]])
    if (isTRUE(p[["smooth"]]))
      .check_chart_feature(ty, "smooth", sprintf("series[[%d]]$smooth", i))
    # A scatter plots x against y, so a series with no categories has no x
    # values.  libxlsxwriter requires them too, but only checks when the range
    # is passed to chart_add_series() as a string; set through
    # chart_series_set_values() it reaches _chart_write_x_val(), which unlike
    # _chart_write_cat() has no NULL guard, and strpbrk() is handed the NULL
    # formula.
    if (identical(.CHART_FAMILY(ty), "scatter") && is.null(p[["categories"]]))
      stop(sprintf(paste0("`series[[%d]]` has no `categories`, which a \"%s\" ",
                          "chart needs: a scatter plots values against ",
                          "categories, so the categories are its x axis."),
                   i, ty), call. = FALSE)
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

  ttl <- if (is.null(title)) NULL
         else if (identical(title, FALSE)) list(off = TRUE)
         else if (is.character(title) && length(title) == 1L && !is.na(title))
           list(text = title)
         else .chart_range(title, "title")
  # A title's only styling is its font, so a fill or a border here has
  # nowhere to go, and a title switched off has nothing to style.
  if (!is.null(title_format)) {
    if (isTRUE(ttl[["off"]]))
      stop("`title_format` styles the title, but `title = FALSE` removes it",
           call. = FALSE)
    # The automatic title of a single-series chart is Excel's own: with no
    # `title` there is no <c:title> element for the font to attach to.
    if (is.null(ttl))
      stop(paste0("`title_format` styles the title, so give a `title` too. ",
                  "The title Excel generates for a single-series chart is not ",
                  "in the file, so there is nothing to style."), call. = FALSE)
    ttl <- c(if (is.null(ttl)) list() else ttl,
             list(format = .chart_format_payload(title_format, "title_format",
                                                 accept = "font",
                                                 part = "a chart title")))
  }

  ttl <- .chart_title_extras(ttl, title_layout, title_overlay)

  structure(.drop_null(list(
    type = ty, series = ss, title = ttl, x_axis = x_axis, y_axis = y_axis,
    legend = legend, data_table = data_table,
    plot_area_format = .chart_format_payload(
      plot_area_format, "plot_area_format",
      accept = c("line", "fill", "pattern"), part = "a plot area"),
    plot_area_layout = .chart_layout(plot_area_layout,
                                     "plot_area_layout"),
    chart_area_format = .chart_format_payload(
      chart_area_format, "chart_area_format",
      accept = c("line", "fill", "pattern"), part = "a chart area"),
    drop_lines = drop_l, high_low_lines = high_l, up_down_bars = updown,
    hole_size = hole, rotation = rot, series_gap = gap, series_overlap = ovl,
    show_blanks = .val_enum(show_blanks, names(.LXW_CHART_BLANKS),
                            "show_blanks"),
    show_hidden_data = .val_flag(show_hidden_data, "show_hidden_data"),
    at = at, scale = sc, offset = pair(offset, "offset", "of pixels as c(x, y)"),
    position = .val_enum(position, names(.LXW_OBJECT_POSITION), "position"),
    description = if (is.null(description)) NULL else {
      if (!is.character(description) || length(description) != 1L ||
          is.na(description))
        stop("`description` must be a single non-NA string", call. = FALSE)
      description
    },
    decorative = .val_flag(decorative, "decorative"),
    style = .val_int(style, "style", min = 1, max = 48)
  )), class = "xl_chart")
}

#' @export
print.xl_chart <- function(x, ...) {
  p <- unclass(x)
  cat(sprintf("<xl_chart: %s, %d series%s>\n", p[["type"]],
              length(p[["series"]]),
              if (!is.null(p[["title"]][["text"]]))
                paste0(", \"", p[["title"]][["text"]], "\"") else ""))
  invisible(x)
}

# --- Resolving a series range against the workbook ---------------------------
#
# A series may plot data from a sheet other than the one the chart sits on, so
# a range cannot be resolved until every sheet is known.  `sheets` maps sheet
# name to the resolved data frame; `own` is the chart's own sheet, used when the
# range names none.
# The header cell of the single column a values spec names, or NULL when there
# is no one column to take it from.
.series_header_name <- function(values, header_offset) {
  if (header_offset < 1L || is.null(values)) return(NULL)
  spec <- values[["spec"]]
  if (!is.list(spec)) return(NULL)             # an A1 range, not a column spec
  cols <- spec[["cols"]]
  if (is.null(cols) || length(cols) != 1L || anyNA(cols)) return(NULL)
  list(spec = list(header = cols), sheet = values[["sheet"]])
}

.resolve_chart_range <- function(rng, arg, sheets, own, header_offset) {
  if (is.null(rng)) return(NULL)
  spec <- rng[["spec"]]
  sheet <- rng[["sheet"]]
  # a string may carry its own qualifier -- "Data!B2:B10" -- which the shared
  # resolver parses; the list form carries it as an element
  if (is.character(spec) && is.null(sheet)) {
    sp <- .split_sheet_ref(spec)
    sheet <- sp$sheet
    spec <- sp$rest
  }
  target <- if (is.null(sheet)) own else sheet
  if (!target %in% names(sheets))
    stop(sprintf(paste0("`%s` names sheet \"%s\", which is not in the ",
                        "workbook.\n  Sheets are: %s."),
                 arg, target, paste(sprintf("\"%s\"", names(sheets)),
                                    collapse = ", ")), call. = FALSE)
  df <- sheets[[target]]
  if (is.list(spec) && !is.null(spec[["header"]])) {
    if (header_offset < 1L)
      stop(sprintf(paste0("`%s` names a header cell, but this workbook is ",
                          "written without headers (col_names = FALSE), so ",
                          "there is no header row."), arg), call. = FALSE)
    col <- .resolve_col_index(spec[["header"]], names(df), arg)
    # 0-based, and the header is the row above data row 1
    return(list(sheet = target,
                range = as.integer(c(0L, col - 1L, 0L, col - 1L))))
  }
  q <- .xl_resolve_range(spec, arg = arg, df = df,
                         header_offset = header_offset, allow_cell = TRUE)
  # A range outside the data plots nothing, and Excel shows an empty chart with
  # no complaint -- so the emptiness is caught here rather than delivered.
  last_row <- (nrow(df) - 1L) + header_offset
  if (q[1L] > last_row || q[2L] > length(df) - 1L)
    stop(sprintf(paste0("`%s` selects no data: sheet \"%s\" has %d row(s) and ",
                        "%d column(s), and the range starts outside them."),
                 arg, target, nrow(df), length(df)), call. = FALSE)
  if (nrow(df) < 1L)
    stop(sprintf("`%s` names sheet \"%s\", which has no data rows", arg,
                 target), call. = FALSE)
  list(sheet = target, range = as.integer(q))
}

# Resolve a sheet's charts to the C payloads.  Called once the whole workbook is
# known, since a series may point at another sheet.
.resolve_charts <- function(el, df, reg, header_offset, props, sheets, own) {
  if (inherits(el, "xl_chartsheet")) {
    # a chartsheet has no cells, so `own` names a sheet with no columns and a
    # range that does not name its sheet has nothing to resolve against
    .check_chartsheet_ranges(unclass(el)$chart, own)
    cs <- list(unclass(el)$chart)
  } else if (!inherits(el, "xl_sheet")) {
    return(list())
  } else {
    cs <- .chart_list(el$chart)
  }
  lapply(seq_along(cs), function(i) {
    p <- unclass(cs[[i]])
    at <- .xl_resolve_range(p[["at"]], arg = sprintf("chart[[%d]] at", i),
                            df = df, header_offset = header_offset,
                            allow_cell = TRUE)
    if (at[1L] != at[3L] || at[2L] != at[4L])
      stop(sprintf(paste0("`chart[[%d]] at` must name a single cell: a chart ",
                          "is anchored to one cell"), i), call. = FALSE)
    ent <- list(type = unname(.LXW_CHART_TYPE[[p[["type"]]]]),
                row = at[1L], col = at[2L])
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
    if (!is.null(p[["style"]]))       ent$style <- as.integer(p[["style"]])

    ttl <- p[["title"]]
    if (!is.null(ttl)) {
      ent$title_format <- ttl[["format"]]
      if (!is.null(ttl[["layout"]])) ent$title_layout <- ttl[["layout"]]
      if (!is.null(ttl[["overlay"]])) ent$title_overlay <- ttl[["overlay"]]
      if (isTRUE(ttl[["off"]]))        ent$title_off <- 1L
      else if (!is.null(ttl[["text"]])) ent$title <- ttl[["text"]]
      # `title_format` alone leaves Excel's automatic title in place and only
      # styles it, so there is no range to resolve
      else if (!is.null(ttl[["spec"]])) {
        r <- .resolve_chart_range(ttl, sprintf("chart[[%d]] title", i),
                                  sheets, own, header_offset)
        ent$title_sheet <- r$sheet
        ent$title_range <- r$range
      }
    }

    ent$legend     <- .legend_payload(p[["legend"]])
    ent$data_table <- .chart_table_payload(p[["data_table"]])
    for (k in c("plot_area_format", "chart_area_format")) {
      f <- p[[k]]
      if (is.null(f)) next
      pre <- sub("_format$", "", k)
      ent[[paste0(pre, "_line")]]    <- f[["line"]]
      ent[[paste0(pre, "_fill")]]    <- f[["fill"]]
      ent[[paste0(pre, "_pattern")]] <- f[["pattern"]]
    }
    ent$plot_area_layout <- p[["plot_area_layout"]]
    ent$drop_lines     <- p[["drop_lines"]]
    ent$high_low_lines <- p[["high_low_lines"]]
    ent$up_down_bars   <- p[["up_down_bars"]]
    ent$hole_size      <- p[["hole_size"]]
    ent$rotation       <- p[["rotation"]]
    ent$series_gap     <- p[["series_gap"]]
    ent$series_overlap <- p[["series_overlap"]]
    if (!is.null(p[["show_blanks"]]))
      ent$show_blanks <- unname(.LXW_CHART_BLANKS[[p[["show_blanks"]]]])
    if (isTRUE(p[["show_hidden_data"]])) ent$show_hidden_data <- 1L

    ent$x_axis <- .axis_payload(p[["x_axis"]], sprintf("chart[[%d]] x_axis", i),
                                sheets, own, header_offset)
    ent$y_axis <- .axis_payload(p[["y_axis"]], sprintf("chart[[%d]] y_axis", i),
                                sheets, own, header_offset)

    ent$series <- lapply(seq_along(p[["series"]]), function(k) {
      q <- unclass(p[["series"]][[k]])
      where <- function(what) sprintf("chart[[%d]] series[[%d]] %s", i, k, what)
      v <- .resolve_chart_range(q[["values"]], where("values"), sheets, own,
                                header_offset)
      se <- list(values_sheet = v$sheet, values_range = v$range)
      cat_ <- .resolve_chart_range(q[["categories"]], where("categories"),
                                   sheets, own, header_offset)
      if (!is.null(cat_)) {
        se$categories_sheet <- cat_$sheet
        se$categories_range <- cat_$range
      }
      # Unnamed, a series is "Series 1" in the legend while the column it
      # plots already has a label in its header -- the name Excel itself uses.
      # Column specs only: an A1 range says nothing about whether the row
      # above it is a header.
      nm <- q[["name"]]
      if (is.null(nm)) nm <- .series_header_name(q[["values"]], header_offset)
      if (!is.null(nm) && !isTRUE(nm[["off"]])) {
        if (!is.null(nm[["text"]])) se$name <- nm[["text"]]
        else {
          r <- .resolve_chart_range(nm, where("name"), sheets, own,
                                    header_offset)
          se$name_sheet <- r$sheet
          se$name_range <- r$range
        }
      }
      se$marker       <- .marker_payload(q[["marker"]])
      se$labels       <- .labels_payload(q[["labels"]])
      se$trendline    <- .trendline_payload(q[["trendline"]])
      se$x_error_bars <- .error_bars_payload(q[["x_error_bars"]])
      se$y_error_bars <- .error_bars_payload(q[["y_error_bars"]])
      se$points       <- q[["points"]]
      if (isTRUE(q[["smooth"]]))             se$smooth <- 1L
      if (isTRUE(q[["invert_if_negative"]])) se$invert_if_negative <- 1L
      # a series is styled by its line and fill; see .chart_format_payload()
      if (!is.null(q[["format"]]))
        se$format <- .chart_format_payload(
          q[["format"]], where("format"),
          accept = c("line", "fill", "pattern"), part = "a chart series")
      se
    })
    ent
  })
}

# Normalise one series or a list of them.
.chart_series_list <- function(series, arg = "series") {
  if (is.null(series)) return(list())
  ss <- if (inherits(series, "xl_chart_series")) list(series) else series
  if (!is.list(ss))
    stop(sprintf("`%s` must be an xl_chart_series object or a list of them",
                 arg), call. = FALSE)
  for (i in seq_along(ss))
    if (!inherits(ss[[i]], "xl_chart_series"))
      stop(sprintf("`%s[[%d]]` must be an xl_chart_series object", arg, i),
           call. = FALSE)
  ss
}

# Normalise one chart or a list of them.
.chart_list <- function(chart, arg = "chart") {
  if (is.null(chart)) return(list())
  cs <- if (inherits(chart, "xl_chart")) list(chart) else chart
  if (!is.list(cs))
    stop(sprintf("`%s` must be an xl_chart object or a list of them", arg),
         call. = FALSE)
  for (i in seq_along(cs))
    if (!inherits(cs[[i]], "xl_chart"))
      stop(sprintf("`%s[[%d]]` must be an xl_chart object", arg, i),
           call. = FALSE)
  cs
}

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
# Verified against libxlsxwriter 1.2.4 before any of this was written.
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

.LXW_CHART_LEGEND <- c(
  right = 1L, left = 2L, top = 3L, bottom = 4L, top_right = 5L,
  overlay_right = 6L, overlay_left = 7L, overlay_top_right = 8L, none = 9L
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
    unknown <- setdiff(nms, c("sheet", "rows", "cols"))
    if (length(unknown))
      stop(sprintf("unknown `%s` element(s): %s", arg,
                   paste(unknown, collapse = ", ")), call. = FALSE)
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
#' categories to plot them against and a name for the legend.
#'
#' Each range may live on a different sheet from the chart, so it takes an
#' optional `sheet`:
#'
#' * `"Data!B2:B10"` --- an A1 range, sheet-qualified;
#' * `list(cols = "revenue")` --- resolved against the chart's own sheet;
#' * `list(sheet = "Data", cols = "revenue")` --- against another sheet.
#'
#' A range that selects no data is an error rather than an empty chart.
#'
#' @param values The range holding the numbers to plot.
#' @param categories The range holding the labels to plot them against.  Omit
#'   for a chart that numbers its points.
#' @param name The series name, shown in the legend.  A string is always taken
#'   literally --- a series may legitimately be called `"Q1!"` --- so to take
#'   the name from a cell, give a range spec:
#'   `name = list(sheet = "Data", rows = 1, cols = 1)`.
#' @param format An [xl_format] styling the series --- its line and fill.  See
#'   [xl_chart()] for which format properties a chart can express.
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
                            format = NULL, smooth = NA,
                            invert_if_negative = NA) {
  if (missing(values) || is.null(values))
    stop("`values` must name the range holding the numbers to plot",
         call. = FALSE)
  if (!is.null(format) && !is_xl_format(format))
    stop("`format` must be an xl_format object", call. = FALSE)
  nm <- if (is.null(name)) NULL
        else if (is.character(name) && length(name) == 1L && !is.na(name))
          list(text = name)
        else .chart_range(name, "name")
  structure(.drop_null(list(
    values = .chart_range(values, "values"),
    categories = .chart_range(categories, "categories"),
    name = nm, format = format,
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
#' @param series One [xl_chart_series()], or a list of them.
#' @param title The chart title: a string, or a range holding one.  `FALSE`
#'   removes the title Excel would otherwise generate.
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
xl_chart <- function(type, series, title = NULL, at = "A1", scale = 1,
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

  for (i in seq_along(ss)) {
    p <- unclass(ss[[i]])
    if (isTRUE(p[["smooth"]]))
      .check_chart_feature(ty, "smooth", sprintf("series[[%d]]$smooth", i))
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

  structure(.drop_null(list(
    type = ty, series = ss, title = ttl,
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
  if (!inherits(el, "xl_sheet")) return(list())
  cs <- .chart_list(el$chart)
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
      if (isTRUE(ttl[["off"]]))        ent$title_off <- 1L
      else if (!is.null(ttl[["text"]])) ent$title <- ttl[["text"]]
      else {
        r <- .resolve_chart_range(ttl, sprintf("chart[[%d]] title", i),
                                  sheets, own, header_offset)
        ent$title_sheet <- r$sheet
        ent$title_range <- r$range
      }
    }

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
      nm <- q[["name"]]
      if (!is.null(nm)) {
        if (!is.null(nm[["text"]])) se$name <- nm[["text"]]
        else {
          r <- .resolve_chart_range(nm, where("name"), sheets, own,
                                    header_offset)
          se$name_sheet <- r$sheet
          se$name_range <- r$range
        }
      }
      if (isTRUE(q[["smooth"]]))             se$smooth <- 1L
      if (isTRUE(q[["invert_if_negative"]])) se$invert_if_negative <- 1L
      if (!is.null(q[["format"]]))
        se$format_id <- .register_format(reg,
          merge_xl_format(props$default_format, q[["format"]]))
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

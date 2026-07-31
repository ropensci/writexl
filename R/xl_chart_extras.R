# =============================================================================
# The chart itself: legend, data table, plot and chart areas, and the options
# that belong to one family of chart
# =============================================================================
#
# Two constructors -- the legend and the data table -- and a set of xl_chart()
# arguments for the rest, which are single values rather than objects with
# parts of their own.
#
# Six apply to one family of chart only: the doughnut hole, pie rotation,
# up-down bars, high-low lines, drop lines, and the bar gap and overlap.
# .CHART_FEATURE_FAMILIES in xl_chart.R holds which, since Excel drops the
# rest without a word.
# -----------------------------------------------------------------------------

.LXW_CHART_BLANKS <- c(gap = 0L, zero = 1L, connected = 2L)

# --- The legend ---------------------------------------------------------------

#' A chart's legend
#'
#' @description
#' `xl_chart_legend()` moves, styles or removes the legend, and can leave
#' individual series out of it.
#'
#' @param position Where the legend sits: `"right"` (Excel's default),
#'   `"left"`, `"top"`, `"bottom"`, `"top_right"`, the `"overlay_*"` variants
#'   that let the legend sit over the plot, or `"none"` to remove it.
#' @param format An [xl_format()] styling the legend text --- the [xl_font()]
#'   group only, since libxlsxwriter gives a legend a font and nothing else.
#' @param layout Where to put the legend by hand, as `c(x, y)` or
#'   `c(x, y, width, height)` --- fractions of the chart, each above 0 and at
#'   most 1.  Excel places it for you otherwise.  `at` is a cell everywhere
#'   else in writexl, so a chart's own fractions are a `layout`.
#' @param delete_series Series to leave out of the legend, by position: `2`
#'   drops the second series' entry while still plotting it.  This is how a
#'   trendline or a helper series is kept out of the key.
#' @return An `xl_chart_legend` object.
#' @family images and charts
#' @seealso [xl_chart]
#' @export
#' @examples
#' xl_chart_legend(position = "bottom")
#' xl_chart_legend(position = "none")
#' xl_chart_legend(delete_series = 2)
xl_chart_legend <- function(position = NULL, format = NULL,
                            layout = NULL, delete_series = NULL) {
  ds <- delete_series
  if (!is.null(ds)) {
    if (!is.numeric(ds) || !length(ds) || anyNA(ds) || any(ds < 1))
      stop(paste0("`delete_series` must be the positions of the series to ",
                  "leave out of the legend, counting from 1"), call. = FALSE)
    ds <- as.integer(ds) - 1L                  # C counts from 0
  }
  structure(.drop_null(list(
    position = .val_enum(position, names(.LXW_CHART_LEGEND), "position"),
    format = .chart_format_payload(format, "format", accept = "font",
                                   part = "a chart legend"),
    layout = .chart_layout(layout, "layout"),
    delete_series = ds
  )), class = "xl_chart_legend")
}

# --- The data table -----------------------------------------------------------

#' The table of values under a chart
#'
#' @description
#' `xl_chart_table()` prints the plotted numbers in a grid beneath the chart,
#' which is Excel's "Data Table" chart element.  It is given to [xl_chart()] as
#' `data_table`.
#'
#' Naming none of the grid options leaves Excel's own: horizontal, vertical and
#' outline borders drawn, and no legend keys.
#'
#' @param show_keys Print each series' legend swatch in the table.
#' @param horizontal,vertical,outline Which borders the grid draws.
#' @param format An [xl_format()] styling the table's text --- the [xl_font()]
#'   group only.
#' @return An `xl_chart_table` object.
#' @family images and charts
#' @seealso [xl_chart]
#' @export
#' @examples
#' xl_chart_table()
#' xl_chart_table(show_keys = TRUE, vertical = FALSE)
xl_chart_table <- function(show_keys = NA, horizontal = NA, vertical = NA,
                           outline = NA, format = NULL) {
  structure(.drop_null(list(
    show_keys = .val_flag(show_keys, "show_keys"),
    horizontal = .val_flag(horizontal, "horizontal"),
    vertical = .val_flag(vertical, "vertical"),
    outline = .val_flag(outline, "outline"),
    format = .chart_format_payload(format, "format", accept = "font",
                                   part = "a chart data table")
  )), class = "xl_chart_table")
}

# --- Shared helpers -----------------------------------------------------------

# lxw_chart_layout: fractions of the chart's own width and height.  Two numbers
# place a thing, four also size it.
.chart_layout <- function(x, arg, size_ok = TRUE) {
  if (is.null(x)) return(NULL)
  n <- if (size_ok) c(2L, 4L) else 2L
  if (!is.numeric(x) || !length(x) %in% n || anyNA(x) ||
      any(x <= 0) || any(x > 1))
    stop(sprintf(paste0("`%s` must be %s, each above 0 and at most 1 -- they ",
                        "are fractions of the chart's width and height"),
                 arg, if (size_ok) "c(x, y) or c(x, y, width, height)"
                      else "c(x, y)"), call. = FALSE)
  as.numeric(x)
}

# TRUE, or an xl_format giving the line these are drawn with.
.chart_line_switch <- function(x, arg) {
  if (.is_unset(x)) return(NULL)
  if (isTRUE(x)) return(list(on = 1L))
  if (identical(x, FALSE)) return(NULL)
  if (is_xl_format(x)) {
    f <- .chart_format_payload(x, arg, accept = "line", part = arg)
    return(.drop_null(list(on = 1L, line = f[["line"]])))
  }
  stop(sprintf(paste0("`%s` must be TRUE, or an xl_format() giving the line to ",
                      "draw them with"), arg), call. = FALSE)
}

# TRUE, or list(up = , down = ) with a format for each bar.
.chart_up_down <- function(x, arg) {
  if (.is_unset(x)) return(NULL)
  if (isTRUE(x)) return(list(on = 1L))
  if (identical(x, FALSE)) return(NULL)
  if (!is.list(x) || !length(x) ||
      length(setdiff(names(x), c("up", "down"))) || is.null(names(x)))
    stop(sprintf(paste0("`%s` must be TRUE, or list(up = , down = ) with an ",
                        "xl_format() for either bar"), arg), call. = FALSE)
  out <- list(on = 1L)
  for (side in c("up", "down")) {
    f <- x[[side]]
    if (is.null(f)) next
    p <- .chart_format_payload(f, sprintf("%s$%s", arg, side),
                               accept = c("line", "fill", "pattern"),
                               part = "an up-down bar")
    out[[paste0(side, "_line")]] <- p[["line"]]
    out[[paste0(side, "_fill")]] <- p[["fill"]]
  }
  .drop_null(out)
}

# --- Payloads -----------------------------------------------------------------

.legend_payload <- function(l) {
  if (is.null(l)) return(NULL)
  p <- unclass(l)
  ent <- list()
  if (!is.null(p[["position"]]))
    ent$position <- unname(.LXW_CHART_LEGEND[[p[["position"]]]])
  ent$font <- p[["format"]]
  if (!is.null(p[["layout"]])) ent$layout <- p[["layout"]]
  ent$delete_series <- p[["delete_series"]]
  .drop_null(ent)
}

.chart_table_payload <- function(t) {
  if (is.null(t)) return(NULL)
  p <- unclass(t)
  ent <- list(on = 1L)
  grid <- c("horizontal", "vertical", "outline", "show_keys")
  if (any(vapply(grid, function(k) !is.null(p[[k]]), logical(1))))
    # Excel's own defaults, for whichever were not named
    ent$grid <- as.integer(c(
      isTRUE(p[["horizontal"]] %||% TRUE), isTRUE(p[["vertical"]] %||% TRUE),
      isTRUE(p[["outline"]] %||% TRUE), isTRUE(p[["show_keys"]] %||% FALSE)))
  ent$font <- p[["format"]]
  .drop_null(ent)
}

# The title carries its layout and overlay alongside its text and font, so
# that .chart_list() has one place to look.
.chart_title_extras <- function(ttl, layout, overlay) {
  lay <- .chart_layout(layout, "title_layout", size_ok = FALSE)
  ov <- .val_flag(overlay, "title_overlay")
  if (is.null(lay) && is.null(ov)) return(ttl)
  if (isTRUE(ttl[["off"]]))
    stop(paste0("`title_layout` and `title_overlay` place the title, but ",
                "`title = FALSE` removes it"), call. = FALSE)
  if (is.null(ttl))
    stop(paste0("`title_layout` and `title_overlay` place the title, so give ",
                "a `title` too. The title Excel generates for a single-series ",
                "chart is not in the file, so there is nothing to place."),
         call. = FALSE)
  c(ttl, .drop_null(list(layout = lay,
                         overlay = if (is.null(ov)) NULL else as.integer(ov))))
}

#' @export
print.xl_chart_legend <- function(x, ...) .print_part(x, "xl_chart_legend")

#' @export
print.xl_chart_table <- function(x, ...) .print_part(x, "xl_chart_table")

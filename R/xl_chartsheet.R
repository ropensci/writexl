# =============================================================================
# Chartsheets: a sheet that is one chart and nothing else
# =============================================================================
#
# A chartsheet holds a single chart full-page and has no cells, so it takes no
# data frame and none of the worksheet options that address cells.
# libxlsxwriter reflects that: lxw_chartsheet wraps a worksheet but exposes
# only 17 functions, against the worksheet's several hundred.
#
# Its chart therefore has to plot another sheet.  Every range in the series must
# name one, since there is no data on the chartsheet for a bare
# list(cols = ) spec to resolve against.
#
# The subsets a chartsheet supports are named here rather than left to Excel:
# xl_page_setup() has 23 options and a chartsheet has setters for five,
# xl_sheet_view() has nine and a chartsheet has four.  The rest are refused by
# name.
# -----------------------------------------------------------------------------

# The xl_page_setup() options a chartsheet can express: chartsheet_set_paper(),
# _set_margins(), _set_landscape()/_set_portrait(), _set_header(), _set_footer()
# and their _opt() variants carrying the margins.
.CHARTSHEET_PAGE <- c("orientation", "paper", "margins", "header", "footer",
                      "header_margin", "footer_margin")

# The xl_sheet_view() options it can express: chartsheet_activate(),
# _select(), _hide() and _set_first_sheet().
.CHARTSHEET_VIEW <- c("active", "selected", "visible", "first_tab")

#' A sheet holding a single chart
#'
#' @description
#' `xl_chartsheet()` is a worksheet-sized chart: a tab of its own holding one
#' chart and no cells.  Give it to [write_xlsx()] in place of a data frame.
#'
#' Because a chartsheet has no cells, every range in the chart's series must
#' name the sheet it plots --- `list(sheet = "Data", cols = "revenue")` or
#' `"Data!B2:B10"`.  A bare `list(cols = )` has nothing to resolve against and
#' is refused.
#'
#' A chartsheet supports only part of what a worksheet does, and the parts it
#' does not are refused rather than dropped: of [xl_page_setup()] it takes the
#' orientation, paper size, margins and the header and footer; of
#' [xl_sheet_view()] it takes `active`, `selected`, `visible` and `first_tab`.
#'
#' @param chart The [xl_chart()] to fill the sheet with.
#' @param tab_color The colour of the sheet tab.
#' @param zoom The zoom level as a percentage, 10 to 400.
#' @param protect Worksheet protection, as [xl_sheet()] takes it.
#' @param page An [xl_page_setup()] describing how it prints.
#' @param view An [xl_sheet_view()] setting the tab state.
#' @return An `xl_chartsheet` object.
#' @family images and charts
#' @seealso [xl_chart], [xl_sheet]
#' @export
#' @examples
#' sales <- data.frame(quarter = c("Q1", "Q2"), revenue = c(10, 25))
#' chart <- xl_chart("column",
#'                   xl_chart_series(values = list(sheet = "Data",
#'                                                 cols = "revenue")))
#' write_xlsx(list(Data = sales, Overview = xl_chartsheet(chart)),
#'            tempfile(fileext = ".xlsx"))
xl_chartsheet <- function(chart, tab_color = NULL, zoom = NA, protect = NULL,
                          page = NULL, view = NULL) {
  if (missing(chart) || is.null(chart))
    stop("`chart` must be the xl_chart() to fill the sheet with", call. = FALSE)
  .validate_protect(protect)
  ch <- .chart_list(chart)
  if (length(ch) != 1L)
    stop(sprintf(paste0("a chartsheet holds exactly one chart, not %d. Put ",
                        "the others on worksheets with xl_sheet(chart = )."),
                 length(ch)), call. = FALSE)

  structure(.drop_null(list(
    chart = ch[[1L]],
    view = .chartsheet_subset(view, "view", "xl_sheet_view",
                              .CHARTSHEET_VIEW),
    page = .chartsheet_subset(page, "page", "xl_page_setup",
                              .CHARTSHEET_PAGE),
    protect = protect,
    tab_color = .val_color(tab_color, "tab_color"),
    zoom = .val_int(zoom, "zoom", min = 10, max = 400)
  )), class = "xl_chartsheet")
}

# Keep the options a chartsheet has a setter for, and name the ones it does not.
.chartsheet_subset <- function(x, arg, cls, keep) {
  if (is.null(x)) return(NULL)
  if (!inherits(x, cls))
    stop(sprintf("`%s` must be an %s() object", arg, cls), call. = FALSE)
  # xl_sheet_view() keeps a slot per option, unset ones NULL, so only what was
  # actually set can be complained about
  p <- Filter(Negate(is.null), unclass(x))
  extra <- setdiff(names(p), keep)
  if (length(extra))
    stop(sprintf(paste0("`%s` sets %s, which a chartsheet has no setter for; ",
                        "Excel would ignore it.\n  A chartsheet takes: %s."),
                 arg, paste(sprintf("`%s`", extra), collapse = ", "),
                 paste(keep, collapse = ", ")), call. = FALSE)
  p
}

#' @export
print.xl_chartsheet <- function(x, ...) {
  p <- unclass(x)
  cat("<xl_chartsheet> ", unclass(p$chart)$type, " chart\n", sep = "")
  set <- setdiff(names(p), "chart")
  if (length(set)) cat("  set:", paste(set, collapse = ", "), "\n")
  invisible(x)
}

# --- Resolution --------------------------------------------------------------

# A chartsheet has no cells, so it stands in the data-frame list as a data frame
# with no columns.  Nothing writes rows for it: the C side branches on the
# plan's `chartsheet` flag before the row loop.
.chartsheet_placeholder <- function() data.frame()

# The plan C reads for one chartsheet.  `charts` carries the single chart in the
# same slot a worksheet uses, so .check_drawing_order() sees it without knowing
# the difference.
.resolve_chartsheet_plan <- function(el, charts) {
  p <- unclass(el)
  ent <- list(chartsheet = 1L, charts = charts)
  vw <- p[["view"]]
  if (!is.null(vw)) {
    if (isTRUE(vw$active))       ent$activate <- 1L
    if (isTRUE(vw$selected))     ent$select <- 1L
    if (identical(vw$visible, FALSE)) ent$hide <- 1L
    if (isTRUE(vw$first_tab))    ent$first_sheet <- 1L
  }
  pg <- p[["page"]]
  if (!is.null(pg)) {
    if (!is.null(pg$orientation))
      ent$landscape <- as.integer(identical(pg$orientation, "landscape"))
    ent$paper <- pg$paper
    if (!is.null(pg$margins)) ent$margins <- unlist(pg$margins, use.names = FALSE)
    ent$header <- pg$header
    ent$footer <- pg$footer
    ent$header_margin <- pg$header_margin
    ent$footer_margin <- pg$footer_margin
  }
  ent$tab_color <- p[["tab_color"]]
  ent$zoom <- p[["zoom"]]
  if (!is.null(p[["protect"]])) {
    pr <- .resolve_protect(p[["protect"]])
    ent$protect <- pr$flag
    ent$protect_password <- pr$password
    ent$protect_options <- pr$options
  }
  .drop_null(ent)
}

# Every range a chartsheet's chart plots must name its sheet: there is nothing
# on the chartsheet itself for a bare list(cols = ) or an unqualified "B2:B10"
# to resolve against.
.check_chartsheet_ranges <- function(chart, own) {
  named <- function(r) {
    if (is.null(r)) return(TRUE)
    if (!is.null(r[["sheet"]])) return(TRUE)
    is.character(r[["spec"]]) && !is.null(.split_sheet_ref(r[["spec"]])$sheet)
  }
  p <- unclass(chart)
  parts <- list(title = p[["title"]])
  for (i in seq_along(p[["series"]])) {
    q <- unclass(p[["series"]][[i]])
    for (k in c("values", "categories", "name"))
      parts[[sprintf("series[[%d]]$%s", i, k)]] <- q[[k]]
  }
  for (nm in names(parts)) {
    r <- parts[[nm]]
    # a literal title or series name carries no range at all
    if (is.null(r) || !is.null(r[["text"]]) || isTRUE(r[["off"]])) next
    if (named(r)) next
    stop(sprintf(paste0("`%s` does not name a sheet, and chartsheet \"%s\" ",
                        "has no cells of its own to resolve it against.\n  ",
                        "Give the sheet holding the data: ",
                        "list(sheet = \"Data\", cols = \"revenue\") or ",
                        "\"Data!B2:B10\"."),
                 nm, own), call. = FALSE)
  }
  invisible(NULL)
}

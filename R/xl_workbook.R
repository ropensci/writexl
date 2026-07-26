# =============================================================================
# Workbook layer: xl_properties() and xl_workbook()
# =============================================================================

# The default number format for POSIXct columns.  Shared with
# .resolve_timezones(), which swaps it out when the time zone is dropped, so
# the two definitions cannot drift apart.
.default_datetime_format <- function() xl_num_format("yyyy-mm-dd HH:mm:ss UTC")

#' Workbook properties, defaults, and metadata
#'
#' @description
#' `xl_properties()` collects everything that applies at the *workbook* level:
#' document metadata, a few native workbook settings, and the formatting
#' defaults that used to be hard-coded.  The formatting defaults are ordinary
#' [xl_format] objects, so you can override them (e.g. change the header style
#' or the default date format) simply by passing a different `xl_format`.
#'
#' The `default_format` is cascaded *under* every cell (an emulated
#' workbook-wide default: libxlsxwriter has no native "Normal style" setter, so
#' it is merged beneath each cell/column format).  `header_format` styles the
#' header row and `hyperlink_format` styles cell hyperlinks; both are also
#' cascaded over `default_format`.
#'
#' @param title,subject,author,manager,company,category,keywords,comments,status,hyperlink_base
#'   Document metadata strings (Excel's "Properties" dialog).
#' @param custom A named list of custom document properties.  Values may be
#'   character, integer, numeric, or logical.
#' @param read_only Logical; mark the workbook read-only recommended.
#' @param window_size Optional integer vector `c(width, height)` for the
#'   workbook window size.
#' @param names A named list of workbook-scoped defined names, each a formula
#'   string (e.g. `list(tax = "=0.2")`).
#' @param default_format An [xl_format] cascaded under every cell (default:
#'   none).
#' @param header_format An [xl_format] for the header row (default: bold,
#'   centered).
#' @param hyperlink_format An [xl_format] for cell hyperlinks (default: blue,
#'   underlined).
#' @param date_format,datetime_format [xl_format] number formats applied to
#'   `Date` / `POSIXct` values.
#' @param date_col_width,datetime_col_width Default column width for `Date` /
#'   `POSIXct` columns.
#' @param header_row_height Height (in points) of the header row.
#' @return An `xl_properties` object.
#' @section Time zones:
#' Excel has no concept of a time zone.  When every `POSIXct` in the workbook
#' shares one time zone, writexl drops the zone and writes local wall-clock
#' time, and the default `datetime_format` loses its `" UTC"` suffix so that
#' nothing is mislabelled.  When the time zones differ, all datetimes are
#' converted to UTC with a warning.  Supplying your own `datetime_format`
#' overrides the label in either case.
#' @family writexl
#' @seealso [xl_workbook], [write_xlsx]
#' @export
#' @examples
#' xl_properties(title = "Quarterly report", author = "Finance",
#'               header_format = xl_font(bold = TRUE, color = "white") +
#'                               xl_fill(background = "navy"))
xl_properties <- function(title = NA, subject = NA, author = NA, manager = NA,
                          company = NA, category = NA, keywords = NA,
                          comments = NA, status = NA, hyperlink_base = NA,
                          custom = NULL, read_only = FALSE, window_size = NULL,
                          names = NULL,
                          default_format   = xl_format(),
                          header_format    = xl_font(bold = TRUE) +
                                             xl_align(horizontal = "center"),
                          hyperlink_format = xl_font(color = "blue",
                                                     underline = "single"),
                          date_format      = xl_num_format("yyyy-mm-dd"),
                          datetime_format  = .default_datetime_format(),
                          date_col_width = 20, datetime_col_width = 20,
                          header_row_height = 15) {
  fmts <- list(default_format = default_format, header_format = header_format,
               hyperlink_format = hyperlink_format, date_format = date_format,
               datetime_format = datetime_format)
  for (nm in names(fmts))
    if (!is_xl_format(fmts[[nm]]))
      stop(sprintf("`%s` must be an xl_format object", nm), call. = FALSE)
  for (nm in c("date_col_width", "datetime_col_width", "header_row_height")) {
    v <- get(nm)
    if (!is.numeric(v) || length(v) != 1L || is.na(v) || v < 0)
      stop(sprintf("`%s` must be a single non-negative number", nm), call. = FALSE)
  }
  if (!is.null(custom) && (is.null(base::names(custom)) || any(base::names(custom) == "")))
    stop("`custom` must be a named list", call. = FALSE)
  if (!is.null(window_size) && length(window_size) != 2L)
    stop("`window_size` must be a length-2 vector c(width, height)", call. = FALSE)
  if (!is.null(names) && (is.null(base::names(names)) || any(base::names(names) == "")))
    stop("`names` must be a named list of defined-name formulas", call. = FALSE)
  structure(
    list(title = title, subject = subject, author = author, manager = manager,
         company = company, category = category, keywords = keywords,
         comments = comments, status = status, hyperlink_base = hyperlink_base,
         custom = custom, read_only = read_only, window_size = window_size,
         names = names,
         default_format = default_format, header_format = header_format,
         hyperlink_format = hyperlink_format, date_format = date_format,
         datetime_format = datetime_format, date_col_width = date_col_width,
         datetime_col_width = datetime_col_width,
         header_row_height = header_row_height),
    class = "xl_properties"
  )
}

#' @export
print.xl_properties <- function(x, ...) {
  cat("<xl_properties>\n")
  meta <- c("title", "author", "company", "subject")
  set <- meta[vapply(meta, function(k) !is.null(x[[k]]) && !is.na(x[[k]]), logical(1))]
  if (length(set))
    for (k in set) cat(sprintf("  %s: %s\n", k, x[[k]]))
  if (isTRUE(x$read_only)) cat("  read_only: TRUE\n")
  invisible(x)
}

#' A workbook: sheets plus workbook-level properties
#'
#' @description
#' `xl_workbook()` binds one or more sheets (data frames or [xl_sheet]s) to a
#' set of [xl_properties].  It is the single place to attach workbook-level
#' formatting defaults and metadata.  When passed to [write_xlsx()], the
#' workbook's `col_names` and `format_headers` settings take precedence over
#' `write_xlsx()`'s own arguments.
#'
#' @param sheets A data frame, an [xl_sheet], or a (named) list of them.
#' @param properties An [xl_properties] object.
#' @param col_names Write column names as the header row?
#' @param format_headers Apply the header format to the header row?
#' @return An `xl_workbook` object.
#' @family writexl
#' @seealso [xl_properties], [xl_sheet], [write_xlsx]
#' @export
#' @examples
#' wb <- xl_workbook(
#'   list(Data = data.frame(x = 1:3)),
#'   properties = xl_properties(title = "Demo", author = "me")
#' )
#' tmp <- write_xlsx(wb)
xl_workbook <- function(sheets, properties = xl_properties(),
                        col_names = TRUE, format_headers = TRUE) {
  if (is.data.frame(sheets) || inherits(sheets, "xl_sheet"))
    sheets <- list(sheets)
  ok <- is.list(sheets) &&
    all(vapply(sheets, function(el) is.data.frame(el) || inherits(el, "xl_sheet"),
               logical(1)))
  if (!ok)
    stop("`sheets` must be a data frame, an xl_sheet, or a list of them",
         call. = FALSE)
  if (!inherits(properties, "xl_properties"))
    stop("`properties` must be an xl_properties object", call. = FALSE)
  structure(
    list(sheets = sheets, properties = properties,
         col_names = col_names, format_headers = format_headers),
    class = "xl_workbook"
  )
}

#' @export
print.xl_workbook <- function(x, ...) {
  cat(sprintf("<xl_workbook: %d sheet%s>\n", length(x$sheets),
              if (length(x$sheets) == 1L) "" else "s"))
  invisible(x)
}

# Build the C-side document-properties payload from an xl_properties object.
.properties_payload <- function(props) {
  meta_keys <- c("title", "subject", "author", "manager", "company",
                 "category", "keywords", "comments", "status", "hyperlink_base")
  out <- list()
  for (k in meta_keys) {
    v <- props[[k]]
    if (!is.null(v) && !(length(v) == 1L && is.na(v))) out[[k]] <- as.character(v)
  }
  if (!is.null(props$custom) && length(props$custom)) {
    out$custom <- lapply(base::names(props$custom), function(nm) {
      v <- props$custom[[nm]]
      type <- if (is.logical(v)) "boolean"
              else if (is.integer(v)) "integer"
              else if (is.numeric(v)) "number"
              else "string"
      list(name = nm, type = type,
           value = if (type == "string") as.character(v) else v)
    })
  }
  out$read_only <- as.integer(isTRUE(props$read_only))
  out$header_row_height <- as.numeric(props$header_row_height)
  if (!is.null(props$window_size))
    out$window_size <- as.integer(props$window_size)
  if (!is.null(props$names) && length(props$names))
    out$names <- lapply(base::names(props$names),
                        function(nm) list(name = nm,
                                          formula = as.character(props$names[[nm]])))
  out
}

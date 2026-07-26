#' Export to xlsx
#'
#' Writes a data frame to an xlsx file. To create an xlsx with (multiple) named
#' sheets, simply set \code{x} to a named list of data frames.
#'
#' Supports strings, numbers, booleans and dates automatically. For cell
#' formatting (fonts, fills, borders, number formats, ...), worksheet layout
#' (column widths, frozen panes, ...), and workbook metadata, wrap columns with
#' \code{\link{xl_cell_general}}, sheets with \code{\link{xl_sheet}}, and the
#' whole workbook with \code{\link{xl_workbook}}. See the "Formatting and
#' workbook properties" vignette and \code{\link{xl_format}}.
#'
#' \if{html}{
#' \out{
#' <link rel="stylesheet" type="text/css" href="https://jeroen.github.io/clippy/clippy.min.css" media="all">
#' <script src="https://jeroen.github.io/clippy/bundle.js"></script>
#' }}
#'
#' @export
#' @aliases writexl
#' @useDynLib writexl C_write_data_frame_list
#' @param x a data frame, an [xl_sheet], an [xl_workbook], or a (named) list of
#'   data frames / `xl_sheet`s that become the sheets in the xlsx
#' @param path a file name to write to
#' @param col_names write column names at the top of the file?
#' @param format_headers make the \code{col_names} in the xlsx centered and bold
#' @param use_zip64 use \href{https://en.wikipedia.org/wiki/Zip_(file_format)#ZIP64}{zip64}
#' to enable support for 4GB+ xlsx files. Not all platforms can read this.
#' @examples # Roundtrip example with single excel sheet named 'mysheet'
#' tmp <- write_xlsx(list(mysheet = iris))
#' readxl::read_xlsx(tmp)
write_xlsx <- function(x, path = tempfile(fileext = ".xlsx"), col_names = TRUE,
                       format_headers = TRUE, use_zip64 = FALSE){
  # Resolve the input to an xl_workbook.  A bare data frame / xl_sheet / list
  # is wrapped in a workbook with default properties; an explicit xl_workbook
  # overrides col_names/format_headers.
  if(inherits(x, "xl_workbook")){
    wb <- x
    col_names <- wb$col_names
    format_headers <- wb$format_headers
  } else {
    wb <- xl_workbook(x, properties = xl_properties(),
                      col_names = col_names, format_headers = format_headers)
  }
  props <- wb$properties
  elems <- wb$sheets
  dfs <- lapply(elems, function(el) if(inherits(el, "xl_sheet")) el$data else el)
  dfs <- lapply(dfs, normalize_df)
  names(dfs) <- names(elems)
  if(any(nchar(names(dfs)) > 31)){
    warning("Truncating sheet name(s) to 31 characters")
    names(dfs) <- substring(names(dfs), 1, 29)
  }
  nm <- names(dfs)
  if(length(unique(nm)) <  length(nm)){
    warning("Deduplicating sheet names")
    names(dfs) <- make.unique(substring(names(dfs), 1, 28), sep = "_")
  }
  stopifnot(is.character(path) && length(path))
  path <- normalizePath(path, mustWork = FALSE)
  # Excel has no concept of a time zone, so decide once, for the whole workbook,
  # how POSIXct values are written (see .resolve_timezones()).
  tz_res <- .resolve_timezones(dfs, props)
  dfs <- tz_res$dfs
  props <- tz_res$props
  # Resolve everything into one deduplicated workbook format table: the header
  # format, each general cell's effective format, and each sheet's column/row
  # plan -- all cascaded over the workbook default_format.
  header_offset <- if(isTRUE(as.logical(col_names))) 1L else 0L
  reg <- .new_format_registry()
  header_id <- .register_format(reg, merge_xl_format(props$default_format,
                                                     props$header_format))
  dfs <- lapply(dfs, .resolve_sheet_formats, reg = reg, props = props)
  sheets <- Map(function(el, df) .resolve_sheet_plan(el, df, reg, header_offset, props),
                elems, dfs)
  ret <- .Call(C_write_data_frame_list, dfs, path, col_names, format_headers,
               use_zip64, reg$table, sheets, header_id, .properties_payload(props))
  invisible(ret)
}

# =============================================================================
# Time zones
# =============================================================================
#
# Excel stores a datetime as a naive serial number: it has no time zone.  Rather
# than silently converting every POSIXct to UTC, writexl decides per workbook:
#
#   * all POSIXct values share one time zone -> drop the zone and write local
#     wall-clock time (2025-12-01 01:00 in Perth is written as 01:00), and drop
#     the " UTC" suffix from the default datetime format so nothing is mislabelled;
#   * the time zones differ -> convert everything to UTC and warn, since no
#     single wall clock can represent them all.
#
# The survey covers plain POSIXct columns and POSIXct values inside
# xl_cell_general cells.

# The time zone of a POSIXct, with NULL/"" (meaning local time) resolved to the
# session zone so that an unset zone and an explicit local zone compare equal.
.tzone_of <- function(x){
  tz <- attr(x, "tzone", exact = TRUE)
  if(is.null(tz) || !length(tz) || !nzchar(tz[1])){
    loc <- Sys.timezone()
    if(is.na(loc)) "" else loc
  } else {
    tz[1]
  }
}

# Every distinct time zone used by a POSIXct anywhere in the workbook.
.collect_tzones <- function(dfs){
  out <- character(0)
  for(df in dfs){
    for(col in df){
      if(inherits(col, "POSIXct")){
        out <- c(out, .tzone_of(col))
      } else if(inherits(col, "xl_cell_general")){
        for(rec in unclass(col))
          if(inherits(rec$value, "POSIXct"))
            out <- c(out, .tzone_of(rec$value))
      }
    }
  }
  unique(out)
}

# Re-express a POSIXct so that its wall-clock reading is what reaches the cell.
# The offset is taken per instant, so daylight saving is handled correctly.
.drop_tzone <- function(x){
  off <- as.POSIXlt(x, tz = .tzone_of(x))$gmtoff
  # Some platforms do not report gmtoff; leave the value alone (i.e. UTC) rather
  # than writing something wrong.
  if(is.null(off) || any(is.na(off) & !is.na(as.numeric(x))))
    return(x)
  off[is.na(off)] <- 0
  structure(as.numeric(x) + off, class = c("POSIXct", "POSIXt"), tzone = "UTC")
}

# Apply .drop_tzone() to every POSIXct in a sheet, including inside cell objects.
.drop_tzones_sheet <- function(df){
  for(j in seq_along(df)){
    col <- df[[j]]
    if(inherits(col, "POSIXct")){
      df[[j]] <- .drop_tzone(col)
    } else if(inherits(col, "xl_cell_general")){
      recs <- unclass(col)
      touched <- FALSE
      for(k in seq_along(recs)){
        if(inherits(recs[[k]]$value, "POSIXct")){
          recs[[k]]$value <- .drop_tzone(recs[[k]]$value)
          touched <- TRUE
        }
      }
      if(touched)
        df[[j]] <- structure(recs, class = class(col))
    }
  }
  df
}

.resolve_timezones <- function(dfs, props){
  tzs <- .collect_tzones(dfs)
  if(length(tzs) > 1L){
    warning("The workbook contains datetimes in ", length(tzs),
            " different time zones (", paste(tzs, collapse = ", "),
            "); all of them were converted to UTC because Excel does not ",
            "support time zones. Convert them yourself beforehand if you want ",
            "different behaviour.", call. = FALSE)
  } else if(length(tzs) == 1L){
    dfs <- lapply(dfs, .drop_tzones_sheet)
    # The default format labels datetimes "UTC"; that is no longer true once the
    # zone has been dropped.  Only replace it if the caller kept the default.
    if(identical(unclass(props$datetime_format),
                 unclass(.default_datetime_format())))
      props$datetime_format <- xl_num_format("yyyy-mm-dd HH:mm:ss")
  }
  list(dfs = dfs, props = props)
}

# Walk a sheet's xl_cell_general columns, register each cell's effective format
# in the workbook registry, attach the resulting integer id vector, and flatten
# each comment's xl_format into the flat fields C applies.
.resolve_sheet_formats <- function(df, reg, props) {
  for (j in seq_along(df)) {
    col <- df[[j]]
    if (!inherits(col, "xl_cell_general")) next
    recs <- unclass(col)
    ids <- vapply(recs, .resolve_cell_format_id, integer(1), reg = reg, props = props)
    recs <- lapply(recs, function(rec) {
      if (!is.null(rec$comment)) rec$comment <- .comment_c_payload(rec$comment)
      rec
    })
    col <- structure(recs, class = class(col))
    attr(col, "writexl_format_ids") <- ids
    df[[j]] <- col
  }
  df
}

# Resolve one cell record's effective format to a registry id (0 = none),
# cascading workbook default_format, then the hyperlink/date default, then the
# cell's own format.
.resolve_cell_format_id <- function(rec, reg, props) {
  fmt <- rec$format
  val <- rec$value
  fm  <- rec$formula
  hl  <- rec$hyperlink
  formula_set   <- !(is.null(fm) || (length(fm) == 1L && is.na(fm)))
  hyperlink_set <- !(is.null(hl) || identical(hl, NA) ||
                     (is.character(hl) && length(hl) == 1L && is.na(hl)))
  base <- props$default_format
  if (hyperlink_set) {
    base <- merge_xl_format(base, props$hyperlink_format)
  } else if (!formula_set && !is.null(val) &&
             (inherits(val, "Date") || inherits(val, "POSIXct"))) {
    has_num <- is_xl_format(fmt) && !is.null(unclass(fmt)$num_format)
    if (!has_num)
      base <- merge_xl_format(base, if (inherits(val, "POSIXct"))
                                props$datetime_format else props$date_format)
  }
  .register_format(reg, merge_xl_format(base, fmt))
}

normalize_df <- function(df){
  if(nrow(df) > 1024^2){
    stop("the xlsx format does not support tables with 1M+ rows")
  }
  # Types to coerce to strings
  for(i in which(vapply(df, inherits, logical(1), c("factor", "hms")))){
    df[[i]] <- as.character(df[[i]])
  }
  for(i in which(vapply(df, function(x){is.integer(x) && inherits(x, "POSIXct")}, logical(1)))){
    df[[i]] <- as.POSIXct(as.double(df[[i]]))
  }
  for(i in which(vapply(df, inherits, logical(1), "POSIXlt"))){
    df[[i]] <- as.POSIXct(df[[i]])
  }
  for(i in which(vapply(df, inherits, logical(1), "integer64"))){
    warning(sprintf("Coercing column %s from int64 to double", names(df)[i]), call. = FALSE)
    getNamespace("bit64")
    df[[i]] <- as.double(df[[i]])
  }
  df
}

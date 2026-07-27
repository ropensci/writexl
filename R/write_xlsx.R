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
  names(dfs) <- .resolve_sheet_names(names(elems), length(dfs))
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
  dfs <- lapply(dfs, .resolve_sheet_formats, reg = reg, props = props,
                header_offset = header_offset)
  sheets <- Map(function(el, df) .resolve_sheet_plan(el, df, reg, header_offset, props),
                elems, dfs)
  cm <- .resolve_constant_memory(dfs, props)
  ret <- .Call(C_write_data_frame_list, dfs, path, col_names, format_headers,
               use_zip64, reg$table, sheets, header_id,
               .properties_payload(props, cm$on))
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
  # Not every platform reports gmtoff: older Windows builds of R return NULL or
  # a zero-length vector.  Fall back to leaving the value alone (i.e. UTC)
  # rather than writing something wrong -- note the length check, without which
  # a zero-length offset would recycle to nothing and silently empty the column.
  if(is.null(off) || length(off) != length(x) ||
     any(is.na(off) & !is.na(as.numeric(x))))
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
# in the workbook registry, attach the resulting integer id vector, flatten each
# comment's xl_format into the flat fields C applies, and resolve each array
# formula's range to the 0-based quad C writes it with.
#
# Multi-cell array ranges are collected on the returned data frame (attribute
# "writexl_array_multicell") because they decide the workbook's constant-memory
# flag -- see .resolve_constant_memory().
.resolve_sheet_formats <- function(df, reg, props, header_offset = 1L) {
  multicell <- list()
  for (j in seq_along(df)) {
    col <- df[[j]]
    if (!inherits(col, "xl_cell_general")) next
    recs <- unclass(col)
    ids <- vapply(recs, .resolve_cell_format_id, integer(1), reg = reg, props = props)
    for (k in seq_along(recs)) {
      rec <- recs[[k]]
      if (!is.null(rec$comment)) rec$comment <- .comment_c_payload(rec$comment)
      fm_set <- !is.null(rec$formula) && length(rec$formula) == 1L &&
                !is.na(rec$formula)
      if ((isTRUE(rec$array) || isTRUE(rec$dynamic)) && fm_set) {
        anchor <- c((k - 1L) + as.integer(header_offset), j - 1L)
        q <- .resolve_array_quad(rec$array_range, anchor, df, header_offset,
                                 names(df)[j], k)
        rec$array_quad <- q
        if (q[3L] > q[1L] || q[4L] > q[2L]) {
          .check_array_no_overlap(q, anchor, df, header_offset, names(df)[j], k)
          multicell <- c(multicell, list(q))
        }
      }
      # a rich string's runs each carry their own font, registered like any
      # other format; C receives the parallel text / id vectors
      if (is_xl_rich_string(rec$value)) {
        rich <- .rich_string_c_payload(rec$value, reg, props)
        rec$rich_text <- rich$rich_text
        rec$rich_format_id <- rich$rich_format_id
        rec$value <- NULL
      }
      # the unresolved user spec has served its purpose; keep the C payload tight
      rec$array_range <- NULL
      recs[[k]] <- rec
    }
    col <- structure(recs, class = class(col))
    attr(col, "writexl_format_ids") <- ids
    df[[j]] <- col
  }
  attr(df, "writexl_array_multicell") <- multicell
  df
}

# Resolve one array formula's range.  With no declared range the formula covers
# only its own cell, which is the modern dynamic-array spelling: Excel spills
# the result itself.  A declared range must start at the cell, because
# libxlsxwriter writes the formula into the range's top-left corner -- were the
# two allowed to differ the formula would land in a different cell than the one
# the caller wrote it to.
.resolve_array_quad <- function(spec, anchor, df, header_offset, colname, k) {
  if (is.null(spec))
    return(as.integer(c(anchor[1L], anchor[2L], anchor[1L], anchor[2L])))
  arg <- sprintf('array_range (column "%s", cell %d)', colname, k)
  q <- .xl_resolve_range(spec, arg = arg, df = df,
                         header_offset = header_offset, allow_cell = TRUE)
  if (q[1L] != anchor[1L] || q[2L] != anchor[2L])
    stop(sprintf(paste0("`%s` must start at the cell that holds the formula ",
                        "(row %d, column %d of the sheet, 1-based)"),
                 arg, anchor[1L] + 1L, anchor[2L] + 1L), call. = FALSE)
  q
}

# Refuse a multi-cell array range that overlaps cells the sheet writes itself.
#
# libxlsxwriter pads a multi-cell array range with formatted zeroes, which the
# row loop would then overwrite (or which would overwrite the sheet's own
# values, depending on ordering).  Rather than emit a file whose array range is
# half data, require the range to spill into cells the sheet does not occupy.
.check_array_no_overlap <- function(q, anchor, df, header_offset, colname, k) {
  nrows <- nrow(df)
  ncols <- length(df)
  if (nrows < 1L || ncols < 1L) return(invisible(NULL))
  # the sheet's own data rectangle, 0-based
  first_row <- as.integer(header_offset)
  last_row  <- first_row + nrows - 1L
  r1 <- max(q[1L], first_row); r2 <- min(q[3L], last_row)
  c1 <- max(q[2L], 0L);        c2 <- min(q[4L], ncols - 1L)
  if (r2 < r1 || c2 < c1) return(invisible(NULL))
  ncell <- (as.numeric(r2 - r1) + 1) * (as.numeric(c2 - c1) + 1)
  if (ncell == 1 && r1 == anchor[1L] && c1 == anchor[2L])
    return(invisible(NULL))
  stop(sprintf(paste0(
    'the `array_range` on column "%s", cell %d covers %.0f cell(s) that the ',
    "sheet writes itself. A multi-cell array range must extend into cells the ",
    "sheet does not occupy, because the range is padded on write and the two ",
    "would overwrite each other. Either drop `array_range` (a single-cell ",
    "dynamic formula spills automatically in modern Excel) or move the range ",
    "beyond the data."), colname, k, ncell), call. = FALSE)
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
  .check_supported_columns(df)
  df
}

# Whether a column can be written at all.  This mirrors get_type() in
# src/write_xlsx.c: anything backed by a character, integer, double or logical
# vector is writable (a classed column such as Date, POSIXct or difftime is
# written from its underlying values), and xl_cell_general carries its own
# writer.  Everything else -- complex, raw, a plain list column -- has no
# representation in xlsx.
.is_supported_column <- function(x){
  if(inherits(x, "xl_cell_general"))
    return(TRUE)
  is.character(x) || is.integer(x) || is.double(x) || is.logical(x)
}

# Fail loudly on column types writexl cannot represent, rather than silently
# writing the underlying values (or nothing at all).
.check_supported_columns <- function(df){
  bad <- which(!vapply(df, .is_supported_column, logical(1)))
  if(!length(bad))
    return(invisible(NULL))
  detail <- vapply(bad, function(i){
    cls <- paste(class(df[[i]]), collapse = "/")
    ty <- typeof(df[[i]])
    sprintf("'%s' (column %d): %s", names(df)[i], i,
            if(identical(cls, ty)) cls else sprintf("%s (%s)", cls, ty))
  }, character(1))
  stop("writexl cannot write column(s) of these types:\n  ",
       paste(detail, collapse = "\n  "), call. = FALSE)
}

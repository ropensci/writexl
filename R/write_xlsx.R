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

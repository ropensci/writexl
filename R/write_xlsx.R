#' Export to xlsx
#'
#' Writes a data frame to an xlsx file. To create an xlsx with (multiple) named
#' sheets, simply set \code{x} to a named list of data frames.
#'
#' Currently supports strings, numbers, booleans and dates. Formatting options
#' may be added in future versions.
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
#' @param x data frame or named list of data frames that will be sheets in the xlsx
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
  if(is.data.frame(x) || inherits(x, "xl_sheet"))
    x <- list(x)
  sheet_like <- is.list(x) &&
    all(vapply(x, function(el) is.data.frame(el) || inherits(el, "xl_sheet"),
               logical(1)))
  if(!sheet_like)
    stop("Argument x must be a data frame, an xl_sheet, or a list of them")
  # Underlying data frames (unwrap any xl_sheet wrappers), keeping the
  # originals so their worksheet options can be resolved below.
  elems <- x
  dfs <- lapply(x, function(el) if(inherits(el, "xl_sheet")) el$data else el)
  dfs <- lapply(dfs, normalize_df)
  names(dfs) <- names(x)
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
  # Resolve per-cell formats for xl_cell_general columns and per-sheet
  # column/row/layout options into a single deduplicated workbook format
  # table, attaching integer ids to each column and building a sheet plan.
  header_offset <- if(isTRUE(as.logical(col_names))) 1L else 0L
  reg <- .new_format_registry()
  dfs <- lapply(dfs, .resolve_sheet_formats, reg = reg)
  sheets <- Map(function(el, df) .resolve_sheet_plan(el, df, reg, header_offset),
                elems, dfs)
  ret <- .Call(C_write_data_frame_list, dfs, path, col_names, format_headers,
               use_zip64, reg$table, sheets)
  invisible(ret)
}

# Default number formats applied to date/time values written through
# xl_cell_general when the cell's own format sets no number format.  These
# mirror the column-level defaults used for plain Date/POSIXct columns.
.writexl_default_date_format     <- function() xl_num_format("yyyy-mm-dd")
.writexl_default_datetime_format <- function() xl_num_format("yyyy-mm-dd HH:mm:ss UTC")

# Walk a sheet's xl_cell_general columns, register each cell's effective format
# in the workbook registry, and attach the resulting integer id vector.
.resolve_sheet_formats <- function(df, reg) {
  for (j in seq_along(df)) {
    col <- df[[j]]
    if (!inherits(col, "xl_cell_general")) next
    recs <- unclass(col)
    ids <- vapply(recs, .resolve_cell_format_id, integer(1), reg = reg)
    attr(col, "writexl_format_ids") <- ids
    df[[j]] <- col
  }
  df
}

# Resolve one cell record's effective format to a registry id (0 = none).
.resolve_cell_format_id <- function(rec, reg) {
  fmt <- rec$format
  val <- rec$value
  fm  <- rec$formula
  hl  <- rec$hyperlink
  formula_set   <- !(is.null(fm) || (length(fm) == 1L && is.na(fm)))
  hyperlink_set <- !(is.null(hl) || identical(hl, NA) ||
                     (is.character(hl) && length(hl) == 1L && is.na(hl)))
  if (!formula_set && !hyperlink_set && !is.null(val) &&
      (inherits(val, "Date") || inherits(val, "POSIXct"))) {
    has_num <- is_xl_format(fmt) && !is.null(unclass(fmt)$num_format)
    if (!has_num) {
      base <- if (inherits(val, "POSIXct")) .writexl_default_datetime_format()
              else .writexl_default_date_format()
      fmt <- if (is.null(fmt)) base else merge_xl_format(base, fmt)
    }
  }
  .register_format(reg, fmt)
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

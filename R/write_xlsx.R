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
#' @param format_headers make the \code{col_names} in the xlsx centred, bold
#' and apply the header colour.
#' @param use_zip64 use \href{https://en.wikipedia.org/wiki/Zip_(file_format)#ZIP64}{zip64}
#' to enable support for 4GB+ xlsx files. Not all platforms can read this.
#' @param freeze_rows number of rows to freeze at the top of the sheet
#' (default: 0 i.e. no freeze)
#' @param freeze_cols number of columns to freeze at the left of the sheet
#' (default: 0 i.e. no freeze)
#' @param autofilter add auto filter to columns (default: FALSE)
#' @param header_bg_color background colour for header row cells (default: NA
#'  meaning no color) other values could include any R colours like
#'   "lightgray", "green", "lightblue" etc.
#' @examples # Roundtrip example with single excel sheet named 'mysheet'
#' tmp <- write_xlsx(list(mysheet = iris))
#' readxl::read_xlsx(tmp)
write_xlsx <- function(x, path = tempfile(fileext = ".xlsx"), col_names = TRUE,
                       format_headers = TRUE, use_zip64 = FALSE,
                       freeze_rows = 0L,  freeze_cols = 0L,
                       autofilter=FALSE, header_bg_color=NA){
  if(is.data.frame(x))
    x <- list(x)
  if(!is.list(x) || !all(vapply(x, is.data.frame, logical(1))))
    stop("Argument x must be a data frame or list of data frames")
  x <- lapply(x, normalize_df)
  if(any(nchar(names(x)) > 31)){
    warning("Truncating sheet name(s) to 31 characters")
    names(x) <- substring(names(x), 1, 29)
  }
  nm <- names(x)
  if(length(unique(nm)) <  length(nm)){
    warning("Deduplicating sheet names")
    names(x) <- make.unique(substring(names(x), 1, 28), sep = "_")
  }
  if(!is.numeric(freeze_rows)){
    stop("freeze_rows must be numeric! Got: ", freeze_rows)
  }
  freeze_rows <- as.integer(freeze_rows)

  if(!is.numeric(freeze_cols)){
    stop("freeze_cols must be numeric! Got: ", freeze_cols)
  }
  freeze_cols <- as.integer(freeze_cols)

  hex_color <- function(color){
    if(is.na(color)){return(-1L)}
    rgb_values <- col2rgb(color)
    hex_color <- rgb(rgb_values[1], rgb_values[2], rgb_values[3],
                     maxColorValue = 255)
    as.integer((paste0("0x", toupper(sub("#", "", hex_color)))))
  }

  if(!is.logical(autofilter)){stop("autofilter expects a logical value!")}

  if(!is.numeric(col_widths)){stop("col_widths expects numeric values!")}

  stopifnot(is.character(path) && length(path))
  path <- normalizePath(path, mustWork = FALSE)
  ret <- .Call(C_write_data_frame_list, x, path, col_names, format_headers,
               use_zip64, freeze_rows, freeze_cols, hex_color(header_bg_color),
               autofilter)
  invisible(ret)
}

normalize_df <- function(df){
  # Types to coerce to strings
  for(i in which(vapply(df, inherits, logical(1), c("factor", "hms")))){
    df[[i]] <- as.character(df[[i]])
  }
  for(i in which(vapply(df, inherits, logical(1), "POSIXlt"))){
    df[[i]] <- as.POSIXct(df[[i]])
  }
  for(i in which(vapply(df, inherits, logical(1), "integer64"))){
    warning(sprintf("Coercing columnn %s from int64 to double", names(df)[i]), call. = FALSE)
    df[[i]] <- bit64::as.double.integer64(df[[i]])
  }
  df
}

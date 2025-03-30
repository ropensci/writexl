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
  stopifnot(is.character(path) && length(path))
  path <- normalizePath(path, mustWork = FALSE)
  ret <- .Call(C_write_data_frame_list, x, path, col_names, format_headers, use_zip64)
  invisible(ret)
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
  # Find any unsupported column classes
  class_unsupported_idx <-
    which(
      !vapply(
        df,
        FUN = inherits,
        FUN.VALUE = logical(1),
        what = c("character", "numeric", "logical", "POSIXct", "Date")
      )
    )
  if (length(class_unsupported_idx) > 0) {
    msg <- character()
    for (col_idx in class_unsupported_idx) {
      col_name <- names(df)[col_idx]
      col_info <- sprintf("column name '%s' (column number %g)", col_name, col_idx)
      msg <-
        c(
          msg,
          sprintf(
            "%s; class: %s", col_info,
            paste(class(df[[col_idx]]), collapse = ", ")
          )
      )
    }
    stop(
      "Unsupported class for the following ",
      ngettext(length(msg), "column:\n", "columns:\n"),
      paste(msg, sep = "\n")
    )
  }
  df
}

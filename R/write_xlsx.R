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
#' @examples # Roundtrip example
#' tmp <- write_xlsx(iris)
#' readxl::read_xlsx(tmp)
write_xlsx <- function(x, path = tempfile(fileext = ".xlsx"), col_names = TRUE){
  if(is.data.frame(x))
    x <- list(x)
  if(!is.list(x) || !all(vapply(x, is.data.frame, logical(1))))
    stop("Argument x must be a data frame or list of data frames")
  x <- lapply(x, normalize_df)
  stopifnot(is.character(path) && length(path))
  path <- normalizePath(path, mustWork = FALSE)
  ret <- .Call(C_write_data_frame_list, x, path, col_names)
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

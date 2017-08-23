#' Export to xlsx
#'
#' Writes a data frame to an xlsx file.
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
#' @useDynLib writexl C_write_data_frame
#' @param x data frame to write to disk
#' @param path a file name to write to
#' @param col_names write column names at the top of the file?
#' @examples # Roundtrip example
#' tmp <- write_xlsx(iris)
#' readxl::read_xlsx(tmp)
write_xlsx <- function(x, path = tempfile(fileext = ".xlsx"), col_names = TRUE){
  stopifnot(is.data.frame(x))
  stopifnot(is.character(path) && length(path))
  path <- normalizePath(path, mustWork = FALSE)
  df <- normalize_df(x)
  headers <- if(isTRUE(col_names))
    colnames(x)
  .Call(C_write_data_frame, df, path, headers)
}

normalize_df <- function(df){
  for(i in which(vapply(df, is.factor, logical(1)))){
    df[[i]] <- as.character(df[[i]])
  }
  for(i in which(vapply(df, inherits, logical(1), "POSIXlt"))){
    df[[i]] <- as.POSIXct(df[[i]])
  }
  df
}

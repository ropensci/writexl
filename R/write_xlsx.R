#' Export to xlsx
#'
#' Writes a data frame to an xlsx file.
#'
#' @export
#' @aliases writexl
#' @useDynLib writexl C_write_data_frame
#' @param x data frame to write to disk
#' @param path a file name to write to
#' @param col_names write column names at the top of the file?
#' @param date_format how timestamps are displayed in excel
#' @examples tmp <- write_xlsx(iris)
#' readxl::read_xlsx(tmp)
write_xlsx <- function(x, path = tempfile(fileext = ".xlsx"), col_names = TRUE, date_format = "yyyy-mm-dd HH:mm:ss UTC"){
  stopifnot(is.data.frame(x))
  stopifnot(is.character(path) && length(path))
  stopifnot(is.character(date_format) && length(date_format))
  path <- normalizePath(path, mustWork = FALSE)
  df <- normalize_df(x)
  headers <- if(isTRUE(col_names))
    colnames(x)
  .Call(C_write_data_frame, df, path, headers, date_format)
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

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
#' @param col_widths A list of numeric vectors with column widths to use
#'  (in inches). A length one vector is recycled. NA gives the default
#'  behaviour. You can have the function guess the column width from the number
#'  of characters in a row for this use one string value per data frame.
#'  The possible string values include:
#'   \itemize{
#'     \item "guess_from_header": The widths are guessed from the character
#'      count of the header.
#'     \item "guess_from_row_2": The widths are guessed from the character
#'      count of the 2nd row. You can use any row number here.
#'   }
#'  @param guessed_column_width_padding numeric inches to add to guessed
#'   column widths (default:2)
#' @examples # Roundtrip example with single excel sheet named 'mysheet'
#' tmp <- write_xlsx(list(mysheet = iris))
#' readxl::read_xlsx(tmp)
write_xlsx <- function(x, path = tempfile(fileext = ".xlsx"), col_names = TRUE,
                       format_headers = TRUE, use_zip64 = FALSE,
                       freeze_rows = 0L,  freeze_cols = 0L,
                       autofilter=FALSE, header_bg_color=NA,
                       col_widths = NA,
                       guessed_column_width_padding = 2){

  if(!is.list(col_widths)){col_widths <- list(col_widths)}

  if(is.data.frame(x)){ x <- list(x)}

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

  if(length(x) != length(col_widths)){
    #Recycling col_widths
    if(length(col_widths) == 1){
      col_widths <- lapply(1:length(x),
                           function(x, cw){col_widths[[1]]}, col_widths)
    } else{
      stop("The col_widths don't have the same number of width vectors as the ",
           "input data frames")
    }
  }
  #convert the column widths to numeric and guess if needed
  col_widths <- lapply(1:length(x), guess_col_widths,
                       cw=col_widths,
                       dfs = x,
                       extra_space = guessed_column_width_padding)


  stopifnot(is.character(path) && length(path))
  path <- normalizePath(path, mustWork = FALSE)
  ret <- .Call(C_write_data_frame_list, x, path, col_names, format_headers,
               use_zip64, freeze_rows, freeze_cols, hex_color(header_bg_color),
               autofilter, col_widths)
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
  for(i in which(vapply(df, inherits, logical(1), "POSIXlt"))){
    df[[i]] <- as.POSIXct(df[[i]])
  }
  for(i in which(vapply(df, inherits, logical(1), "integer64"))){
    warning(sprintf("Coercing columnn %s from int64 to double", names(df)[i]), call. = FALSE)
    df[[i]] <- bit64::as.double.integer64(df[[i]])
  }
  df
}


guess_col_widths <- function(i, cws, dfs, extra_space = 2){
  cw <- cws[[i]]

  if(is.numeric(cw)){return(cw)}
  if(length(cw) != 1){stop("If not numeric, col_widths should have lenght 1!")}
  if(is.na(cw)){return(NA)}
  guess_names <- NULL
  if(grepl("guess_from_head", cw)){
    guess_names <- colnames(dfs[[i]])
  }
  if(grepl("guess_from_row",cw)){
    row_num <- as.numeric(gsub("[^0-9]", "", cw))
    if(nrow(dfs[[i]]) < row_num){
      stop("Rows in data frame: ", nrow(dfs[[i]]),
           "\nGuess row asked: ", row_num)
      }
    guess_names <- dfs[[i]][row_num, ]
  }
  if(is.null(guess_names)){stop("Unexpected col_widths string: ", cw)}

  nchar(as.character(guess_names)) + extra_space
}

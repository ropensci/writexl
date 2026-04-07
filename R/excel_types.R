#' Excel Types
#'
#' Create special column types to write to a spreadsheet
#'
#' @family writexl
#' @param x character vector to be interpreted as formula
#' @export
#' @rdname xl_formula
#' @examples
#' df <- data.frame(
#'   name = c("UCLA", "Berkeley", "Jeroen"),
#'   founded = c(1919, 1868, 2030),
#'   website = xl_hyperlink(c("http://www.ucla.edu", "http://www.berkeley.edu", NA), "homepage")
#' )
#' df$age <- xl_formula('=(YEAR(TODAY()) - INDIRECT("B" & ROW()))')
#' write_xlsx(df, 'universities.xlsx')
#'
#' # cleanup
#' unlink('universities.xlsx')
xl_formula <- function(x){
  if(is.factor(x))
    x <- as.character(x)
  stopifnot(is.character(x))
  if(!all(grepl("^=", x) | is.na(x)))
    stop("Formulas must start with '='")
  xl_cell_general(formula = x)
}

#' @rdname xl_formula
#' @export
#' @param url character vector of URLs
#' @param name character vector of friendly names
xl_hyperlink <- function(url, name = NULL){
  if(is.factor(url))
    url <- as.character(url)
  stopifnot(is.character(url))
  if(is.null(name)){
    xl_cell_general(hyperlink = url)
  } else {
    hlinks <- mapply(
      function(u, n) if(is.na(u)) NA else list(url = u, string = n),
      url, name,
      SIMPLIFY = FALSE
    )
    xl_cell_general(hyperlink = hlinks)
  }
}

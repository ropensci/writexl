#' Version
#'
#' Shows version of bundled libxlsxwriter.
#'
#' @export
#' @rdname writexl
#' @useDynLib writexl C_lxw_version
lxw_version <- function(){
  version <- .Call(C_lxw_version)
  as.numeric_version(version)
}

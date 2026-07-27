# =============================================================================
# Rich strings: one cell, several differently formatted runs
# =============================================================================
#
# A rich string is a single cell whose text is split into runs, each with its
# own font.  libxlsxwriter renders a run's format with lxw_styles_write_rich_font(),
# which emits only the <font> element -- so a run carries font properties and
# nothing else.  As with xl_comment(), the styling input is an ordinary
# xl_format and the unsupported groups are warned about rather than silently
# dropped (see .check_rich_string_format()).
#
# Excel itself imposes the shape: a run must have text, and a "rich" string
# needs at least two runs.  A single-format string is a plain character value.
# -----------------------------------------------------------------------------

#' One run of a rich (multi-format) string
#'
#' @description
#' `xl_rich_run()` is one fragment of an [xl_rich_string()]: a piece of text
#' plus the font it is drawn in.
#'
#' @param text A single non-`NA`, non-empty string.
#' @param format An optional [xl_format].  Only its **font** properties apply
#'   (see [xl_font()]); a run has no fill, border, alignment or number format,
#'   and supplying one warns.  `NULL` draws the run in the cell's own font.
#' @return An `xl_rich_run` object.
#' @family writexl
#' @seealso [xl_rich_string], [xl_format]
#' @export
#' @examples
#' xl_rich_run("bold", xl_font(bold = TRUE))
#' xl_rich_run("plain")
xl_rich_run <- function(text, format = NULL) {
  if (!is.character(text) || length(text) != 1L || is.na(text))
    stop("`text` must be a single non-NA string", call. = FALSE)
  if (!nzchar(text))
    stop("`text` must not be empty: Excel has no representation for an empty ",
         "run of a rich string", call. = FALSE)
  if (!is.null(format)) {
    if (!is_xl_format(format))
      stop("`format` must be an xl_format object", call. = FALSE)
    .check_rich_string_format(format)
  }
  structure(list(text = text, format = format), class = "xl_rich_run")
}

#' A cell whose text has several formats
#'
#' @description
#' `xl_rich_string()` builds the value of a single cell out of differently
#' formatted runs, so that one cell can read "This is **bold** text".  Pass it
#' as the `value` of [xl_cell_general()].
#'
#' Excel requires at least two runs: a string with one format is an ordinary
#' character value, so pass it as one.
#'
#' @param ... Runs, in order.  A bare string is taken as an unformatted run; an
#'   [xl_rich_run()] carries its own font.  Lists of either are flattened, so
#'   runs can be assembled programmatically.
#' @return An `xl_rich_string` object: a list of runs.
#' @family writexl
#' @seealso [xl_rich_run], [xl_cell_general], [xl_font]
#' @export
#' @examples
#' xl_rich_string("This is ", xl_rich_run("bold", xl_font(bold = TRUE)), " text")
#'
#' # in a cell, with a cell-wide format alongside the per-run fonts
#' xl_cell_general(
#'   value  = xl_rich_string("2 H", xl_rich_run("2", xl_font(script = "sub")), "O"),
#'   format = xl_align(horizontal = "center")
#' )
xl_rich_string <- function(...) {
  runs <- .flatten_rich_runs(list(...))
  if (length(runs) < 2L)
    stop("a rich string needs at least 2 runs (got ", length(runs),
         "); a string with a single format is an ordinary character value",
         call. = FALSE)
  structure(runs, class = "xl_rich_string")
}

#' @rdname xl_rich_string
#' @param x An object to test.
#' @export
is_xl_rich_string <- function(x) inherits(x, "xl_rich_string")

# Turn the `...` of xl_rich_string() into a flat list of xl_rich_run objects.
# Bare strings become unformatted runs; nested lists are flattened so runs may
# be built up with lapply() and spliced in.
.flatten_rich_runs <- function(x) {
  out <- list()
  for (el in x) {
    if (inherits(el, "xl_rich_run")) {
      out <- c(out, list(el))
    } else if (is.character(el)) {
      if (!length(el))
        stop("a character run must contain at least one string", call. = FALSE)
      out <- c(out, lapply(el, xl_rich_run))
    } else if (is.list(el)) {
      out <- c(out, .flatten_rich_runs(el))
    } else {
      stop("each run must be a string or an xl_rich_run object, not ",
           paste(class(el), collapse = "/"), call. = FALSE)
    }
  }
  out
}

# Warn about xl_format properties a rich-string run cannot render.  A run is
# drawn by lxw_styles_write_rich_font(), which emits a <font> element only, so
# the whole font group applies and no other group does.
.check_rich_string_format <- function(format) {
  f <- unclass(format)
  bad <- character(0)
  for (g in c("fill", "border", "align", "num_format", "protection"))
    if (!is.null(f[[g]])) bad <- c(bad, g)
  if (isTRUE(f$quote_prefix)) bad <- c(bad, "quote_prefix")
  if (length(bad))
    warning("xl_rich_run(): a rich string run is drawn with a font only, so ",
            "these format properties will be ignored: ",
            paste(bad, collapse = ", "), call. = FALSE)
  invisible(bad)
}

#' @export
print.xl_rich_run <- function(x, ...) {
  cat(sprintf("<xl_rich_run: %s%s>\n", encodeString(x$text, quote = '"'),
              if (is.null(x$format)) "" else " (formatted)"))
  invisible(x)
}

#' @export
print.xl_rich_string <- function(x, ...) {
  runs <- unclass(x)
  cat(sprintf("<xl_rich_string: %d runs>\n", length(runs)))
  for (r in runs)
    cat(sprintf("  %s%s\n", encodeString(r$text, quote = '"'),
                if (is.null(r$format)) "" else "  <formatted>"))
  invisible(x)
}

#' @export
format.xl_rich_string <- function(x, ...) {
  paste(vapply(unclass(x), function(r) r$text, character(1)), collapse = "")
}

# Flatten a rich string to the parallel text / format-id vectors C consumes,
# registering each run's format in the workbook registry.  Run formats cascade
# over the workbook default_format like every other format path.
.rich_string_c_payload <- function(rs, reg, props) {
  runs <- unclass(rs)
  list(
    rich_text = vapply(runs, function(r) r$text, character(1)),
    rich_format_id = vapply(runs, function(r)
      .register_format(reg, merge_xl_format(props$default_format, r$format)),
      integer(1))
  )
}

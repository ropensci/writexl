# =============================================================================
# Merged cells
# =============================================================================
#
# worksheet_merge_range() does more than record a merge: it writes the top-left
# cell's text itself and blanks the rest of the range with the same format.  A
# merge therefore carries its own content, which is why xl_merge() takes `value`
# rather than reusing whatever the data frame put there.
#
# That also settles when merges are applied.  They run *after* the sheet's rows
# have been written, so a merge over cells the data frame filled discards all
# but the top-left value -- which is exactly what merging does in Excel.
# Applying them before the rows would instead have the row loop overwrite the
# merge's own blanks.
#
# Writing to already-written rows is not possible while libxlsxwriter is
# streaming rows to disk, so any merge turns constant memory off; see
# .resolve_constant_memory().
# -----------------------------------------------------------------------------

#' Merge a range of cells
#'
#' @description
#' `xl_merge()` merges a rectangle of cells into one, as Excel's "Merge and
#' Centre" does.  Pass one or a list of them as `xl_sheet(merge = )`.
#'
#' A merged range holds a single value, so `xl_merge()` carries its own `value`
#' rather than taking it from the data frame.  Merging over cells the data frame
#' filled keeps only the merged value, exactly as merging in Excel discards
#' everything but the top-left value.
#'
#' @param range The cells to merge: an Excel range string such as `"A1:C1"`, or
#'   a `list(rows = , cols = )` spec.  It must cover more than one cell ---
#'   Excel has no single-cell merge.
#' @param value The value shown in the merged cell: a string, or anything with
#'   an [as.character()] method such as an [xl_rich_string()].  `NULL` leaves
#'   the cell empty.
#' @param format An optional [xl_format] applied to the whole merged range.
#'   Merged cells usually want `xl_align(horizontal = "center")`.
#' @return An `xl_merge` object.
#' @family worksheet features
#' @seealso [xl_sheet], [xl_format]
#' @export
#' @examples
#' xl_merge("A1:C1", "Quarterly results",
#'          format = xl_align(horizontal = "center") + xl_font(bold = TRUE))
#'
#' df <- data.frame(a = 1:3, b = 4:6)
#' sheet <- xl_sheet(df, merge = xl_merge("A5:B5", "Total",
#'                                        format = xl_font(bold = TRUE)))
#' tmp <- write_xlsx(list(Data = sheet))
xl_merge <- function(range, value = NULL, format = NULL) {
  if (missing(range) || is.null(range))
    stop("`range` must name the cells to merge", call. = FALSE)
  value <- .as_display_string(value, "value")
  if (!is.null(format) && !is_xl_format(format))
    stop("`format` must be an xl_format object", call. = FALSE)
  structure(list(range = range, value = value, format = format),
            class = "xl_merge")
}

#' @export
print.xl_merge <- function(x, ...) {
  cat(sprintf("<xl_merge: %s%s>\n",
              if (is.character(x$range)) x$range else "<spec>",
              if (is.null(x$value)) ""
              else sprintf(" %s", encodeString(x$value, quote = '"'))))
  invisible(x)
}

# Resolve a sheet's merges to the payloads C applies, registering each format.
# Normalise one merge or a list of them to a checked list.
.merge_list <- function(merge, arg = "merge") {
  if (is.null(merge)) return(list())
  ms <- if (inherits(merge, "xl_merge")) list(merge) else merge
  if (!is.list(ms))
    stop(sprintf("`%s` must be an xl_merge object or a list of them", arg),
         call. = FALSE)
  for (i in seq_along(ms))
    if (!inherits(ms[[i]], "xl_merge"))
      stop(sprintf("`%s[[%d]]` must be an xl_merge object", arg, i),
           call. = FALSE)
  ms
}

.resolve_merges <- function(el, df, reg, header_offset, props) {
  if (!inherits(el, "xl_sheet")) return(list())
  ms <- .merge_list(el$merge)
  lapply(seq_along(ms), function(i) {
    m <- ms[[i]]
    arg <- sprintf("merge[[%d]] range", i)
    q <- .xl_resolve_range(m$range, arg = arg, df = df,
                           header_offset = header_offset, allow_cell = FALSE)
    # libxlsxwriter refuses a single-cell merge, and so does Excel; catch it
    # here so the message names the merge rather than coming back as a bare
    # parameter-validation error
    if (q[1L] == q[3L] && q[2L] == q[4L])
      stop(sprintf("`%s` covers a single cell; a merge needs more than one",
                   arg), call. = FALSE)
    list(kind = "merge", range = as.integer(q),
         value = if (is.null(m$value)) NA_character_ else m$value,
         format_id = .register_format(reg,
                                      merge_xl_format(props$default_format,
                                                      m$format)))
  })
}

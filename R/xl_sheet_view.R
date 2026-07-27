# =============================================================================
# Sheet view: the tab strip and what the user sees when the sheet opens
# =============================================================================
#
# Grouped into an object for the same reason xl_page_setup() is: these are a
# cluster of rarely-used settings that would otherwise widen xl_sheet()'s
# signature past the point of being readable.
#
# `freeze` deliberately stays on xl_sheet() rather than moving here.  It is the
# single most common worksheet option -- keeping the header row visible -- and
# making the common case wordier to tidy the rare ones would be a poor trade.
# `split`, its rarely-used divider variant, does live here; the two remain
# mutually exclusive and the check spans both.
# -----------------------------------------------------------------------------

#' How a worksheet appears when it opens
#'
#' @description
#' `xl_sheet_view()` collects a worksheet's tab state and opening view: which
#' tab is active, selected or hidden, where the sheet is scrolled and selected,
#' and a few display options.  Pass it as `xl_sheet(view = )`.
#'
#' None of these affect the cell data.
#'
#' @param active Logical; make this the tab Excel opens on.  At most one sheet
#'   in a workbook may be active.
#' @param selected Logical; include this tab in the selected group.  The active
#'   sheet is always selected.
#' @param visible Logical; `FALSE` hides the sheet's tab.  A hidden sheet cannot
#'   be active or selected, the first sheet cannot be hidden unless another is
#'   made active, and at least one sheet must stay visible or Excel will not
#'   open the file.  All four rules are checked before writing.
#' @param first_tab Logical; make this the leftmost visible tab in the tab
#'   strip.  Independent of which sheet is active.
#' @param selection The cell or range selected when the sheet opens, as an Excel
#'   reference (`"B2"`, `"B2:D10"`) or a `list(rows = , cols = )` spec.
#'
#'   Excel also uses the order of a selection's corners to mark which cell in it
#'   is active; writexl does not expose that, because ranges are normalised by
#'   the shared range parser, which rejects an inverted range.
#' @param top_left The cell scrolled to the top-left of the window when the
#'   sheet opens, as an Excel reference such as `"A5"`.
#' @param hide_zero Logical; display zero values as blank cells.
#' @param right_to_left Logical; order the columns right to left, for a sheet in
#'   a right-to-left language.
#' @param split Split the sheet into scrollable panes with a visible, movable
#'   divider, given as the cell reference the split sits above and to the left
#'   of --- `"B3"` splits above row 3 and left of column B.  Mutually exclusive
#'   with `xl_sheet(freeze = )`, which does the same thing without the divider.
#'
#'   libxlsxwriter positions a split by distance, in row-height and column-width
#'   units, not by row and column number.  writexl converts the cell reference
#'   using the sheet's actual row heights and column widths, so the split lands
#'   where you asked even after resizing.  Pass
#'   `list(vertical = , horizontal = )` to give those units directly.
#'
#'   Note that libxlsxwriter derives the pane's scroll anchor back from that
#'   distance assuming default row heights, so on a sheet with resized rows or
#'   columns the divider is placed correctly but the anchor cell may be a row or
#'   two out.
#' @return An `xl_sheet_view` object.
#' @family writexl
#' @seealso [xl_sheet], [xl_page_setup]
#' @export
#' @examples
#' xl_sheet_view(active = TRUE, selection = "B2")
#' xl_sheet_view(visible = FALSE)
#'
#' df <- data.frame(x = 1:3)
#' tmp <- write_xlsx(list(
#'   Summary = xl_sheet(df, view = xl_sheet_view(active = TRUE)),
#'   Working = xl_sheet(df, view = xl_sheet_view(visible = FALSE))
#' ))
xl_sheet_view <- function(active = NA, selected = NA, visible = NA,
                          first_tab = NA, selection = NULL, top_left = NULL,
                          hide_zero = NA, right_to_left = NA, split = NULL) {
  structure(
    list(active        = .val_flag(active, "active"),
         selected      = .val_flag(selected, "selected"),
         visible       = .val_flag(visible, "visible"),
         first_tab     = .val_flag(first_tab, "first_tab"),
         selection     = selection,
         top_left      = top_left,
         hide_zero     = .val_flag(hide_zero, "hide_zero"),
         right_to_left = .val_flag(right_to_left, "right_to_left"),
         split         = split),
    class = "xl_sheet_view"
  )
}

#' @export
print.xl_sheet_view <- function(x, ...) {
  p <- .drop_null(unclass(x))
  cat(sprintf("<xl_sheet_view: %d setting%s>\n", length(p),
              if (length(p) == 1L) "" else "s"))
  for (k in names(p))
    cat(sprintf("  %s: %s\n", k, paste(format(unlist(p[[k]])), collapse = ", ")))
  invisible(x)
}

# The view settings of a sheet element, with absent ones as NULL.  Accepts a
# sheet whose `view` is unset so callers need not special-case it.
.sheet_view_of <- function(el) {
  if (!inherits(el, "xl_sheet") || is.null(el$view)) return(NULL)
  if (!inherits(el$view, "xl_sheet_view"))
    stop("`view` must be an xl_sheet_view object", call. = FALSE)
  unclass(el$view)
}

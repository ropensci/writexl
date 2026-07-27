# =============================================================================
# Sheet visibility: which tab opens, which are selected, which are hidden
# =============================================================================
#
# These are the only worksheet settings whose rules span the whole workbook, so
# unlike everything else in the sheet plan they cannot be validated one sheet at
# a time.  Excel's rules, as documented for worksheet_hide():
#
#   * a hidden sheet cannot be activated or selected -- the calls are mutually
#     exclusive;
#   * the first sheet is active by default, so it cannot be hidden unless some
#     other sheet is made active;
#   * at most one sheet may be active;
#   * at least one sheet must stay visible, or Excel refuses to open the file.
#
# libxlsxwriter enforces none of them: worksheet_hide(), worksheet_activate()
# and worksheet_select() all return void and simply set a flag.  A bad
# combination therefore produces a broken workbook with no diagnostic, so the
# checks live here and name the sheet at fault.
# -----------------------------------------------------------------------------

# One sheet's visibility settings, with absent ones as NA.
.sheet_view_flags <- function(el) {
  if (!inherits(el, "xl_sheet"))
    return(list(active = NA, selected = NA, visible = NA, first_tab = NA))
  list(active    = el$active,
       selected  = el$selected,
       visible   = el$visible,
       first_tab = el$first_tab)
}

# A readable name for a sheet in an error message.
.sheet_label <- function(nms, i) {
  if (!is.null(nms) && length(nms) >= i && !is.na(nms[i]) && nzchar(nms[i]))
    sprintf('"%s"', nms[i])
  else
    sprintf("%d", i)
}

# Validate the workbook-wide visibility rules.  Returns nothing; it exists to
# fail loudly before anything is written.
.resolve_sheet_visibility <- function(elems, nms = names(elems)) {
  n <- length(elems)
  if (!n) return(invisible(NULL))
  f <- lapply(elems, .sheet_view_flags)
  is_set <- function(k, want) vapply(f, function(x) identical(x[[k]], want),
                                     logical(1))

  hidden   <- is_set("visible", FALSE)
  active   <- is_set("active", TRUE)
  selected <- is_set("selected", TRUE)

  # every sheet hidden -> Excel will not open the file
  if (all(hidden))
    stop("every sheet is hidden; a workbook needs at least one visible sheet",
         call. = FALSE)

  # a hidden sheet cannot also be the active or a selected one
  bad <- which(hidden & (active | selected))
  if (length(bad))
    stop(sprintf(paste0("sheet %s is hidden but also marked active/selected; ",
                        "Excel cannot show a hidden sheet"),
                 .sheet_label(nms, bad[1L])), call. = FALSE)

  # only one active sheet
  if (sum(active) > 1L)
    stop(sprintf("sheets %s are both marked active; only one sheet may be active",
                 paste(vapply(which(active), function(i) .sheet_label(nms, i),
                              character(1)), collapse = " and ")),
         call. = FALSE)

  # the first sheet is active by default, so hiding it needs another active one
  if (hidden[1L] && !any(active))
    stop(sprintf(paste0("sheet %s is the first sheet and cannot be hidden ",
                        "unless another sheet is given active = TRUE, since ",
                        "Excel opens on the first sheet by default"),
                 .sheet_label(nms, 1L)), call. = FALSE)

  invisible(NULL)
}

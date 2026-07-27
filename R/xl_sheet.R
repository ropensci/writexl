# =============================================================================
# Worksheet-level formatting: xl_sheet(), xl_col_spec(), xl_row_spec()
# =============================================================================
#
# xl_col_spec and xl_row_spec are *subclasses* of xl_format: they carry the
# same formatting groups plus targeting (which columns/rows) and geometry
# (width/height/hidden/level), stored in the `xl_target` / `xl_geometry`
# attributes.  Because they are xl_format objects they combine with `+` just
# like any other format group.
# -----------------------------------------------------------------------------

# Shared constructor for column and row specs.
.new_colrow_spec <- function(kind, target, geometry, format) {
  base <- if (is.null(format)) new_xl_format() else format
  if (!is_xl_format(base))
    stop("`format` must be an xl_format object", call. = FALSE)
  out <- unclass(base)
  attr(out, "xl_geometry") <- geometry
  attr(out, "xl_target")   <- target
  class(out) <- c(paste0("xl_", kind, "_spec"), "xl_colrow_spec", "xl_format")
  out
}

#' Column and row specifications for a worksheet
#'
#' @description
#' `xl_col_spec()` and `xl_row_spec()` describe formatting and geometry for a
#' set of columns or rows within a sheet built by [xl_sheet()].  They are
#' *subclasses* of [xl_format]: they carry the usual formatting groups (so they
#' combine with `+` and the group constructors) plus a target (which
#' columns/rows) and geometry (width/height, hidden, outline level).
#'
#' @param cols Columns to target: a character vector of column names or a
#'   numeric vector of 1-based positions.
#' @param rows Rows to target: a numeric vector of 1-based data-row indices
#'   (row 1 is the first data row, ignoring the header).
#' @param width Column width (in Excel character units).
#' @param height Row height (in points).
#' @param hidden Logical; hide the column/row.
#' @param level Integer outline (grouping) level, 0--7.
#' @param format An optional [xl_format] applied to the column/row as its
#'   default cell format.  Combine groups with `+` (e.g.
#'   `xl_font(bold = TRUE) + xl_fill(background = "yellow")`).
#' @return An `xl_col_spec` / `xl_row_spec` object (also an [xl_format]).
#' @family writexl
#' @seealso [xl_sheet], [xl_format]
#' @examples
#' xl_col_spec("revenue", width = 14, format = xl_num_format("#,##0.00"))
#' xl_col_spec(c(1, 2), width = 10) + xl_font(bold = TRUE)
#' xl_row_spec(1, height = 24, format = xl_font(bold = TRUE))
#' @name xl_colrow_spec
NULL

#' @rdname xl_colrow_spec
#' @export
xl_col_spec <- function(cols, width = NA, hidden = NA, level = NA,
                        format = NULL) {
  if (missing(cols) || length(cols) < 1L)
    stop("`cols` must name or index at least one column", call. = FALSE)
  if (!is.character(cols) && !is.numeric(cols))
    stop("`cols` must be a character (names) or numeric (positions) vector",
         call. = FALSE)
  geometry <- .drop_null(list(
    width  = .val_num(width, "width", min = 0),
    hidden = .val_flag(hidden, "hidden"),
    level  = .val_int(level, "level", min = 0, max = 7)
  ))
  .new_colrow_spec("col", list(kind = "col", index = cols), geometry, format)
}

#' @rdname xl_colrow_spec
#' @export
xl_row_spec <- function(rows, height = NA, hidden = NA, level = NA,
                        format = NULL) {
  if (missing(rows) || length(rows) < 1L)
    stop("`rows` must index at least one row", call. = FALSE)
  if (!is.numeric(rows))
    stop("`rows` must be a numeric vector of 1-based data-row indices",
         call. = FALSE)
  if (any(rows < 1))
    stop("`rows` must be positive (1-based data-row indices)", call. = FALSE)
  geometry <- .drop_null(list(
    height = .val_num(height, "height", min = 0),
    hidden = .val_flag(hidden, "hidden"),
    level  = .val_int(level, "level", min = 0, max = 7)
  ))
  .new_colrow_spec("row", list(kind = "row", index = rows), geometry, format)
}

# Normalise a single spec or a list of specs to a list of specs.
.as_spec_list <- function(x, cls, arg) {
  if (is.null(x)) return(list())
  if (inherits(x, cls)) return(list(x))
  if (is.list(x)) {
    for (el in x)
      if (!inherits(el, cls))
        stop(sprintf("each element of `%s` must be an %s object", arg, cls),
             call. = FALSE)
    return(x)
  }
  stop(sprintf("`%s` must be an %s object or a list of them", arg, cls),
       call. = FALSE)
}

#' A worksheet with formatting and layout options
#'
#' @description
#' `xl_sheet()` wraps a data frame together with worksheet-level options
#' (column/row formatting and geometry, frozen panes, gridlines, tab color,
#' zoom).  Pass it anywhere [write_xlsx()] accepts a data frame; a plain data
#' frame continues to behave exactly as before.
#'
#' @param data A data frame (the sheet contents).
#' @param cols An [xl_col_spec()], or a list of them.
#' @param rows An [xl_row_spec()], or a list of them.
#' @param freeze Frozen panes: an Excel cell reference such as `"A2"` (freeze
#'   the rows above and columns left of that cell), or `list(row =, col =)`
#'   giving the number of rows/columns to freeze.
#' @param gridlines Logical; show (`TRUE`) or hide (`FALSE`) screen gridlines.
#'   `NA` leaves Excel's default.
#' @param tab_color Sheet tab color (see [xl_color]).
#' @param zoom Zoom level as a percentage (10--400).
#' @param default_row_height Default height (in points) for rows in the sheet.
#' @param auto_colwidth If `TRUE`, size each column to fit its contents (a
#'   character-count heuristic, since the xlsx format has no true "AutoFit").
#'   Columns given an explicit width via [xl_col_spec()] are left untouched.
#' @param autofilter Add an autofilter (filter dropdowns). `TRUE` covers the
#'   whole used range (header plus data); an Excel range string such as
#'   `"A1:D51"` restricts it; `FALSE` (default) adds none.
#' @param protect Protect the worksheet: `FALSE` (default) leaves it
#'   unprotected, `TRUE` applies the standard protection, and a string sets a
#'   password. A named list gives fine-grained control, e.g.
#'   `list(password = "secret", format_cells = TRUE)`; an editing option set to
#'   `TRUE` *allows* that action on the protected sheet. Available option names:
#'   `format_cells`, `format_columns`, `format_rows`, `insert_columns`,
#'   `insert_rows`, `insert_hyperlinks`, `delete_columns`, `delete_rows`,
#'   `sort`, `autofilter`, `pivot_tables`, `scenarios`, `objects`,
#'   `no_select_locked_cells`, `no_select_unlocked_cells`. Cell locking via
#'   [xl_protection()] only has an effect on a protected sheet.
#' @param comment_author Default author for this sheet's cell comments (a
#'   per-comment `author` overrides it).
#' @param show_comments If `TRUE`, all comments on the sheet are initially
#'   shown (individual comments can still be forced via `xl_comment(visible=)`).
#' @param active Logical; make this the tab Excel opens on.  At most one sheet
#'   in a workbook may be active.
#' @param selected Logical; include this tab in the selected group.  The active
#'   sheet is always selected.
#' @param visible Logical; `FALSE` hides the sheet's tab.  A hidden sheet cannot
#'   be active or selected, the first sheet cannot be hidden unless another is
#'   made active, and at least one sheet must stay visible or Excel will not
#'   open the file.
#' @param first_tab Logical; make this the leftmost visible tab in the tab
#'   strip.  This is independent of which sheet is active.
#' @param hide_zero Logical; display zero values as blank cells.
#' @param right_to_left Logical; order the columns right to left, for a sheet in
#'   a right-to-left language.
#' @param selection The cell or range selected when the sheet opens, as an Excel
#'   reference (`"B2"`, `"B2:D10"`) or a `list(rows = , cols = )` spec.
#'
#'   Excel also uses the order of a selection's corners to mark which cell in it
#'   is active; writexl does not expose that, because ranges are normalised by
#'   the shared range parser, which rejects an inverted range.
#' @param top_left The cell scrolled to the top-left of the window when the
#'   sheet opens, as an Excel reference such as `"A5"`.
#' @param page An [xl_page_setup()] describing how the sheet prints
#'   (orientation, paper size, margins, scaling, header and footer). Affects
#'   printing only, never the cell data.
#' @return An `xl_sheet` object.
#' @family writexl
#' @seealso [xl_col_spec], [xl_row_spec], [write_xlsx]
#' @export
#' @examples
#' df <- data.frame(name = c("a", "b"), revenue = c(1000.5, 2000.25))
#' sheet <- xl_sheet(
#'   df,
#'   cols   = xl_col_spec("revenue", width = 14, format = xl_num_format("#,##0.00")),
#'   freeze = "A2",
#'   tab_color = "steelblue"
#' )
#' tmp <- write_xlsx(list(Data = sheet))
xl_sheet <- function(data, cols = NULL, rows = NULL, freeze = NULL,
                     gridlines = NA, tab_color = NA, zoom = NA,
                     default_row_height = NA, auto_colwidth = FALSE,
                     autofilter = FALSE, protect = FALSE,
                     comment_author = NA, show_comments = FALSE,
                     page = NULL, active = NA, selected = NA, visible = NA,
                     first_tab = NA, hide_zero = NA, right_to_left = NA,
                     selection = NULL, top_left = NULL) {
  if (!is.data.frame(data))
    stop("`data` must be a data frame", call. = FALSE)
  if (!is.logical(auto_colwidth) || length(auto_colwidth) != 1L || is.na(auto_colwidth))
    stop("`auto_colwidth` must be TRUE or FALSE", call. = FALSE)
  if (!(is.logical(autofilter) || is.character(autofilter)) ||
      length(autofilter) != 1L)
    stop('`autofilter` must be TRUE/FALSE or an Excel range string like "A1:D51"',
         call. = FALSE)
  .validate_protect(protect)
  if (!is.logical(show_comments) || length(show_comments) != 1L || is.na(show_comments))
    stop("`show_comments` must be TRUE or FALSE", call. = FALSE)
  structure(
    list(
      data      = data,
      cols      = .as_spec_list(cols, "xl_col_spec", "cols"),
      rows      = .as_spec_list(rows, "xl_row_spec", "rows"),
      freeze    = freeze,
      gridlines = gridlines,
      tab_color = tab_color,
      zoom      = zoom,
      default_row_height = default_row_height,
      auto_colwidth = auto_colwidth,
      autofilter = autofilter,
      protect    = protect,
      comment_author = comment_author,
      show_comments  = show_comments,
      page           = page,
      active         = .val_flag(active, "active"),
      selected       = .val_flag(selected, "selected"),
      visible        = .val_flag(visible, "visible"),
      first_tab      = .val_flag(first_tab, "first_tab"),
      hide_zero      = .val_flag(hide_zero, "hide_zero"),
      right_to_left  = .val_flag(right_to_left, "right_to_left"),
      selection      = selection,
      top_left       = top_left
    ),
    class = "xl_sheet"
  )
}

# The subset of lxw_protection options exposed (chartsheet-only fields omitted).
.LXW_PROTECT_OPTS <- c(
  "no_select_locked_cells", "no_select_unlocked_cells", "format_cells",
  "format_columns", "format_rows", "insert_columns", "insert_rows",
  "insert_hyperlinks", "delete_columns", "delete_rows", "sort", "autofilter",
  "pivot_tables", "scenarios", "objects"
)

.validate_protect <- function(protect) {
  if (is.null(protect)) return(invisible())
  if (is.logical(protect) && length(protect) == 1L) return(invisible())
  if (is.character(protect) && length(protect) == 1L) return(invisible())
  if (is.list(protect)) {
    nms <- names(protect)
    if (is.null(nms) || any(nms == ""))
      stop("`protect` list must be fully named", call. = FALSE)
    bad <- setdiff(nms, c("password", .LXW_PROTECT_OPTS))
    if (length(bad))
      stop("unknown `protect` option(s): ", paste(bad, collapse = ", "),
           call. = FALSE)
    return(invisible())
  }
  stop("`protect` must be TRUE/FALSE, a password string, or a named list",
       call. = FALSE)
}

#' @export
print.xl_sheet <- function(x, ...) {
  cat(sprintf("<xl_sheet: %d rows x %d cols>\n", nrow(x$data), ncol(x$data)))
  if (length(x$cols)) cat("  column specs:", length(x$cols), "\n")
  if (length(x$rows)) cat("  row specs:", length(x$rows), "\n")
  if (!is.null(x$freeze)) cat("  freeze:", format(x$freeze), "\n")
  invisible(x)
}

# --- resolution: turn a sheet into the plan C consumes --------------------

# Build the (flag, password, options) triple C uses for worksheet protection.
.resolve_protect <- function(protect) {
  out <- list(flag = 0L, password = NA_character_, options = NULL)
  if (isTRUE(protect)) {
    out$flag <- 1L
  } else if (is.character(protect) && length(protect) == 1L) {
    out$flag <- 1L; out$password <- protect
  } else if (is.list(protect)) {
    out$flag <- 1L
    if (!is.null(protect$password)) out$password <- as.character(protect$password)
    optnames <- setdiff(names(protect), "password")
    if (length(optnames)) {
      # a named list (not a bare vector) so the C side can look up by name
      out$options <- stats::setNames(
        lapply(.LXW_PROTECT_OPTS, function(nm) as.integer(isTRUE(protect[[nm]]))),
        .LXW_PROTECT_OPTS)
    }
  }
  out
}

# Estimate a column width (in Excel character units) from its rendered content.
# There is no native "AutoFit" in the file format, so this is a heuristic based
# on character counts of the displayed values and the header.
.auto_col_width <- function(col, header) {
  vals_w <- .content_nchar(col)
  hdr_w  <- if (nzchar(header)) nchar(header, type = "width") else 0L
  w <- max(c(hdr_w, vals_w, 0L), na.rm = TRUE)
  min(w + 1L, 255L)   # small padding; Excel's maximum column width is 255
}

# Per-cell display width for a column (0 for blank/NA cells).
.content_nchar <- function(col) {
  if (inherits(col, "xl_cell_general")) {
    vapply(unclass(col), function(rec) {
      s <- .cell_display_string(rec)
      if (is.na(s)) 0L else nchar(s, type = "width")
    }, integer(1))
  } else {
    x <- format(col, trim = TRUE)
    w <- nchar(x, type = "width")
    w[is.na(col)] <- 0L
    w
  }
}

# The string a general cell displays (hyperlink display / value / formula).
.cell_display_string <- function(rec) {
  # A rich string has been flattened to its runs before widths are measured, so
  # the displayed text is the runs joined -- rec$value is gone by now.
  if (!is.null(rec$rich_text))
    return(paste(rec$rich_text, collapse = ""))
  v <- rec$value
  if (!is.null(v) && length(v) && !all(is.na(v)))
    return(format(v, trim = TRUE)[1L])
  fm <- rec$formula
  if (!is.null(fm) && length(fm) && !is.na(fm)) return(as.character(fm))
  hl <- rec$hyperlink
  if (is.character(hl) && length(hl) && !is.na(hl)) return(hl)
  if (is.list(hl) && !is.null(hl$url)) return(as.character(hl$url))
  NA_character_
}

# Resolve targeted columns (names or positions) to 1-based indices.  `arg` names
# the calling argument in the error messages.
.resolve_col_index <- function(index, colnames, arg = "cols") {
  if (is.character(index)) {
    idx <- match(index, colnames)
    if (anyNA(idx))
      stop(sprintf("`%s`: unknown column(s): ", arg),
           paste(index[is.na(idx)], collapse = ", "), call. = FALSE)
    idx
  } else {
    idx <- as.integer(index)
    if (any(idx < 1L | idx > length(colnames)))
      stop(sprintf("`%s`: column index out of range", arg), call. = FALSE)
    idx
  }
}

# --- sheet overlays -------------------------------------------------------
#
# A sheet overlay is a range-scoped worksheet feature applied in one pass after
# the per-sheet scalar options and *before* the row loop -- required under
# libxlsxwriter's constant-memory mode, where each row is flushed as it is
# written.  Each payload is a named list carrying a `kind` string that C
# dispatches on, so a new feature adds a payload kind rather than a new
# argument to C_write_data_frame_list().
#
# The autofilter is the first (and currently only) kind.  `xl_sheet()` has no
# public `overlay` argument yet; an `overlay` element set on the sheet object
# is picked up here so the mechanism is reachable for testing and for the
# features that will populate it.

# The autofilter payload for an already-resolved 0-based range.
.overlay_autofilter <- function(range) {
  list(kind = "autofilter", range = as.integer(range))
}

# Normalise a single payload or a list of payloads to a list of payloads.
.as_overlay_list <- function(x) {
  if (is.null(x)) return(list())
  if (!is.list(x))
    stop("`overlay` must be an overlay payload or a list of them", call. = FALSE)
  if (!is.null(x$kind)) list(x) else x
}

# Build the per-sheet plan (column/row/scalar options) that C applies.
.resolve_sheet_plan <- function(el, df, reg, header_offset, props) {
  ncols  <- length(df)
  cnames <- names(df)

  # per-column base format: workbook default_format, plus the date/time number
  # format for Date/POSIXct columns
  col_fmt  <- vector("list", ncols)
  col_width  <- rep(NA_real_, ncols)
  col_hidden <- rep(NA_integer_, ncols)
  col_level  <- rep(NA_integer_, ncols)
  explicit_width <- rep(FALSE, ncols)   # columns whose width the user set
  for (i in seq_len(ncols)) {
    base <- props$default_format
    if (inherits(df[[i]], "POSIXct")) {
      base <- merge_xl_format(base, props$datetime_format)
      col_width[i] <- props$datetime_col_width
    } else if (inherits(df[[i]], "Date")) {
      base <- merge_xl_format(base, props$date_format)
      col_width[i] <- props$date_col_width
    }
    col_fmt[[i]] <- base
  }

  row_row <- integer(0); row_height <- numeric(0)
  row_fmt_id <- integer(0); row_hidden <- integer(0); row_level <- integer(0)
  freeze <- c(-1L, -1L)
  gridlines <- -1L; tab_color <- -1L; zoom <- 0L; default_row_height <- NA_real_
  overlay <- list()
  protect <- list(flag = 0L, password = NA_character_, options = NULL)
  # comment defaults are worksheet-scoped (as in libxlsxwriter); a comment's own
  # author overrides the sheet default
  comment_author <- NA_character_
  show_comments <- FALSE
  page_payload <- NULL
  view <- list()

  if (inherits(el, "xl_sheet")) {
    # column specs
    for (spec in el$cols) {
      idx <- .resolve_col_index(attr(spec, "xl_target")$index, cnames)
      geo <- attr(spec, "xl_geometry")
      for (i in idx) {
        col_fmt[[i]] <- merge_xl_format(col_fmt[[i]], spec)
        if (!is.null(geo$width)) { col_width[i] <- geo$width; explicit_width[i] <- TRUE }
        if (!is.null(geo$hidden)) col_hidden[i] <- as.integer(isTRUE(geo$hidden))
        if (!is.null(geo$level))  col_level[i]  <- as.integer(geo$level)
      }
    }
    # row specs
    for (spec in el$rows) {
      geo <- attr(spec, "xl_geometry")
      fid <- .register_format(reg, spec)
      for (r in as.integer(attr(spec, "xl_target")$index)) {
        row_row    <- c(row_row, (r - 1L) + header_offset)
        row_height <- c(row_height, if (!is.null(geo$height)) geo$height else NA_real_)
        row_fmt_id <- c(row_fmt_id, fid)
        row_hidden <- c(row_hidden, if (!is.null(geo$hidden)) as.integer(isTRUE(geo$hidden)) else NA_integer_)
        row_level  <- c(row_level, if (!is.null(geo$level)) as.integer(geo$level) else NA_integer_)
      }
    }
    freeze <- .parse_freeze(el$freeze)
    gridlines <- if (is.na(el$gridlines)) -1L else if (isTRUE(el$gridlines)) 3L else 0L
    tab_color <- if (is.na(el$tab_color)) -1L else xl_color(el$tab_color)
    zoom <- if (is.na(el$zoom)) 0L else as.integer(el$zoom)
    default_row_height <- if (is.na(el$default_row_height)) NA_real_ else as.numeric(el$default_row_height)
    af <- el$autofilter
    if (is.character(af)) {
      overlay <- c(overlay, list(.overlay_autofilter(
        .parse_range(af, "autofilter", df, header_offset))))
    } else if (isTRUE(af)) {
      last_row <- (nrow(df) - 1L) + header_offset
      if (ncols > 0L && last_row >= 0L)
        overlay <- c(overlay, list(.overlay_autofilter(
          c(0L, 0L, as.integer(last_row), ncols - 1L))))
    }
    overlay <- c(overlay, .as_overlay_list(el$overlay))
    protect <- .resolve_protect(el$protect)
    page_payload <- .page_setup_payload(el$page, df, header_offset)
    for (k in c("active", "selected", "visible", "first_tab", "hide_zero",
                "right_to_left"))
      if (!is.null(el[[k]])) view[[k]] <- as.integer(isTRUE(el[[k]]))
    if (!is.null(el$selection))
      view$selection <- .xl_resolve_range(el$selection, arg = "selection",
                                          df = df, header_offset = header_offset,
                                          allow_cell = TRUE)
    if (!is.null(el$top_left))
      view$top_left <- .xl_resolve_range(el$top_left, arg = "top_left",
                                         df = df, header_offset = header_offset,
                                         allow_cell = TRUE)[1:2]
    comment_author <- el$comment_author
    show_comments <- isTRUE(el$show_comments)
    # auto column widths (for columns the user did not size explicitly)
    if (isTRUE(el$auto_colwidth)) {
      for (i in seq_len(ncols)) {
        if (explicit_width[i]) next
        header <- if (header_offset > 0L) cnames[i] else ""
        col_width[i] <- .auto_col_width(df[[i]], header)
      }
    }
  }

  col_format_id <- vapply(col_fmt, function(f) .register_format(reg, f), integer(1))

  list(
    col_width = col_width, col_format_id = col_format_id,
    col_hidden = col_hidden, col_level = col_level,
    row_row = row_row, row_height = row_height, row_format_id = row_fmt_id,
    row_hidden = row_hidden, row_level = row_level,
    freeze_row = freeze[1L], freeze_col = freeze[2L],
    gridlines = gridlines, tab_color = tab_color, zoom = zoom,
    default_row_height = default_row_height,
    overlay = overlay, protect = protect$flag,
    protect_password = protect$password, protect_options = protect$options,
    comment_author = as.character(comment_author),
    show_comments = as.integer(show_comments),
    page = page_payload,
    view = if (length(view)) view else NULL
  )
}

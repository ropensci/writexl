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
                     autofilter = FALSE, protect = FALSE) {
  if (!is.data.frame(data))
    stop("`data` must be a data frame", call. = FALSE)
  if (!is.logical(auto_colwidth) || length(auto_colwidth) != 1L || is.na(auto_colwidth))
    stop("`auto_colwidth` must be TRUE or FALSE", call. = FALSE)
  if (!(is.logical(autofilter) || is.character(autofilter)) ||
      length(autofilter) != 1L)
    stop('`autofilter` must be TRUE/FALSE or an Excel range string like "A1:D51"',
         call. = FALSE)
  .validate_protect(protect)
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
      protect    = protect
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

# Convert column letters ("A", "AB") to a 0-based column index.
.col_letters_to_index <- function(letters_str) {
  chars <- utf8ToInt(toupper(letters_str)) - utf8ToInt("A") + 1L
  idx <- 0L
  for (v in chars) idx <- idx * 26L + v
  idx - 1L
}

# Parse a freeze specification into c(freeze_row, freeze_col) (counts), or
# c(-1, -1) for none.
.parse_freeze <- function(freeze) {
  if (is.null(freeze) || (length(freeze) == 1L && is.na(freeze)))
    return(c(-1L, -1L))
  if (is.list(freeze)) {
    r <- if (!is.null(freeze$row)) as.integer(freeze$row) else 0L
    cc <- if (!is.null(freeze$col)) as.integer(freeze$col) else 0L
    return(c(r, cc))
  }
  if (is.character(freeze) && length(freeze) == 1L) {
    m <- regmatches(freeze, regexec("^([A-Za-z]+)([0-9]+)$", freeze))[[1]]
    if (length(m) != 3L)
      stop('`freeze` must be a cell reference like "A2"', call. = FALSE)
    return(c(as.integer(m[3L]) - 1L, .col_letters_to_index(m[2L])))
  }
  stop('`freeze` must be a cell reference ("A2") or list(row =, col =)',
       call. = FALSE)
}

# Parse an Excel range ("A1:D51") into 0-based c(first_row, first_col,
# last_row, last_col).
.parse_range <- function(rng) {
  parts <- strsplit(rng, ":", fixed = TRUE)[[1]]
  cell <- function(s) {
    m <- regmatches(s, regexec("^([A-Za-z]+)([0-9]+)$", s))[[1]]
    if (length(m) != 3L)
      stop('`autofilter` range must look like "A1:D51"', call. = FALSE)
    c(as.integer(m[3L]) - 1L, .col_letters_to_index(m[2L]))
  }
  if (length(parts) != 2L)
    stop('`autofilter` range must look like "A1:D51"', call. = FALSE)
  a <- cell(parts[1L]); b <- cell(parts[2L])
  c(a[1L], a[2L], b[1L], b[2L])
}

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

# Resolve targeted columns (names or positions) to 1-based indices.
.resolve_col_index <- function(index, colnames) {
  if (is.character(index)) {
    idx <- match(index, colnames)
    if (anyNA(idx))
      stop("unknown column(s): ", paste(index[is.na(idx)], collapse = ", "),
           call. = FALSE)
    idx
  } else {
    idx <- as.integer(index)
    if (any(idx < 1L | idx > length(colnames)))
      stop("column index out of range", call. = FALSE)
    idx
  }
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
  autofilter <- c(-1L, -1L, -1L, -1L)
  protect <- list(flag = 0L, password = NA_character_, options = NULL)

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
      autofilter <- .parse_range(af)
    } else if (isTRUE(af)) {
      last_row <- (nrow(df) - 1L) + header_offset
      if (ncols > 0L && last_row >= 0L)
        autofilter <- c(0L, 0L, as.integer(last_row), ncols - 1L)
    }
    protect <- .resolve_protect(el$protect)
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
    autofilter = autofilter, protect = protect$flag,
    protect_password = protect$password, protect_options = protect$options
  )
}

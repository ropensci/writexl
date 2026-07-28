# =============================================================================
# Worksheet tables
# =============================================================================
#
# A table is a named, styled range that Excel treats as a unit: banded rows, a
# filter dropdown in the header, structured references in formulas, and an
# optional total row.
#
# The thing to know before reading further is that worksheet_add_table() writes
# cells of its own.  _write_table_column_data() in worksheet.c writes each
# column's header string into the header row, and _set_default_table_columns()
# defaults those strings to "Column 1", "Column 2", ...  A table added over a
# writexl data frame would therefore replace the header row with generic
# captions.
#
# It is worse than losing the captions: the table part records the column names
# independently of the cells, and Excel treats a mismatch between the two as a
# corrupt file.  So xl_table() always sends the data frame's column names
# through as the headers -- the generic captions are not reachable by accident.
#
# That also settles the ordering.  Tables are applied *after* the sheet's rows,
# beside merges, so the table's header and total-row writes land last; applying
# them first would let the row loop overwrite them and lose header_format.
# Writing over already-written rows is impossible while libxlsxwriter streams,
# and worksheet_add_table() refuses outright in that mode, so any table turns
# constant memory off -- see .resolve_constant_memory().
# -----------------------------------------------------------------------------

# Style types, and the style numbers each one allows.  libxlsxwriter only warns
# on a bad number and silently substitutes Medium 9, so the checking is ours.
.LXW_TABLE_STYLE_TYPE <- c(none = 0L, light = 1L, medium = 2L, dark = 3L)
.TABLE_STYLE_RANGE <- list(light = c(0L, 21L), medium = c(1L, 28L),
                           dark = c(1L, 11L))

# The total-row functions.  The enum is sparse (101-105, 107, 109, 110), which
# is a good reason not to make callers know it.
.LXW_TABLE_FUNCTION <- c(
  average = 101L, count_nums = 102L, count = 103L, max = 104L, min = 105L,
  std_dev = 107L, sum = 109L, var = 110L
)

# Parse "medium 9" / "light 21" / "none" into the type and number C wants.
.parse_table_style <- function(style) {
  if (is.null(style)) return(list(type = 2L, number = 9L))
  if (!is.character(style) || length(style) != 1L || is.na(style))
    stop('`style` must be a single string like "medium 9" or "none"',
         call. = FALSE)
  s <- tolower(trimws(style))
  # "none" is Light 0 in libxlsxwriter's numbering
  if (identical(s, "none")) return(list(type = 1L, number = 0L))
  parts <- strsplit(s, "[[:space:]]+")[[1L]]
  if (length(parts) != 2L)
    stop(sprintf(paste0('`style` must be "none" or a type and number like ',
                        '"medium 9" (got "%s")'), style), call. = FALSE)
  type <- parts[1L]
  if (!type %in% names(.TABLE_STYLE_RANGE))
    stop(sprintf('`style` type must be "light", "medium" or "dark" (got "%s")',
                 parts[1L]), call. = FALSE)
  num <- suppressWarnings(as.integer(parts[2L]))
  rng <- .TABLE_STYLE_RANGE[[type]]
  if (is.na(num) || num < rng[1L] || num > rng[2L])
    stop(sprintf("`style`: %s styles are numbered %d to %d (got \"%s\")",
                 type, rng[1L], rng[2L], parts[2L]), call. = FALSE)
  list(type = unname(.LXW_TABLE_STYLE_TYPE[[type]]), number = num)
}

#' Describe a column of a worksheet table
#'
#' @description
#' `xl_table_column()` overrides what [xl_table()] does with one column: its
#' header caption, a formula filling the column, what its total-row cell shows,
#' and the formats applied to the header and the data cells.
#'
#' Only the columns you name need an entry; the rest take the data frame's
#' column name as their header and are left otherwise alone.
#'
#' @param col The column: a name or a 1-based position within the table's range.
#' @param header The header caption.  Defaults to the data frame's column name
#'   --- which is also what makes a table safe to add, since Excel rejects a
#'   file whose table definition and header cells disagree.
#' @param formula A formula filling every data cell of the column, usually with
#'   a structured reference such as `"=SUM(Sales[@[Q1]:[Q4]])"`.  Because those
#'   name the table, see the `name` argument of [xl_table()].
#' @param total The function shown in this column's total-row cell: one of
#'   `"sum"`, `"average"`, `"count"`, `"count_nums"`, `"max"`, `"min"`,
#'   `"std_dev"` or `"var"`.  Needs `total_row = TRUE` on the table.
#' @param total_label A string for the total-row cell instead of a function ---
#'   typically `"Total"` under the first column.
#' @param format An [xl_format] applied to the column's data cells.
#' @param header_format An [xl_format] applied to the column's header cell.
#' @return An `xl_table_column` object.
#' @family writexl
#' @seealso [xl_table], [xl_sheet]
#' @export
#' @examples
#' xl_table_column("qty", total = "sum")
#' xl_table_column("fruit", total_label = "Total")
#' xl_table_column("margin", formula = "=Sales[@revenue] * 0.3")
xl_table_column <- function(col, header = NULL, formula = NULL, total = NULL,
                            total_label = NULL, format = NULL,
                            header_format = NULL) {
  if (missing(col) || is.null(col))
    stop("`col` must name or index the table column", call. = FALSE)
  str1 <- function(x, arg) {
    if (is.null(x)) return(NULL)
    if (!is.character(x) || length(x) != 1L || is.na(x))
      stop(sprintf("`%s` must be a single non-NA string", arg), call. = FALSE)
    x
  }
  fmt <- function(x, arg) {
    if (is.null(x)) return(NULL)
    if (!is_xl_format(x))
      stop(sprintf("`%s` must be an xl_format object", arg), call. = FALSE)
    x
  }
  formula <- str1(formula, "formula")
  if (!is.null(formula) && !startsWith(formula, "="))
    stop("`formula` must start with '='", call. = FALSE)
  tot <- .val_enum(total, names(.LXW_TABLE_FUNCTION), "total")
  if (!is.null(tot) && !is.null(total_label))
    stop("give either `total` or `total_label`, not both", call. = FALSE)
  structure(.drop_null(list(
    col = col, header = str1(header, "header"), formula = formula,
    total = tot, total_label = str1(total_label, "total_label"),
    format = fmt(format, "format"),
    header_format = fmt(header_format, "header_format")
  )), class = "xl_table_column")
}

#' @export
print.xl_table_column <- function(x, ...) {
  p <- unclass(x)
  # `[[` not `$`: see .table_columns_payload() -- `$total` would partial-match
  # `total_label` on a column that only carries a label
  cat(sprintf("<xl_table_column: %s%s%s>\n", p[["col"]],
              if (!is.null(p[["formula"]])) " formula" else "",
              if (!is.null(p[["total"]])) paste0(" total=", p[["total"]]) else ""))
  invisible(x)
}

#' Add a worksheet table
#'
#' @description
#' `xl_table()` turns a range into an Excel table: a named, styled block with
#' banded rows, a filter dropdown in its header, and an optional total row.
#' Pass one or a list of them as `xl_sheet(table = )`.
#'
#' Column headers default to the data frame's column names.  That is not just a
#' convenience: Excel records a table's column names separately from the header
#' cells and rejects a file where the two disagree, so the default is what makes
#' a table safe to add at all.
#'
#' @section Table names:
#' Excel requires table names to be unique across the workbook, and formulas
#' refer to a table by name.  `name = NULL` (the default) has writexl generate a
#' unique name from the sheet name and write it explicitly, so it cannot shift
#' when another table is added elsewhere.  Give a string to choose your own; it
#' is validated and checked for collisions.  Give `NA` to write no name at all
#' and let Excel assign `Table1`, `Table2`, ... by insertion order --- which
#' makes any formula naming the table fragile, so that combination warns.
#'
#' @section The total row:
#' A total row is written by libxlsxwriter below the data, and the sheet's row
#' plan does not know about it: `auto_colwidth` does not measure it, and an
#' [xl_row_spec()] aimed at that row will fight it.  It has no cells of its own
#' until a column gives a `total` or `total_label`.
#'
#' @param range The cells the table covers, as an Excel range string or a
#'   `list(rows = , cols = )` spec.  Defaults to the sheet's used range,
#'   including the header row and, with `total_row = TRUE`, one row below.
#' @param name The table's name.  See "Table names".
#' @param style The table style: `"none"`, or a type and number such as
#'   `"medium 9"` (the default), `"light 21"` or `"dark 11"`.  Light styles are
#'   numbered 0--21, medium 1--28 and dark 1--11.
#' @param header_row Show the header row.  Turning it off also removes the
#'   filter dropdown, as it does in Excel.
#' @param autofilter Show the filter dropdown in the header row.
#' @param banded_rows,banded_columns Alternating row / column shading.
#' @param first_column,last_column Highlight the first / last column.
#' @param total_row Add a total row below the data.  See "The total row".
#' @param columns One [xl_table_column()], or a list of them, overriding
#'   individual columns.
#' @return An `xl_table` object.
#' @family writexl
#' @seealso [xl_table_column], [xl_sheet]
#' @export
#' @examples
#' xl_table(style = "light 9")
#' xl_table(name = "Sales", total_row = TRUE,
#'          columns = list(xl_table_column("fruit", total_label = "Total"),
#'                         xl_table_column("qty", total = "sum")))
xl_table <- function(range = NULL, name = NULL, style = "medium 9",
                     header_row = TRUE, autofilter = TRUE, banded_rows = TRUE,
                     banded_columns = FALSE, first_column = FALSE,
                     last_column = FALSE, total_row = FALSE, columns = NULL) {
  flag <- function(x, arg) {
    if (!is.logical(x) || length(x) != 1L || is.na(x))
      stop(sprintf("`%s` must be TRUE or FALSE", arg), call. = FALSE)
    isTRUE(x)
  }
  # NA means "write no name"; NULL means "generate one"
  if (!is.null(name)) {
    if (length(name) != 1L)
      stop("`name` must be a single string, NULL or NA", call. = FALSE)
    if (!is.na(name)) .check_table_name(name)
  }
  sty <- .parse_table_style(style)
  cols <- .table_column_list(columns)

  header <- flag(header_row, "header_row")
  af <- flag(autofilter, "autofilter")
  # libxlsxwriter forces this itself; mirroring it keeps the object honest
  if (!header) af <- FALSE
  tot <- flag(total_row, "total_row")

  if (!tot) {
    # `[[` not `$`: `$total` partial-matches `total_label` once the NULLs are
    # dropped, so the two conditions would not be independent
    bad <- vapply(cols, function(cc)
      !is.null(cc[["total"]]) || !is.null(cc[["total_label"]]), logical(1))
    if (any(bad))
      stop(sprintf(paste0("column %s gives a total, but the table has no total ",
                          "row; set `total_row = TRUE`"),
                   paste(vapply(cols[bad], function(cc) sprintf('"%s"', cc$col),
                                character(1)), collapse = ", ")),
           call. = FALSE)
  }
  if (!is.null(name) && length(name) == 1L && is.na(name) &&
      any(vapply(cols, function(cc) !is.null(cc$formula), logical(1))))
    warning("`name = NA` leaves Excel to name the table (Table1, Table2, ...), ",
            "but a column formula refers to the table by name, so it will ",
            "break if the numbering shifts. Give `name` a string instead.",
            call. = FALSE)

  structure(.drop_null(list(
    range = range, name = name,
    style_type = sty$type, style_number = sty$number,
    header_row = header, autofilter = af, banded_rows = flag(banded_rows, "banded_rows"),
    banded_columns = flag(banded_columns, "banded_columns"),
    first_column = flag(first_column, "first_column"),
    last_column = flag(last_column, "last_column"),
    total_row = tot,
    columns = if (length(cols)) cols else NULL
  )), class = "xl_table")
}

#' @export
print.xl_table <- function(x, ...) {
  p <- unclass(x)
  cat(sprintf("<xl_table: %s%s%s>\n",
              if (is.null(p$name)) "unnamed" else
                if (is.na(p$name)) "named by Excel" else p$name,
              if (isTRUE(p$total_row)) ", total row" else "",
              sprintf(", %d column override(s)", length(p$columns))))
  invisible(x)
}

# Normalise one table or a list of them.
.table_list <- function(table, arg = "table") {
  if (is.null(table)) return(list())
  ts <- if (inherits(table, "xl_table")) list(table) else table
  if (!is.list(ts))
    stop(sprintf("`%s` must be an xl_table object or a list of them", arg),
         call. = FALSE)
  for (i in seq_along(ts))
    if (!inherits(ts[[i]], "xl_table"))
      stop(sprintf("`%s[[%d]]` must be an xl_table object", arg, i),
           call. = FALSE)
  ts
}

# --- Table names, resolved across the whole workbook -------------------------
#
# Excel requires table names to be unique workbook-wide, and libxlsxwriter never
# checks that -- two tables sharing a name reach the file and Excel offers to
# repair it.  Uniqueness therefore cannot be settled inside xl_table(), which
# sees one table; it is settled here, where the sheets are assembled, next to
# the sheet-name dedupe that has the same shape.

# Turn any string into something Excel will accept as a table name.
.sanitize_table_name <- function(x) {
  chars <- strsplit(x, "", fixed = TRUE)[[1L]]
  chars[chars %in% .TABLE_NAME_BAD_CHARS] <- "_"
  out <- paste(chars, collapse = "")
  if (!nzchar(out)) out <- "Table"
  if (grepl("^[0-9]", out)) out <- paste0("T_", out)
  # a bare C/R or a cell-reference lookalike is reserved; suffixing clears both
  if (toupper(out) %in% c("C", "R") || grepl("^[A-Za-z]{1,3}[0-9]+$", out))
    out <- paste0(out, "_tbl")
  if (nchar(out) > 250L) out <- substr(out, 1L, 250L)
  out
}

# Append _2, _3, ... until the name is free.  Excel compares table names
# case-insensitively, so "Sales" collides with "sales" and the comparison has
# to be folded on both sides -- folding only the stored set silently lets every
# generated name through.
.unique_table_name <- function(base, taken) {
  taken <- tolower(taken)
  if (!tolower(base) %in% taken) return(base)
  i <- 2L
  repeat {
    cand <- paste0(base, "_", i)
    if (!tolower(cand) %in% taken) return(cand)
    i <- i + 1L
  }
}

# Resolve every table's name across the workbook.  Returns one character vector
# per sheet, NA meaning "write no name and let Excel number it".
.resolve_table_names <- function(elems, sheet_names) {
  tabs <- lapply(elems, function(el)
    if (inherits(el, "xl_sheet")) .table_list(el$table) else list())
  if (!length(unlist(tabs, recursive = FALSE))) return(rep(list(character(0)), length(elems)))

  # explicit names first: they are fixed points, so a collision between two of
  # them cannot be resolved by renaming either one
  explicit <- list()
  for (i in seq_along(tabs)) {
    for (k in seq_along(tabs[[i]])) {
      nm <- unclass(tabs[[i]][[k]])$name
      if (is.null(nm) || is.na(nm)) next
      prev <- explicit[[tolower(nm)]]
      if (!is.null(prev))
        stop(sprintf(paste0("two tables are both named \"%s\" (on sheet \"%s\" ",
                            "and sheet \"%s\"); Excel requires table names to ",
                            "be unique across the workbook"),
                     nm, prev, sheet_names[i]), call. = FALSE)
      explicit[[tolower(nm)]] <- sheet_names[i]
    }
  }

  taken <- names(explicit)
  out <- vector("list", length(elems))
  for (i in seq_along(tabs)) {
    nms <- character(length(tabs[[i]]))
    base <- .sanitize_table_name(sheet_names[i])
    for (k in seq_along(tabs[[i]])) {
      nm <- unclass(tabs[[i]][[k]])$name
      if (is.null(nm)) {
        gen <- .unique_table_name(base, taken)
        taken <- c(taken, tolower(gen))
        nms[k] <- gen
      } else if (is.na(nm)) {
        nms[k] <- NA_character_
      } else {
        nms[k] <- nm
      }
    }
    out[[i]] <- nms
  }
  out
}

# --- Conflicts -----------------------------------------------------------
#
# Excel drops or repairs each of these rather than reporting them, and
# libxlsxwriter passes them all through, so they are refused here.  Each
# message names the table's range so the offending pair can be found.

# Do two 0-based c(first_row, first_col, last_row, last_col) quads overlap?
.ranges_overlap <- function(a, b) {
  a[1L] <= b[3L] && b[1L] <= a[3L] && a[2L] <= b[4L] && b[2L] <= a[4L]
}

.range_a1 <- function(r) {
  cell <- function(row, col) {
    letters_of <- function(i) {
      out <- ""
      repeat {
        out <- paste0(LETTERS[(i %% 26L) + 1L], out)
        i <- i %/% 26L - 1L
        if (i < 0L) break
      }
      out
    }
    paste0(letters_of(col), row + 1L)
  }
  paste0(cell(r[1L], r[2L]), ":", cell(r[3L], r[4L]))
}

# `tables` is the resolved payload list; the rest come off the sheet element.
.check_table_conflicts <- function(el, df, header_offset, tables) {
  if (!length(tables)) return(invisible(NULL))
  ranges <- lapply(tables, function(t) t$range)

  refuse <- function(rng, what)
    stop(sprintf("the table at %s %s", .range_a1(rng), what), call. = FALSE)

  for (rng in ranges) {
    # a table brings its own autofilter; a second one over the same cells is
    # what Excel repairs the file over
    af <- el$autofilter
    if (isTRUE(af) || is.character(af)) {
      afr <- if (is.character(af))
        .parse_range(af, "autofilter", df, header_offset)
      else c(0L, 0L, (nrow(df) - 1L) + header_offset, length(df) - 1L)
      if (.ranges_overlap(rng, afr))
        # the remedy is the sheet's autofilter, not the table's: turning the
        # table's off leaves the sheet-level one over the same cells
        refuse(rng, paste0("overlaps `xl_sheet(autofilter =)`; a table brings ",
                           "its own filter dropdown, and Excel allows only one ",
                           "over a range. Drop the sheet's `autofilter`, or ",
                           "point it at cells outside the table."))
    }
    # libxlsxwriter: "Filter conditions within the table are not supported"
    for (f in .filter_list(el$filter)) {
      idx <- .resolve_col_index(unclass(f)$col, names(df), "filter col") - 1L
      if (idx >= rng[2L] && idx <= rng[4L])
        refuse(rng, sprintf(paste0("covers column \"%s\", which `filter` also ",
                                   "filters; filter criteria inside a table ",
                                   "are not supported"),
                            names(df)[idx + 1L]))
    }
    # Excel forbids merged cells inside a table
    for (m in .merge_list(el$merge)) {
      mr <- .xl_resolve_range(unclass(m)$range, arg = "merge range", df = df,
                              header_offset = header_offset, allow_cell = FALSE)
      if (.ranges_overlap(rng, mr))
        refuse(rng, sprintf(paste0("overlaps the merge at %s; Excel does not ",
                                   "allow merged cells inside a table"),
                            .range_a1(mr)))
    }
  }
  # two tables may not overlap each other either
  if (length(ranges) > 1L)
    for (i in seq_len(length(ranges) - 1L))
      for (j in seq(i + 1L, length(ranges)))
        if (.ranges_overlap(ranges[[i]], ranges[[j]]))
          stop(sprintf("the tables at %s and %s overlap",
                       .range_a1(ranges[[i]]), .range_a1(ranges[[j]])),
               call. = FALSE)
  invisible(NULL)
}

# --- The C payload -----------------------------------------------------------
#
# Every column of the table's range gets an entry, whether or not the caller
# named it, because _write_table_column_data() writes a header string for each
# one and defaults it to "Column 1", "Column 2", ...  Sending the data frame's
# own names through is what keeps the table definition and the header cells in
# agreement, which is the difference between a file Excel opens and one it
# offers to repair.
.resolve_tables <- function(el, df, reg, header_offset, props, table_names) {
  if (!inherits(el, "xl_sheet")) return(list())
  ts <- .table_list(el$table)
  if (!length(ts)) return(list())
  if (nrow(df) < 1L)
    stop("a table needs at least one data row", call. = FALSE)

  out <- lapply(seq_along(ts), function(i) {
    p <- unclass(ts[[i]])
    rng <- if (is.null(p$range)) {
      last_row <- (nrow(df) - 1L) + header_offset + as.integer(p$total_row)
      c(0L, 0L, as.integer(last_row), length(df) - 1L)
    } else {
      .xl_resolve_range(p$range, arg = "table range", df = df,
                        header_offset = header_offset, allow_cell = FALSE)
    }
    # libxlsxwriter counts the header row but not the total row as data
    non_header <- rng[3L] - rng[1L] - as.integer(p$total_row) +
      as.integer(!p$header_row)
    if (non_header < 1L)
      stop("a table needs at least one row that is not the header or the ",
           "total row", call. = FALSE)

    # A column past the data has no name to take, and the generic "Column N"
    # caption is exactly what this design exists to make unreachable.
    if (rng[4L] > length(df) - 1L)
      stop(sprintf(paste0("the table at %s reaches column %d, but the data ",
                          "frame has %d; writexl has no header name for the ",
                          "extra column(s)"),
                   .range_a1(rng), rng[4L] + 1L, length(df)), call. = FALSE)
    # worksheet_add_table() writes its header captions into the range's first
    # row.  If that is not the sheet's header row it lands on data, silently
    # replacing values with column names.
    if (p$header_row && header_offset > 0L && rng[1L] != 0L)
      stop(sprintf(paste0("the table at %s has a header row, but its first row ",
                          "is not the sheet's header row; the table's captions ",
                          "would overwrite data. Start the table at row 1, or ",
                          "set `header_row = FALSE`."),
                   .range_a1(rng)), call. = FALSE)

    cols <- .table_columns_payload(p, df, rng, reg, props)
    ent <- list(range = as.integer(rng),
                style_type = p$style_type, style_number = p$style_number,
                header_row = as.integer(p$header_row),
                autofilter = as.integer(p$autofilter),
                banded_rows = as.integer(p$banded_rows),
                banded_columns = as.integer(p$banded_columns),
                first_column = as.integer(p$first_column),
                last_column = as.integer(p$last_column),
                total_row = as.integer(p$total_row),
                columns = cols)
    nm <- if (i <= length(table_names)) table_names[[i]] else NA_character_
    if (!is.na(nm)) ent$name <- nm
    ent
  })
  .check_table_conflicts(el, df, header_offset, out)
  out
}

# One entry per column of the range, with the caller's overrides merged over
# the data frame's column names.
.table_columns_payload <- function(p, df, rng, reg, props) {
  first <- rng[2L] + 1L                    # 1-based data frame column
  n <- rng[4L] - rng[2L] + 1L
  specs <- vector("list", n)
  for (cc in p$columns) {
    q <- unclass(cc)
    idx <- .resolve_col_index(q$col, names(df), "table column")
    k <- idx - first + 1L
    if (length(k) != 1L || k < 1L || k > n)
      stop(sprintf("table column \"%s\" is outside the table's range",
                   as.character(q$col)), call. = FALSE)
    specs[[k]] <- q
  }
  # `[[` throughout, never `$`: these lists have had their NULLs dropped, and
  # `$` partial-matches, so on a column carrying only `total_label` a `$total`
  # would quietly return the label and a `$format` would return the header
  # format.  Both are silent wrong answers rather than errors.
  lapply(seq_len(n), function(k) {
    q <- specs[[k]]
    dfcol <- first + k - 1L
    default_header <- if (dfcol <= length(df)) names(df)[dfcol]
                      else sprintf("Column %d", k)
    hdr <- q[["header"]]
    ent <- list(header = if (!is.null(hdr)) hdr else default_header)
    if (!is.null(q[["formula"]]))     ent$formula <- q[["formula"]]
    if (!is.null(q[["total"]]))
      ent$total_function <- unname(.LXW_TABLE_FUNCTION[[q[["total"]]]])
    if (!is.null(q[["total_label"]])) ent$total_string <- q[["total_label"]]
    if (!is.null(q[["format"]]))
      ent$format_id <- .register_format(reg,
        merge_xl_format(props$default_format, q[["format"]]))
    if (!is.null(q[["header_format"]]))
      ent$header_format_id <- .register_format(reg,
        merge_xl_format(props$default_format, q[["header_format"]]))
    ent
  })
}

# Normalise one column spec or a list of them.
.table_column_list <- function(columns, arg = "columns") {
  if (is.null(columns)) return(list())
  cs <- if (inherits(columns, "xl_table_column")) list(columns) else columns
  if (!is.list(cs))
    stop(sprintf("`%s` must be an xl_table_column object or a list of them",
                 arg), call. = FALSE)
  for (i in seq_along(cs))
    if (!inherits(cs[[i]], "xl_table_column"))
      stop(sprintf("`%s[[%d]]` must be an xl_table_column object", arg, i),
           call. = FALSE)
  cs
}

# Excel's rules for a table name.  libxlsxwriter checks length, a bare C or R,
# and punctuation; it does not reject a cell reference, which Excel does.
#
# The punctuation set is the one _check_table_name() passes to strpbrk() in
# worksheet.c.  Testing membership rather than matching a regex mirrors what
# strpbrk does and keeps the set readable -- as a bracket expression it needs
# escaping that is easy to get wrong ("," "-" "/" reads as a character range).
.TABLE_NAME_BAD_CHARS <- strsplit(
  " !\"#$%&'()*+,-/:;<=>?@[\\]^`{|}~", "", fixed = TRUE)[[1L]]

.check_table_name <- function(name, arg = "name") {
  if (!is.character(name) || length(name) != 1L || !nzchar(name))
    stop(sprintf("`%s` must be a single non-empty string, NULL or NA", arg),
         call. = FALSE)
  if (nchar(name) > 255L)
    stop(sprintf("`%s` exceeds Excel's limit of 255 characters", arg),
         call. = FALSE)
  if (grepl("^[0-9]", name))
    stop(sprintf("`%s` may not start with a digit (got \"%s\")", arg, name),
         call. = FALSE)
  if (toupper(name) %in% c("C", "R"))
    stop(sprintf("`%s` may not be \"C\" or \"R\": Excel reserves them", arg),
         call. = FALSE)
  if (any(strsplit(name, "", fixed = TRUE)[[1L]] %in% .TABLE_NAME_BAD_CHARS))
    stop(sprintf(paste0("`%s` may not contain spaces or punctuation (got ",
                        "\"%s\"); use letters, digits, \".\" or \"_\""),
                 arg, name), call. = FALSE)
  # a name that reads as a cell reference is rejected by Excel but not by
  # libxlsxwriter, so it would reach the file and prompt a repair
  if (grepl("^[A-Za-z]{1,3}[0-9]+$", name))
    stop(sprintf(paste0("`%s` looks like a cell reference (\"%s\"), which ",
                        "Excel will not accept as a table name"), arg, name),
         call. = FALSE)
  invisible(name)
}

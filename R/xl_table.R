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
  cat(sprintf("<xl_table_column: %s%s%s>\n", p$col,
              if (!is.null(p$formula)) " formula" else "",
              if (!is.null(p$total)) paste0(" total=", p$total) else ""))
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
    bad <- vapply(cols, function(cc)
      !is.null(cc$total) || !is.null(cc$total_label), logical(1))
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

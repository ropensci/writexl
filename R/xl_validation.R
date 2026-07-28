# =============================================================================
# Data validation: restricting what may be typed into a range
# =============================================================================
#
# libxlsxwriter exposes 17 LXW_VALIDATION_TYPE_* values, but they are really 5
# kinds crossed with 3 ways of supplying the limit:
#
#            plain value        formula              serial number
#   integer  INTEGER            INTEGER_FORMULA      --
#   decimal  DECIMAL            DECIMAL_FORMULA      --
#   date     DATE               DATE_FORMULA         DATE_NUMBER
#   time     TIME               TIME_FORMULA         TIME_NUMBER
#   length   LENGTH             LENGTH_FORMULA       --
#
# plus LIST, LIST_FORMULA, ANY and CUSTOM_FORMULA.
#
# writexl exposes the five kinds and picks the variant from the R type of the
# limit: a character starting with "=" is a formula, a Date/POSIXct is a
# datetime.  The enum never reaches the user.
#
# libxlsxwriter checks a lot of this itself, but reports it as a warning on
# stderr plus a bare LXW_ERROR_PARAMETER_VALIDATION.  Every rule is therefore
# re-checked here, where the message can name the argument.
# -----------------------------------------------------------------------------

# Enum values, in the header's declaration order.
.LXW_VALIDATION <- c(
  none = 0L, integer = 1L, integer_formula = 2L, decimal = 3L,
  decimal_formula = 4L, list = 5L, list_formula = 6L, date = 7L,
  date_formula = 8L, date_number = 9L, time = 10L, time_formula = 11L,
  time_number = 12L, length = 13L, length_formula = 14L,
  custom_formula = 15L, any = 16L
)

.LXW_VALIDATION_CRITERIA <- c(
  none = 0L, between = 1L, "not between" = 2L, "==" = 3L, "!=" = 4L,
  ">" = 5L, "<" = 6L, ">=" = 7L, "<=" = 8L
)

.LXW_VALIDATION_ERROR_TYPE <- c(stop = 0L, warning = 1L, information = 2L)

# lxw_validation_boolean is three-valued, not a plain flag: 0 leaves
# libxlsxwriter's own default (which is "on" for all four toggles), so a FALSE
# has to be sent as OFF rather than as zero.
.LXW_VALIDATION_BOOL <- c("default" = 0L, off = 1L, on = 2L)
.validation_bool <- function(x) if (isTRUE(x)) 2L else 1L

# The kinds a user names, and whether each has a formula / serial variant.
.VALIDATION_KINDS <- c("integer", "decimal", "date", "time", "length",
                       "list", "custom", "any")

# Excel's own limits, enforced in libxlsxwriter as
# LXW_VALIDATION_MAX_TITLE_LENGTH / MAX_STRING_LENGTH.
.VALIDATION_TITLE_MAX  <- 32L
.VALIDATION_STRING_MAX <- 255L

# Which criteria need two bounds.
.is_between <- function(criteria) criteria %in% c("between", "not between")

# Is this limit an Excel formula rather than a literal?
.is_formula_limit <- function(x) {
  is.character(x) && length(x) == 1L && !is.na(x) && startsWith(x, "=")
}

.is_datetime_limit <- function(x) inherits(x, "Date") || inherits(x, "POSIXct")

# Check a title or message against Excel's length limit, in characters (which
# is what lxw_utf8_strlen() counts).
.check_validation_text <- function(x, arg, max) {
  if (.is_unset(x)) return(NULL)
  s <- .val_str(x, arg)
  if (!is.null(s) && nchar(s, type = "chars") > max)
    stop(sprintf("`%s` must be at most %d characters (got %d)", arg, max,
                 nchar(s, type = "chars")), call. = FALSE)
  s
}

# Resolve the libxlsxwriter type from the kind plus the R type of the limits.
.validation_type <- function(kind, limits) {
  if (kind == "any")    return("any")
  if (kind == "custom") return("custom_formula")
  if (kind == "list")
    return(if (length(limits) && .is_formula_limit(limits[[1L]]))
             "list_formula" else "list")
  supplied <- limits[!vapply(limits, is.null, logical(1))]
  if (length(supplied) && all(vapply(supplied, .is_formula_limit, logical(1))))
    return(paste0(kind, "_formula"))
  # A Date/POSIXct bound on a date/time validation is written as a serial
  # number, which is what the _NUMBER variants are for.
  if (kind %in% c("date", "time") &&
      length(supplied) && all(vapply(supplied, .is_datetime_limit, logical(1))))
    return(kind)
  kind
}

#' Restrict what can be typed into a range
#'
#' @description
#' `xl_validation()` adds Excel data validation to a range: a dropdown list, a
#' numeric or date bound, a text-length limit, or a custom formula.  Pass one or
#' a list of them as `xl_sheet(validation = )`.
#'
#' Excel's 17 internal validation types are collapsed into the five `type`
#' kinds below.  Whether a limit is a literal, a cell formula or a date is
#' inferred from what you pass: a string starting with `"="` is a formula, and a
#' `Date` or `POSIXct` is a date/time bound.
#'
#' @param range The cells to validate: an Excel range string such as
#'   `"B2:B100"`, a single cell, or a `list(rows = , cols = )` spec.
#' @param type The kind of value allowed: `"integer"`, `"decimal"`, `"date"`,
#'   `"time"`, `"length"` (of the text entered), `"custom"` (any formula that
#'   must evaluate `TRUE`), or `"any"` (no restriction, useful when you only
#'   want the input message).  Ignored when `list` is given.
#' @param list A dropdown of allowed values: a character vector of choices, or a
#'   single `"=..."` formula naming a range that holds them.  The choices are
#'   stored joined by commas, and Excel limits that joined string to 255
#'   characters.
#' @param criteria How `value` (or `min`/`max`) limits the entry: `"between"`,
#'   `"not between"`, `"=="`, `"!="`, `">"`, `"<"`, `">="` or `"<="`.  Required
#'   for the numeric, date, time and length kinds; must not be given for
#'   `list`, `custom` or `any`, which carry their own meaning.  Supplying `min`
#'   and `max` implies `"between"`.
#' @param value The single limit the criteria applies to.  A number, a `Date` or
#'   `POSIXct`, or a `"=..."` formula.
#' @param min,max The two limits for `"between"` / `"not between"`.  Supplying
#'   both and omitting `criteria` implies `"between"`.
#' @param input_title,input_message Text shown in a tooltip when the cell is
#'   selected.  Titles are limited to 32 characters and messages to 255.
#' @param error_title,error_message Text shown when an invalid entry is made.
#'   Same limits.
#' @param error_type What Excel does on an invalid entry: `"stop"` (refuse it),
#'   `"warning"` or `"information"` (both allow it through).
#' @param ignore_blank Logical; allow an empty cell.  `TRUE` by default, as in
#'   Excel.
#' @param show_input,show_error Logical; whether the input tooltip and the error
#'   alert are shown at all.  Both `TRUE` by default.
#' @param dropdown Logical; show the in-cell dropdown arrow for a `list`
#'   validation.  `TRUE` by default.
#' @return An `xl_validation` object.
#' @family writexl
#' @seealso [xl_sheet]
#' @export
#' @examples
#' # a dropdown
#' xl_validation("C2:C100", list = c("open", "high", "close"))
#'
#' # a numeric bound, with the message Excel shows on a bad entry
#' xl_validation("B2:B100", type = "integer", min = 1, max = 10,
#'               error_message = "Enter a whole number from 1 to 10")
#'
#' # a date bound
#' xl_validation("D2:D100", type = "date", criteria = ">=",
#'               value = as.Date("2024-01-01"))
#'
#' df <- data.frame(qty = 1:3)
#' tmp <- write_xlsx(list(Data = xl_sheet(df,
#'   validation = xl_validation("A2:A4", type = "integer", min = 0, max = 99))))
xl_validation <- function(range, type = "any", list = NULL, criteria = NA,
                          value = NULL, min = NULL, max = NULL,
                          input_title = NA, input_message = NA,
                          error_title = NA, error_message = NA,
                          error_type = NA, ignore_blank = NA,
                          show_input = NA, show_error = NA, dropdown = NA) {
  if (missing(range) || is.null(range))
    stop("`range` must name the cells to validate", call. = FALSE)
  kind <- .val_enum(type, .VALIDATION_KINDS, "type")
  if (is.null(kind)) kind <- "any"
  if (!is.null(list)) kind <- "list"

  crit <- .val_enum(criteria, names(.LXW_VALIDATION_CRITERIA)[-1L], "criteria")
  # Supplying both bounds is unambiguous, so let it stand in for the criteria.
  if (is.null(crit) && !is.null(min) && !is.null(max)) crit <- "between"

  .check_validation_args(kind, crit, list, value, min, max)

  structure(
    .drop_null(base::list(
      range = range, kind = kind, criteria = crit, list = list,
      value = value, min = min, max = max,
      input_title   = .check_validation_text(input_title, "input_title",
                                             .VALIDATION_TITLE_MAX),
      input_message = .check_validation_text(input_message, "input_message",
                                             .VALIDATION_STRING_MAX),
      error_title   = .check_validation_text(error_title, "error_title",
                                             .VALIDATION_TITLE_MAX),
      error_message = .check_validation_text(error_message, "error_message",
                                             .VALIDATION_STRING_MAX),
      error_type    = .val_enum(error_type, names(.LXW_VALIDATION_ERROR_TYPE),
                                "error_type"),
      ignore_blank  = .val_flag(ignore_blank, "ignore_blank"),
      show_input    = .val_flag(show_input, "show_input"),
      show_error    = .val_flag(show_error, "show_error"),
      dropdown      = .val_flag(dropdown, "dropdown")
    )),
    class = "xl_validation"
  )
}

# The rules libxlsxwriter enforces with a stderr warning and a bare error code.
.check_validation_args <- function(kind, crit, lst, value, min, max) {
  needs_criteria <- kind %in% c("integer", "decimal", "date", "time", "length")

  if (needs_criteria && is.null(crit))
    stop(sprintf(paste0("`criteria` is required for type = \"%s\"; give one of ",
                        "\"between\", \"not between\", \"==\", \"!=\", \">\", ",
                        "\"<\", \">=\", \"<=\", or supply both `min` and `max`"),
                 kind), call. = FALSE)
  if (!needs_criteria && !is.null(crit))
    stop(sprintf(paste0("`criteria` does not apply to type = \"%s\": that kind ",
                        "carries its own meaning"), kind), call. = FALSE)

  if (kind == "list") {
    if (is.null(lst) || !length(lst))
      stop("`list` must give at least one choice", call. = FALSE)
    if (!is.character(lst))
      stop("`list` must be a character vector of choices, or a single \"=\" ",
           "formula naming the range that holds them", call. = FALSE)
    if (anyNA(lst))
      stop("`list` choices must not be NA", call. = FALSE)
    if (!.is_formula_limit(lst[[1L]])) {
      joined <- nchar(paste(lst, collapse = ","), type = "chars")
      if (joined > .VALIDATION_STRING_MAX)
        stop(sprintf(paste0("the `list` choices joined by commas come to %d ",
                            "characters; Excel allows %d. Put them in a range ",
                            "and pass a \"=\" formula instead"),
                     joined, .VALIDATION_STRING_MAX), call. = FALSE)
    }
    return(invisible(NULL))
  }

  if (kind == "custom" && is.null(value))
    stop("type = \"custom\" needs a `value` holding the formula to evaluate",
         call. = FALSE)
  if (kind == "any") return(invisible(NULL))

  if (.is_between(crit %||% "")) {
    if (is.null(min) || is.null(max))
      stop(sprintf("`criteria = \"%s\"` needs both `min` and `max`", crit),
           call. = FALSE)
  } else if (needs_criteria && is.null(value)) {
    stop(sprintf("`criteria = \"%s\"` needs a `value` to compare against", crit),
         call. = FALSE)
  }
  invisible(NULL)
}

#' @export
print.xl_validation <- function(x, ...) {
  p <- unclass(x)
  cat(sprintf("<xl_validation: %s on %s>\n", p$kind,
              if (is.character(p$range)) p$range else "<spec>"))
  invisible(x)
}

# One limit, flattened into the (number, formula, datetime) triple the
# lxw_data_validation struct carries for each of value / minimum / maximum.
.validation_limit <- function(x, prefix) {
  if (is.null(x)) return(base::list())
  if (.is_formula_limit(x))
    return(stats::setNames(base::list(x), paste0(prefix, "_formula")))
  if (.is_datetime_limit(x))
    return(stats::setNames(base::list(.datetime_fields(x)),
                           paste0(prefix, "_datetime")))
  if (is.character(x))
    return(stats::setNames(base::list(x), paste0(prefix, "_formula")))
  stats::setNames(base::list(as.numeric(x)), paste0(prefix, "_number"))
}

# Resolve a sheet's validations to the overlay payloads C applies.
.resolve_validations <- function(el, df, header_offset) {
  if (!inherits(el, "xl_sheet") || is.null(el$validation)) return(list())
  vs <- if (inherits(el$validation, "xl_validation")) list(el$validation)
        else el$validation
  if (!is.list(vs))
    stop("`validation` must be an xl_validation object or a list of them",
         call. = FALSE)
  lapply(seq_along(vs), function(i) {
    v <- vs[[i]]
    if (!inherits(v, "xl_validation"))
      stop(sprintf("`validation[[%d]]` must be an xl_validation object", i),
           call. = FALSE)
    p <- unclass(v)
    arg <- sprintf("validation[[%d]] range", i)
    q <- .xl_resolve_range(p$range, arg = arg, df = df,
                           header_offset = header_offset, allow_cell = TRUE)
    limits <- if (identical(p$kind, "list")) list(p$list[[1L]])
              else list(p$value, p$min, p$max)
    type <- .validation_type(p$kind, limits)

    out <- list(kind = "validation", range = as.integer(q),
                validate = unname(.LXW_VALIDATION[[type]]))
    if (!is.null(p$criteria))
      out$criteria <- unname(.LXW_VALIDATION_CRITERIA[[p$criteria]])
    if (!is.null(p$error_type))
      out$error_type <- unname(.LXW_VALIDATION_ERROR_TYPE[[p$error_type]])
    # A dropdown's choices travel as a character vector; C builds the
    # NULL-terminated array libxlsxwriter turns into a CSV formula.
    if (identical(p$kind, "list")) {
      if (.is_formula_limit(p$list[[1L]])) out$value_formula <- p$list[[1L]]
      else                                 out$value_list <- as.character(p$list)
    }
    if (identical(p$kind, "custom")) out$value_formula <- as.character(p$value)
    else {
      out <- c(out, .validation_limit(p$value, "value"),
               .validation_limit(p$min, "minimum"),
               .validation_limit(p$max, "maximum"))
    }
    for (k in c("input_title", "input_message", "error_title", "error_message"))
      if (!is.null(p[[k]])) out[[k]] <- p[[k]]
    # These three are on by default in libxlsxwriter, so only a FALSE is worth
    # sending; the C side distinguishes absent from 0.
    for (k in c("ignore_blank", "show_input", "show_error", "dropdown"))
      if (!is.null(p[[k]])) out[[k]] <- .validation_bool(p[[k]])
    out
  })
}

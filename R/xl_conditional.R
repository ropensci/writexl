# =============================================================================
# Conditional formatting
# =============================================================================
#
# lxw_conditional_format carries about thirty fields, but they fall into four
# clusters that share almost nothing: the simple rules (type, criteria, value,
# format), colour scales (min/mid/max x value, rule type, colour), data bars
# (a dozen bar_* fields) and icon sets (style, reverse, icons only).  Rather
# than one constructor with thirty arguments, there is one per cluster; each
# returns an xl_conditional carrying its `cluster`, and one overlay kind in C
# applies them all.
#
# The trap here is that the 34 criteria are partitioned by type -- text criteria
# only mean anything for a text rule, the time-period ones only for a time
# rule, and so on -- and libxlsxwriter does not check the pairing.  A mismatch
# produces a rule Excel silently ignores, so .CONDITIONAL_CRITERIA_FOR maps each
# type to what it accepts and the mismatch is an error naming both.
# -----------------------------------------------------------------------------

# Enum values, in the header's declaration order.
.LXW_COND_TYPE <- c(
  none = 0L, cell = 1L, text = 2L, time_period = 3L, average = 4L,
  duplicate = 5L, unique = 6L, top = 7L, bottom = 8L, blanks = 9L,
  no_blanks = 10L, errors = 11L, no_errors = 12L, formula = 13L,
  icon_sets = 14L
)

.LXW_COND_CRITERIA <- c(
  none = 0L,
  "==" = 1L, "!=" = 2L, ">" = 3L, "<" = 4L, ">=" = 5L, "<=" = 6L,
  between = 7L, "not between" = 8L,
  contains = 9L, "not contains" = 10L, "begins with" = 11L, "ends with" = 12L,
  yesterday = 13L, today = 14L, tomorrow = 15L, "last 7 days" = 16L,
  "last week" = 17L, "this week" = 18L, "next week" = 19L,
  "last month" = 20L, "this month" = 21L, "next month" = 22L,
  "above" = 23L, "below" = 24L, "above or equal" = 25L, "below or equal" = 26L,
  "1 std dev above" = 27L, "1 std dev below" = 28L,
  "2 std dev above" = 29L, "2 std dev below" = 30L,
  "3 std dev above" = 31L, "3 std dev below" = 32L,
  "percent" = 33L
)

.LXW_COND_RULE_TYPE <- c(
  none = 0L, minimum = 1L, number = 2L, percent = 3L, percentile = 4L,
  formula = 5L, maximum = 6L, auto_min = 7L, auto_max = 8L
)

.LXW_COND_ICONS <- c(
  "3_arrows"                  = 0L,
  "3_arrows_gray"             = 1L,
  "3_flags"                   = 2L,
  "3_traffic_lights"          = 3L,
  "3_traffic_lights_rimmed"   = 4L,
  "3_signs"                   = 5L,
  "3_symbols_circled"         = 6L,
  "3_symbols"                 = 7L,
  "4_arrows"                  = 8L,
  "4_arrows_gray"             = 9L,
  "4_red_to_black"            = 10L,
  "4_ratings"                 = 11L,
  "4_traffic_lights"          = 12L,
  "5_arrows"                  = 13L,
  "5_arrows_gray"             = 14L,
  "5_ratings"                 = 15L,
  "5_quarters"                = 16L
)

# Which criteria each rule type accepts.  libxlsxwriter checks none of this, and
# a mismatched pair produces a rule Excel quietly ignores.
.COND_COMPARISONS <- c("==", "!=", ">", "<", ">=", "<=", "between",
                       "not between")
.COND_TEXT        <- c("contains", "not contains", "begins with", "ends with")
.COND_TIME        <- c("yesterday", "today", "tomorrow", "last 7 days",
                       "last week", "this week", "next week", "last month",
                       "this month", "next month")
.COND_AVERAGE     <- c("above", "below", "above or equal", "below or equal",
                       "1 std dev above", "1 std dev below",
                       "2 std dev above", "2 std dev below",
                       "3 std dev above", "3 std dev below")

.CONDITIONAL_CRITERIA_FOR <- list(
  cell        = .COND_COMPARISONS,
  text        = .COND_TEXT,
  time_period = .COND_TIME,
  average     = .COND_AVERAGE,
  top         = "percent",          # optional: absent means "top N"
  bottom      = "percent",
  duplicate   = character(0),
  unique      = character(0),
  blanks      = character(0),
  no_blanks   = character(0),
  errors      = character(0),
  no_errors   = character(0),
  formula     = character(0)
)

# The criteria that only make sense with two bounds.
.cond_is_between <- function(x) !is.null(x) && x %in% c("between", "not between")

# Infer the rule type from the criteria when the caller did not name one: the
# two are largely redundant, and a comparison plainly means a cell rule.
.cond_infer_type <- function(type, criteria) {
  if (!is.null(type)) return(type)
  if (is.null(criteria)) return("cell")
  if (criteria %in% .COND_TEXT)    return("text")
  if (criteria %in% .COND_TIME)    return("time_period")
  if (criteria %in% .COND_AVERAGE) return("average")
  "cell"
}

.check_conditional_criteria <- function(type, criteria) {
  allowed <- .CONDITIONAL_CRITERIA_FOR[[type]]
  if (is.null(allowed)) return(invisible(NULL))
  if (is.null(criteria)) {
    # only the comparison and text kinds insist on one
    if (type %in% c("cell", "text", "time_period", "average"))
      stop(sprintf(paste0("`criteria` is required for type = \"%s\"; one of: %s"),
                   type, paste(allowed, collapse = ", ")), call. = FALSE)
    return(invisible(NULL))
  }
  if (!length(allowed))
    stop(sprintf(paste0("`criteria` does not apply to type = \"%s\": that rule ",
                        "carries its own meaning"), type), call. = FALSE)
  if (!criteria %in% allowed)
    stop(sprintf(paste0("`criteria = \"%s\"` is not valid for type = \"%s\". ",
                        "Excel would ignore the rule. Valid here: %s"),
                 criteria, type, paste(allowed, collapse = ", ")),
         call. = FALSE)
  invisible(NULL)
}

# Shared tail of every constructor.
.new_conditional <- function(cluster, range, fields, stop_if_true, multi_range) {
  if (missing(range) || is.null(range))
    stop("`range` must name the cells to format", call. = FALSE)
  structure(
    c(list(cluster = cluster, range = range),
      fields,
      .drop_null(list(stop_if_true = .val_flag(stop_if_true, "stop_if_true"),
                      multi_range  = .val_str(multi_range, "multi_range")))),
    class = "xl_conditional"
  )
}

#' Format cells according to their contents
#'
#' @description
#' Excel's conditional formatting, in four flavours:
#'
#' * `xl_cond_cell()` --- a rule with a format: comparisons, text matches, time
#'   periods, above/below average, top/bottom N, duplicates, blanks, errors, or
#'   an arbitrary formula.
#' * `xl_cond_scale()` --- a two- or three-colour scale across the range.
#' * `xl_cond_bar()` --- in-cell data bars.
#' * `xl_cond_icons()` --- one of Excel's built-in icon sets.
#'
#' Pass one or a list of them as `xl_sheet(conditional = )`.
#'
#' @param range The cells the rule applies to: an Excel range string such as
#'   `"B2:B100"`, a single cell, or a `list(rows = , cols = )` spec.
#' @param criteria How the rule decides, which depends on `type`:
#'   `"=="`, `"!="`, `">"`, `"<"`, `">="`, `"<="`, `"between"`, `"not between"`
#'   for `"cell"`; `"contains"`, `"not contains"`, `"begins with"`,
#'   `"ends with"` for `"text"`; `"yesterday"`, `"today"`, `"tomorrow"`,
#'   `"last 7 days"`, `"last week"`, `"this week"`, `"next week"`,
#'   `"last month"`, `"this month"`, `"next month"` for `"time_period"`;
#'   `"above"`, `"below"`, `"above or equal"`, `"below or equal"` and the
#'   `"N std dev above"` / `"below"` variants for `"average"`; and `"percent"`
#'   for `"top"` / `"bottom"`.
#'
#'   Pairing a criteria with the wrong `type` is an error --- Excel would accept
#'   the file and silently ignore the rule.
#' @param value What the criteria compares against: a number, a string (for the
#'   text criteria), or an `"=..."` formula.  For `"top"` / `"bottom"` it is the
#'   N.  Use `min` and `max` for `"between"` / `"not between"`.
#' @param min,max The two bounds for `"between"` / `"not between"`.
#' @param type The kind of rule, when it cannot be inferred from `criteria`:
#'   `"cell"`, `"text"`, `"time_period"`, `"average"`, `"top"`, `"bottom"`,
#'   `"duplicate"`, `"unique"`, `"blanks"`, `"no_blanks"`, `"errors"`,
#'   `"no_errors"` or `"formula"`.
#' @param format The [xl_format] applied to cells that match.
#' @param stop_if_true Logical; if this rule matches, skip the later rules on
#'   the same cells.
#' @param multi_range A further set of ranges the rule also covers, as an Excel
#'   multi-range string such as `"B3:K6 B9:K12"`.
#' @return An `xl_conditional` object.
#' @family writexl
#' @seealso [xl_sheet], [xl_format]
#' @name xl_conditional
#' @examples
#' xl_cond_cell("B2:B100", criteria = ">", value = 100, format = xl_fill(background = "red"))
#' xl_cond_cell("C2:C100", type = "text", criteria = "contains", value = "urgent",
#'              format = xl_font(bold = TRUE))
#' xl_cond_cell("D2:D100", type = "duplicate",
#'              format = xl_fill(background = "yellow"))
NULL

#' @rdname xl_conditional
#' @export
xl_cond_cell <- function(range, type = NA, criteria = NA, value = NULL,
                         min = NULL, max = NULL, format = NULL,
                         stop_if_true = NA, multi_range = NA) {
  # the range is the subject of the call, so check it before anything else
  if (missing(range) || is.null(range))
    stop("`range` must name the cells to format", call. = FALSE)
  crit <- .val_enum(criteria, names(.LXW_COND_CRITERIA)[-1L], "criteria")
  kind <- .val_enum(type, names(.CONDITIONAL_CRITERIA_FOR), "type")
  kind <- .cond_infer_type(kind, crit)
  .check_conditional_criteria(kind, crit)

  if (.cond_is_between(crit) && (is.null(min) || is.null(max)))
    stop(sprintf("`criteria = \"%s\"` needs both `min` and `max`", crit),
         call. = FALSE)
  if (!is.null(crit) && !.cond_is_between(crit) && is.null(value) &&
      kind %in% c("cell", "text"))
    stop(sprintf("`criteria = \"%s\"` needs a `value` to compare against", crit),
         call. = FALSE)
  if (kind == "formula" && is.null(value))
    stop("type = \"formula\" needs a `value` holding the formula", call. = FALSE)
  if (!is.null(format) && !is_xl_format(format))
    stop("`format` must be an xl_format object", call. = FALSE)

  .new_conditional("cell", range,
                   .drop_null(list(type = kind, criteria = crit, value = value,
                                   min = min, max = max, format = format)),
                   stop_if_true, multi_range)
}

#' @export
print.xl_conditional <- function(x, ...) {
  p <- unclass(x)
  cat(sprintf("<xl_conditional: %s on %s>\n", p$cluster,
              if (is.character(p$range)) p$range else "<spec>"))
  invisible(x)
}

# One limit, flattened into the (number, string) pair the struct carries.  A
# string is either an Excel formula or the literal text a text-rule matches;
# libxlsxwriter takes both through value_string.
.cond_value <- function(x, prefix) {
  if (is.null(x)) return(list())
  if (is.character(x))
    return(stats::setNames(list(as.character(x)), paste0(prefix, "_string")))
  stats::setNames(list(as.numeric(x)), prefix)
}

# Resolve a sheet's conditional formats to the overlay payloads C applies.
.resolve_conditionals <- function(el, df, reg, header_offset, props) {
  if (!inherits(el, "xl_sheet") || is.null(el$conditional)) return(list())
  cs <- if (inherits(el$conditional, "xl_conditional")) list(el$conditional)
        else el$conditional
  if (!is.list(cs))
    stop("`conditional` must be an xl_conditional object or a list of them",
         call. = FALSE)
  lapply(seq_along(cs), function(i) {
    cf <- cs[[i]]
    if (!inherits(cf, "xl_conditional"))
      stop(sprintf("`conditional[[%d]]` must be an xl_conditional object", i),
           call. = FALSE)
    p <- unclass(cf)
    q <- .xl_resolve_range(p$range, arg = sprintf("conditional[[%d]] range", i),
                           df = df, header_offset = header_offset,
                           allow_cell = TRUE)
    out <- list(kind = "conditional", range = as.integer(q))

    if (identical(p$cluster, "cell")) {
      out$type <- unname(.LXW_COND_TYPE[[p$type]])
      if (!is.null(p$criteria))
        out$criteria <- unname(.LXW_COND_CRITERIA[[p$criteria]])
      out <- c(out, .cond_value(p$value, "value"),
               .cond_value(p$min, "min_value"),
               .cond_value(p$max, "max_value"))
      # the format is an ordinary xl_format, so it goes through the same
      # registry as every other format -- libxlsxwriter emits it as a
      # differential format (<dxf>) rather than a cell style
      out$format_id <- .register_format(reg,
                                        merge_xl_format(props$default_format,
                                                        p$format))
    }
    for (k in c("stop_if_true"))
      if (!is.null(p[[k]])) out[[k]] <- as.integer(isTRUE(p[[k]]))
    if (!is.null(p$multi_range)) out$multi_range <- p$multi_range
    out
  })
}

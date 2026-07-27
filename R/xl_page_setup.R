# =============================================================================
# Page setup: how a worksheet prints
# =============================================================================
#
# xl_page_setup() collects everything Excel's "Page Layout" ribbon controls:
# orientation, paper size, margins, scaling, centring, the print options, and
# the header/footer.  None of it affects the cells -- it only affects printing
# and the page-break preview -- so it is a plain bag of scalars that flattens
# into the per-sheet plan and is applied before any row is written.
# -----------------------------------------------------------------------------

# Excel's paper types, from the table in worksheet.h.  Only the sizes worth
# naming are here; the ~40 envelope and specialist types are reachable by
# passing the index directly.
.LXW_PAPER <- c(
  "default"   = 0L,
  "letter"    = 1L,
  "tabloid"   = 3L,
  "ledger"    = 4L,
  "legal"     = 5L,
  "statement" = 6L,
  "executive" = 7L,
  "A3"        = 8L,
  "A4"        = 9L,
  "A5"        = 11L,
  "B4"        = 12L,
  "B5"        = 13L,
  "folio"     = 14L,
  "quarto"    = 15L
)

# libxlsxwriter caps a header/footer string at 255 characters
# (LXW_HEADER_FOOTER_MAX in worksheet.h).
.HEADER_FOOTER_MAX <- 255L

# Resolve a paper size to Excel's numeric type.  A name is looked up
# case-insensitively; an integer passes through so the unnamed sizes stay
# reachable.
.resolve_paper <- function(x) {
  if (.is_unset(x)) return(NULL)
  if (is.character(x) && length(x) == 1L) {
    i <- match(tolower(x), tolower(names(.LXW_PAPER)))
    if (is.na(i))
      stop("`paper` must be one of ", paste(names(.LXW_PAPER), collapse = ", "),
           ", or an Excel paper-type integer", call. = FALSE)
    return(unname(.LXW_PAPER[[i]]))
  }
  # libxlsxwriter's own valid range; outside it worksheet_set_paper() warns and
  # ignores the value, which would be a silent no-op from R
  .val_int(x, "paper", min = 0, max = 118)
}

# Validate a header or footer string.  Excel's &-codes are passed through
# untouched; the length and the image placeholders are our business.
.check_header_footer <- function(x, arg) {
  if (.is_unset(x)) return(NULL)
  s <- .val_str(x, arg)
  if (is.null(s)) return(NULL)
  # libxlsxwriter measures this in UTF-8 characters, as nchar(type = "chars")
  # does, and errors above the limit
  if (nchar(s, type = "chars") > .HEADER_FOOTER_MAX)
    stop(sprintf("`%s` must be at most %d characters (got %d)", arg,
                 .HEADER_FOOTER_MAX, nchar(s, type = "chars")), call. = FALSE)
  # An &G / &[Picture] placeholder needs a matching image, which writexl cannot
  # supply yet.  libxlsxwriter would reject it with an opaque parameter-
  # validation error, so say what is actually wrong.
  if (grepl("&G", s, fixed = TRUE) || grepl("&[Picture]", s, fixed = TRUE))
    stop(sprintf(paste0("`%s` contains an image placeholder (&G or &[Picture]),",
                        " but header and footer images are not supported yet"),
                 arg), call. = FALSE)
  s
}

# Validate a header/footer margin.  libxlsxwriter applies it only when it is
# greater than zero (`if (options->margin > 0.0)`), so a zero would be silently
# ignored rather than doing what it says.
.check_hf_margin <- function(x, arg) {
  v <- .val_num(x, arg, min = 0)
  if (!is.null(v) && v <= 0)
    stop(sprintf("`%s` must be greater than 0 (Excel's default is 0.3)", arg),
         call. = FALSE)
  v
}

# Normalise a margins argument to the four named sides.
.resolve_margins <- function(x) {
  if (is.null(x) || (length(x) == 1L && is.na(x))) return(NULL)
  sides <- c("left", "right", "top", "bottom")
  if (!is.numeric(x))
    stop("`margins` must be a numeric vector of inches", call. = FALSE)
  if (is.null(names(x))) {
    if (length(x) == 1L) x <- stats::setNames(rep(x, 4L), sides)
    else if (length(x) == 4L) x <- stats::setNames(x, sides)
    else stop("`margins` must be one value, four values, or a named vector ",
              "with any of: ", paste(sides, collapse = ", "), call. = FALSE)
  }
  bad <- setdiff(names(x), sides)
  if (length(bad))
    stop("unknown `margins` element(s): ", paste(bad, collapse = ", "),
         call. = FALSE)
  if (anyNA(x) || any(x < 0))
    stop("`margins` must be non-negative", call. = FALSE)
  as.list(x)
}

# Normalise fit_to to c(width, height); 0 means "unconstrained in that
# direction", which is how Excel spells "fit to 1 page wide, any number tall".
.resolve_fit_to <- function(x) {
  if (is.null(x) || (length(x) == 1L && is.na(x))) return(NULL)
  if (!is.numeric(x) || !length(x) || length(x) > 2L || anyNA(x) || any(x < 0))
    stop("`fit_to` must be one or two non-negative numbers, c(width, height)",
         call. = FALSE)
  if (!is.null(names(x))) {
    bad <- setdiff(names(x), c("width", "height"))
    if (length(bad))
      stop("unknown `fit_to` element(s): ", paste(bad, collapse = ", "),
           call. = FALSE)
    w <- if (is.na(x["width"])) 1L else as.integer(x[["width"]])
    h <- if (is.na(x["height"])) 1L else as.integer(x[["height"]])
    return(c(width = w, height = h))
  }
  if (length(x) == 1L) x <- c(x, 1)
  c(width = as.integer(x[1L]), height = as.integer(x[2L]))
}

#' How a worksheet prints
#'
#' @description
#' `xl_page_setup()` collects Excel's page-layout settings for one worksheet:
#' orientation, paper size, margins, scaling, centring, the print options, and
#' the header and footer.  Pass it as `xl_sheet(page = )`.
#'
#' None of these affect the cell data --- they change only how the sheet prints
#' and how it looks in Excel's page-break preview.
#'
#' @param orientation `"portrait"` or `"landscape"`.
#' @param paper Paper size: a name (`"A4"`, `"letter"`, `"legal"`, `"tabloid"`,
#'   `"ledger"`, `"statement"`, `"executive"`, `"A3"`, `"A5"`, `"B4"`, `"B5"`,
#'   `"folio"`, `"quarto"`, `"default"`), or an Excel paper-type integer for the
#'   envelope and specialist sizes that have no name here.
#' @param margins Page margins in inches: a single number for all four sides,
#'   four numbers in the order left, right, top, bottom, or a named vector using
#'   any of `left`, `right`, `top`, `bottom`.
#' @param scale Print scaling as a percentage (10--400).  Ignored by Excel when
#'   `fit_to` is set.
#' @param fit_to Fit the printout to `c(width, height)` pages.  A `0` means
#'   "as many pages as needed in that direction", so `c(width = 1, height = 0)`
#'   is Excel's "fit all columns on one page".
#' @param center_horizontally,center_vertically Logical; centre the printed
#'   output on the page.
#' @param header,footer Header and footer text, at most 255 characters, using
#'   Excel's own codes: `&L`, `&C`, `&R` start the left, centre and right
#'   sections, `&P` is the page number, `&N` the page count, `&D` the date, `&A`
#'   the sheet name, and `&&` a literal ampersand.  For example
#'   `"&LQ1 report&RPage &P of &N"`.  Image placeholders (`&G` /
#'   `&[Picture]`) are not supported yet and are rejected.
#' @param header_margin,footer_margin Header/footer margin in inches (Excel's
#'   default is 0.3).  Must be greater than 0.
#' @param page_view Logical; open the sheet in Excel's page-layout view rather
#'   than normal view.
#' @param first_page The page number to start numbering from.
#' @param across Logical; print pages left-to-right before top-to-bottom
#'   (Excel's "over, then down").
#' @param black_and_white Logical; print without colour.
#' @param row_col_headers Logical; print the row numbers and column letters.
#' @return An `xl_page_setup` object.
#' @family writexl
#' @seealso [xl_sheet], [write_xlsx]
#' @export
#' @examples
#' xl_page_setup(orientation = "landscape", paper = "A4",
#'               fit_to = c(width = 1, height = 0))
#'
#' # margins in inches, and a header with page numbers
#' xl_page_setup(margins = c(left = 1, right = 1),
#'               header = "&LQuarterly report&RPage &P of &N")
#'
#' df <- data.frame(x = 1:3)
#' tmp <- write_xlsx(list(Data = xl_sheet(df, page = xl_page_setup(
#'   orientation = "landscape", header = "&CDraft"
#' ))))
xl_page_setup <- function(orientation = NA, paper = NA, margins = NULL,
                          scale = NA, fit_to = NULL,
                          center_horizontally = NA, center_vertically = NA,
                          header = NA, footer = NA,
                          header_margin = NA, footer_margin = NA,
                          page_view = NA, first_page = NA, across = NA,
                          black_and_white = NA, row_col_headers = NA) {
  out <- .drop_null(list(
    orientation         = .val_enum(orientation, c("portrait", "landscape"),
                                    "orientation"),
    paper               = .resolve_paper(paper),
    margins             = .resolve_margins(margins),
    scale               = .val_int(scale, "scale", min = 10, max = 400),
    fit_to              = .resolve_fit_to(fit_to),
    center_horizontally = .val_flag(center_horizontally, "center_horizontally"),
    center_vertically   = .val_flag(center_vertically, "center_vertically"),
    header              = .check_header_footer(header, "header"),
    footer              = .check_header_footer(footer, "footer"),
    header_margin       = .check_hf_margin(header_margin, "header_margin"),
    footer_margin       = .check_hf_margin(footer_margin, "footer_margin"),
    page_view           = .val_flag(page_view, "page_view"),
    first_page          = .val_int(first_page, "first_page", min = 0),
    across              = .val_flag(across, "across"),
    black_and_white     = .val_flag(black_and_white, "black_and_white"),
    row_col_headers     = .val_flag(row_col_headers, "row_col_headers")
  ))
  structure(out, class = "xl_page_setup")
}

#' @export
print.xl_page_setup <- function(x, ...) {
  p <- unclass(x)
  cat(sprintf("<xl_page_setup: %d setting%s>\n", length(p),
              if (length(p) == 1L) "" else "s"))
  for (k in names(p))
    cat(sprintf("  %s: %s\n", k, paste(format(unlist(p[[k]])), collapse = ", ")))
  invisible(x)
}

# Flatten a page setup into the flat scalars C applies.  Margins and fit_to are
# spread into individual keys so the C side only ever reads scalars.
.page_setup_payload <- function(page) {
  if (is.null(page)) return(NULL)
  if (!inherits(page, "xl_page_setup"))
    stop("`page` must be an xl_page_setup object", call. = FALSE)
  p <- unclass(page)
  out <- list()
  if (!is.null(p$orientation))
    out$landscape <- as.integer(identical(p$orientation, "landscape"))
  if (!is.null(p$paper))  out$paper <- as.integer(p$paper)
  if (!is.null(p$scale))  out$scale <- as.integer(p$scale)
  if (!is.null(p$margins))
    for (side in names(p$margins))
      out[[paste0("margin_", side)]] <- as.numeric(p$margins[[side]])
  if (!is.null(p$fit_to)) {
    out$fit_width  <- as.integer(p$fit_to[["width"]])
    out$fit_height <- as.integer(p$fit_to[["height"]])
  }
  for (k in c("center_horizontally", "center_vertically", "page_view",
              "across", "black_and_white", "row_col_headers"))
    if (!is.null(p[[k]])) out[[k]] <- as.integer(isTRUE(p[[k]]))
  if (!is.null(p$first_page)) out$first_page <- as.integer(p$first_page)
  if (!is.null(p$header)) out$header <- p$header
  if (!is.null(p$footer)) out$footer <- p$footer
  if (!is.null(p$header_margin)) out$header_margin <- as.numeric(p$header_margin)
  if (!is.null(p$footer_margin)) out$footer_margin <- as.numeric(p$footer_margin)
  if (length(out)) out else NULL
}

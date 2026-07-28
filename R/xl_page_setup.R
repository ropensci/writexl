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

# Excel's limit on page breaks per worksheet (LXW_BREAKS_MAX in worksheet.h).
.BREAKS_MAX <- 1023L

# Resolve `repeat_rows` / `repeat_cols` to the 0-based c(first, last) pair
# libxlsxwriter wants.
#
# These are *sheet* rows and columns, counted from 1 with the header row
# included -- unlike xl_row_spec(), which indexes data rows.  That is
# deliberate: repeating the header row on every printed page is the whole point
# of the feature, and data-row numbering could not name it.
.resolve_repeat <- function(x, arg, kind) {
  if (.is_unset(x)) return(NULL)
  if (is.numeric(x) && length(x) == 1L) {
    n <- .val_int(x, arg, min = 1)
    return(c(0L, as.integer(n) - 1L))
  }
  if (is.character(x) && length(x) == 1L) {
    ok <- if (kind == "row") grepl("^[$]?[0-9]+:[$]?[0-9]+$", x)
          else               grepl("^[$]?[A-Za-z]+:[$]?[A-Za-z]+$", x)
    if (!ok)
      stop(sprintf('`%s` must be a count, or a %s range like "%s"', arg,
                   if (kind == "row") "row" else "column",
                   if (kind == "row") "1:2" else "A:B"), call. = FALSE)
    q <- .xl_resolve_range(x, arg = arg, df = NULL, allow_cell = FALSE)
    return(if (kind == "row") c(q[1L], q[3L]) else c(q[2L], q[4L]))
  }
  stop(sprintf("`%s` must be a single count or range string", arg),
       call. = FALSE)
}

# Resolve a page-break vector to the 0-based positions libxlsxwriter wants.
#
# A break "at" row N means the new page starts at row N, which is 0-based N-1.
# Row/column 1 is therefore rejected: 0 is the array terminator in
# libxlsxwriter, so a break there would silently truncate the whole list rather
# than doing nothing visible.
.resolve_breaks <- function(x, arg, kind) {
  if (is.null(x) || (length(x) == 1L && !is.list(x) && is.na(x))) return(NULL)
  if (is.character(x) && kind == "col") x <- .col_letters_to_index(x) + 1L
  if (!is.numeric(x) || !length(x) || anyNA(x))
    stop(sprintf("`%s` must be a numeric vector of 1-based %s positions", arg,
                 if (kind == "row") "row" else "column"), call. = FALSE)
  v <- sort(unique(as.integer(x)))
  if (any(v < 2L))
    stop(sprintf(paste0("`%s` positions must be 2 or greater: a break at %s 1 ",
                        "would come before any content, and libxlsxwriter uses ",
                        "0 to terminate the break list"), arg,
                 if (kind == "row") "row" else "column"), call. = FALSE)
  if (length(v) > .BREAKS_MAX)
    stop(sprintf("`%s` may have at most %d breaks (got %d)", arg,
                 .BREAKS_MAX, length(v)), call. = FALSE)
  v - 1L
}

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
  s
}

# How many image placeholders a header/footer string carries.  Excel spells the
# same thing two ways, and libxlsxwriter counts both.
.hf_placeholders <- function(s) {
  if (is.null(s)) return(0L)
  n <- function(pat) length(gregexpr(pat, s, fixed = TRUE)[[1L]][
    gregexpr(pat, s, fixed = TRUE)[[1L]] > 0L])
  n("&G") + n("&[Picture]")
}

# The three positions a header/footer image can occupy.
.HF_POSITIONS <- c("left", "center", "right")

# Resolve a header/footer image argument to file paths.
#
# libxlsxwriter takes only filenames here -- lxw_header_footer_options has no
# buffer form (upstream issue 384) -- so an in-memory image is written to a
# temporary file, which lives as long as the session's tempdir.
.resolve_hf_images <- function(x, arg) {
  if (is.null(x) || (length(x) == 1L && !is.list(x) && is.na(x)))
    return(NULL)
  if (!is.list(x)) x <- as.list(x)
  nms <- names(x)
  if (is.null(nms) || any(!nzchar(nms)))
    stop(sprintf(paste0("`%s` must be named by position, e.g. ",
                        "list(left = \"logo.png\")"), arg), call. = FALSE)
  bad <- setdiff(nms, .HF_POSITIONS)
  if (length(bad))
    stop(sprintf("`%s`: unknown position(s): %s. Use %s.", arg,
                 paste(bad, collapse = ", "),
                 paste(.HF_POSITIONS, collapse = ", ")), call. = FALSE)
  if (anyDuplicated(nms))
    stop(sprintf("`%s`: each position may be given only once", arg),
         call. = FALSE)

  out <- list()
  for (pos in nms) {
    img <- x[[pos]]
    where <- sprintf("%s$%s", arg, pos)
    if (.is_raster_like(img)) img <- .raster_to_png(img, where)
    .check_image(img, where)
    out[[pos]] <- if (is.raw(img)) {
      p <- tempfile(fileext = ".img")
      writeBin(img, p)
      p
    } else {
      img
    }
  }
  out
}

# libxlsxwriter requires the placeholder count to equal the image count, and
# rejects a mismatch with a parameter-validation error that says nothing about
# which side is short.  Checking here says which.
.check_hf_images <- function(text, images, text_arg, image_arg) {
  n_ph <- .hf_placeholders(text)
  n_im <- length(images)
  if (n_ph == n_im) return(invisible(NULL))
  if (n_im > n_ph)
    stop(sprintf(paste0("`%s` gives %d image(s) but `%s` has %d image ",
                        "placeholder(s); add &G where each image should sit, ",
                        "e.g. \"&L&G\" for a left image"),
                 image_arg, n_im, text_arg, n_ph), call. = FALSE)
  stop(sprintf(paste0("`%s` has %d image placeholder(s) (&G or &[Picture]) but ",
                      "`%s` gives %d image(s)"),
               text_arg, n_ph, image_arg, n_im), call. = FALSE)
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
#' @param header_image,footer_image Images to place in the header or footer,
#'   named by position: `list(left = , center = , right = )`.  Each may be a
#'   file path, a raw vector, or an in-memory image, exactly as [xl_image()]
#'   accepts.  Each image needs a matching `&G` placeholder in the
#'   corresponding section of `header`/`footer` --- `"&L&G"` puts one on the
#'   left --- and the counts must agree.
#' @param header_margin,footer_margin Header/footer margin in inches (Excel's
#'   default is 0.3).  Must be greater than 0.
#' @param page_view Logical; open the sheet in Excel's page-layout view rather
#'   than normal view.
#' @param first_page The page number to start numbering from.
#' @param across Logical; print pages left-to-right before top-to-bottom
#'   (Excel's "over, then down").
#' @param black_and_white Logical; print without colour.
#' @param row_col_headers Logical; print the row numbers and column letters.
#' @param print_area The range to print: an Excel range string such as
#'   `"A1:F50"`, or a `list(rows = , cols = )` spec naming data rows and
#'   columns, as elsewhere in writexl.
#' @param repeat_rows,repeat_cols Rows/columns to repeat at the top or left of
#'   every printed page.  Either a count (`repeat_rows = 1` repeats the first
#'   sheet row, which is the header when `col_names = TRUE`) or a range string
#'   (`"1:2"`, `"A:B"`).
#'
#'   Note these count *sheet* rows from 1 with the header included, unlike
#'   [xl_row_spec()], which indexes data rows --- repeating the header row is
#'   the usual reason to use this, and data-row numbering could not name it.
#' @param h_breaks,v_breaks Manual page breaks: 1-based sheet positions at which
#'   a new page starts, so `h_breaks = 21` breaks between rows 20 and 21.
#'   `v_breaks` also accepts column letters.  Positions must be 2 or greater,
#'   and at most 1023 breaks are allowed.  Excel ignores manual breaks when
#'   `fit_to` is set, which warns.
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
                          header_image = NULL, footer_image = NULL,
                          page_view = NA, first_page = NA, across = NA,
                          black_and_white = NA, row_col_headers = NA,
                          print_area = NULL, repeat_rows = NA,
                          repeat_cols = NA, h_breaks = NULL, v_breaks = NULL) {
  hdr_images <- .resolve_hf_images(header_image, "header_image")
  ftr_images <- .resolve_hf_images(footer_image, "footer_image")
  .check_hf_images(.check_header_footer(header, "header"), hdr_images,
                   "header", "header_image")
  .check_hf_images(.check_header_footer(footer, "footer"), ftr_images,
                   "footer", "footer_image")
  if (!is.null(h_breaks) || !is.null(v_breaks)) {
    fit <- .resolve_fit_to(fit_to)
    if (!is.null(fit))
      warning("Excel ignores manual page breaks when `fit_to` is set",
              call. = FALSE)
  }
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
    header_image        = hdr_images,
    footer_image        = ftr_images,
    page_view           = .val_flag(page_view, "page_view"),
    first_page          = .val_int(first_page, "first_page", min = 0),
    across              = .val_flag(across, "across"),
    black_and_white     = .val_flag(black_and_white, "black_and_white"),
    row_col_headers     = .val_flag(row_col_headers, "row_col_headers"),
    # the print area is resolved later, when the sheet's data frame is known
    print_area          = print_area,
    repeat_rows         = .resolve_repeat(repeat_rows, "repeat_rows", "row"),
    repeat_cols         = .resolve_repeat(repeat_cols, "repeat_cols", "col"),
    h_breaks            = .resolve_breaks(h_breaks, "h_breaks", "row"),
    v_breaks            = .resolve_breaks(v_breaks, "v_breaks", "col")
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
.page_setup_payload <- function(page, df = NULL, header_offset = 1L) {
  if (is.null(page)) return(NULL)
  if (!inherits(page, "xl_page_setup"))
    stop("`page` must be an xl_page_setup object", call. = FALSE)
  p <- unclass(page)
  out <- list()
  # The print area is the one setting that needs the sheet, so it is resolved
  # here rather than in the constructor.
  if (!is.null(p$print_area))
    out$print_area <- .xl_resolve_range(p$print_area, arg = "print_area",
                                        df = df, header_offset = header_offset,
                                        allow_cell = FALSE)
  for (k in c("repeat_rows", "repeat_cols", "h_breaks", "v_breaks"))
    if (!is.null(p[[k]])) out[[k]] <- as.integer(p[[k]])
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
  for (side in c("header", "footer"))
    for (pos in .HF_POSITIONS) {
      f <- p[[paste0(side, "_image")]][[pos]]
      if (!is.null(f))
        out[[sprintf("%s_image_%s", side, pos)]] <- f
    }
  if (length(out)) out else NULL
}

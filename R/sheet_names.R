# =============================================================================
# Sheet names: fix them up so Excel will open the file
# =============================================================================
#
# Excel's rules, mirrored from workbook_validate_sheet_name() in
# src/libxlsxwriter/src/workbook.c: a name must be non-empty, at most 31
# characters, must not contain any of [ ] : * ? / \, and must not start or end
# with an apostrophe.  Duplicates are also rejected.
#
# writexl fixes names up and warns rather than erroring, which is how it has
# always treated over-long and duplicate names -- a sheet name is often derived
# from data, and failing the whole export over one stray "/" would be unkind.
# The warnings name the original and the replacement so a rename is never
# silent.
#
# The stages run in this order, because each one has to see the previous one's
# output: sanitize -> truncate -> deduplicate.  Deduplicating last matters:
# "a/b" and "a-b" both sanitize to "a-b", and only a final dedup pass catches
# the collision that sanitizing created.
#
# An empty name is left alone rather than repaired.  writexl uses "" to mean
# "let libxlsxwriter name this sheet", and C passes NULL for it.
# -----------------------------------------------------------------------------

.SHEETNAME_MAX <- 31L

# The characters Excel forbids in a sheet name.
.SHEETNAME_BAD_CHARS <- c("[", "]", ":", "*", "?", "/", "\\")

# Report a set of renames as one warning, listing at most `max_show` of them.
.warn_renamed <- function(what, orig, new, changed, max_show = 5L) {
  shown <- utils::head(changed, max_show)
  detail <- paste(sprintf('"%s" -> "%s"', orig[shown], new[shown]),
                  collapse = ", ")
  if (length(changed) > max_show)
    detail <- paste0(detail, sprintf(", and %d more", length(changed) - max_show))
  warning(what, ": ", detail, call. = FALSE)
}

# Replace the characters Excel forbids, and strip apostrophes from the ends.
.sanitize_sheet_names <- function(x) {
  out <- x
  for (ch in .SHEETNAME_BAD_CHARS)
    out <- gsub(ch, "-", out, fixed = TRUE)
  out <- sub("^'+", "", out)
  out <- sub("'+$", "", out)
  changed <- which(out != x & nzchar(x))
  if (length(changed))
    .warn_renamed("Replacing characters Excel does not allow in sheet name(s)",
                  x, out, changed)
  out
}

# Cut to Excel's 31-character maximum.
.truncate_sheet_names <- function(x) {
  out <- x
  long <- nzchar(x) & nchar(x, type = "chars") > .SHEETNAME_MAX
  out[long] <- substr(x[long], 1L, .SHEETNAME_MAX)
  if (any(long))
    .warn_renamed("Truncating sheet name(s) to 31 characters", x, out,
                  which(long))
  out
}

# Make names unique, keeping every result within the 31-character maximum by
# shortening the base rather than letting the suffix push it over.
.dedupe_sheet_names <- function(x) {
  out <- x
  seen <- character(0)
  for (i in seq_along(x)) {
    nm <- x[i]
    # empty means "libxlsxwriter names this one"; two empties do not collide
    if (!nzchar(nm)) next
    cand <- nm
    k <- 0L
    while (cand %in% seen) {
      k <- k + 1L
      sfx <- paste0("_", k)
      cand <- paste0(substr(nm, 1L, .SHEETNAME_MAX - nchar(sfx, type = "chars")),
                     sfx)
    }
    seen <- c(seen, cand)
    out[i] <- cand
  }
  changed <- which(out != x)
  if (length(changed))
    .warn_renamed("Deduplicating sheet names", x, out, changed)
  out
}

# Resolve the user's sheet names to ones Excel will accept.
.resolve_sheet_names <- function(nms, n) {
  if (is.null(nms)) return(rep("", n))
  nms <- as.character(nms)
  nms[is.na(nms)] <- ""
  .dedupe_sheet_names(.truncate_sheet_names(.sanitize_sheet_names(nms)))
}

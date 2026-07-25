# Test helpers: write an xlsx and read back parts of its XML for inspection.

# Unzip a written workbook and return the parsed/raw contents of one part.
xlsx_part <- function(path, part = "xl/styles.xml", raw = FALSE) {
  d <- tempfile("xlsx_")
  dir.create(d)
  utils::unzip(path, exdir = d)
  f <- file.path(d, part)
  if (!file.exists(f)) return(if (raw) "" else NULL)
  if (raw) paste(readLines(f, warn = FALSE), collapse = "\n") else xml2::read_xml(f)
}

# Convenience: write a data frame (or workbook) and return styles.xml as a
# single string for token matching.
styles_string <- function(x, ...) {
  tmp <- tempfile(fileext = ".xlsx")
  write_xlsx(x, tmp, ...)
  xlsx_part(tmp, "xl/styles.xml", raw = TRUE)
}

# Write and return the path (for readback / multi-part inspection).
write_tmp <- function(x, ...) {
  tmp <- tempfile(fileext = ".xlsx")
  write_xlsx(x, tmp, ...)
  tmp
}

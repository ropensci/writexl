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

# A minimal valid 1x1 PNG.  Hardcoded rather than shipped as a fixture or drawn
# with png(), which needs a working graphics device that CRAN machines may not
# have.
PNG_1x1 <- as.raw(c(
  0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
  0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
  0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00,
  0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
  0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82))

png_file <- function() {
  p <- tempfile(fileext = ".png")
  writeBin(PNG_1x1, p)
  p
}

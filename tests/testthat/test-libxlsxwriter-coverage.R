# Coverage of the bundled libxlsxwriter's option structs.
#
# A caller reaches most of libxlsxwriter through a handful of option structs:
# fill one in, hand it over, and every field left at zero is a feature writexl
# does not offer.  Nothing fails when a field is missed -- the file is written,
# Excel opens it, and the option simply is not there -- so the gap is invisible
# until someone reads the header.
#
# The chart gates elsewhere read the same header for *function* coverage.  This
# one reads it for *field* coverage, which is where the misses actually were: a
# chart font's typeface and rotation, a chartsheet's two protection options, an
# outline group's collapsed state, a table column's cached total.  Each was a
# documented field nothing in writexl ever assigned.
#
# A field is "reached" when src/write_xlsx.c assigns it.  UNUSED lists the ones
# that are deliberately not, each with the reason, so adding a field upstream
# fails here rather than passing unnoticed.
#
# This is a floor, not a proof: the check is textual, so a field whose name
# another struct also carries -- `name`, `color`, `size` -- reads as reached
# once any of them is set.  Telling them apart needs the type of the variable
# being assigned, which needs a C parser.  Removing each of the five fields
# above and re-running catches four; the fifth, lxw_chart_font.name, is pinned
# instead by test-xl_chart_format.R asserting typeface= in the chart XML.

lxw_header <- function(name) {
  for (p in c(file.path("../..", "src/libxlsxwriter/include/xlsxwriter", name),
              file.path("../../..", "src/libxlsxwriter/include/xlsxwriter", name)))
    if (file.exists(p)) return(paste(readLines(p, warn = FALSE), collapse = "\n"))
  NULL
}

# The documented fields of one option struct.  Undocumented members (no /** */)
# are libxlsxwriter's own bookkeeping and are not part of the API.
struct_fields <- function(header, struct) {
  txt <- lxw_header(header)
  if (is.null(txt)) return(NULL)
  m <- regmatches(txt, regexpr(sprintf("typedef struct %s \\{.*?\\n\\}", struct),
                               txt))
  if (!length(m)) return(NA_character_)
  documented <- regmatches(m, gregexpr("/\\*\\*.*?\\*/[^;]*;", m))[[1]]
  decl <- sub(".*\\*/\\s*", "", documented)
  trimws(gsub("^.*?([A-Za-z_][A-Za-z0-9_]*)\\s*(\\[[^]]*\\])?\\s*;$", "\\1",
              gsub("\n", " ", decl)))
}

# Fields writexl does not set, and why.  Anything not listed here must be
# assigned in src/write_xlsx.c.
UNUSED <- list(
  lxw_doc_properties = character(0),
  # writing to a buffer instead of a path is not write_xlsx()'s contract
  lxw_workbook_options = c("output_buffer", "output_buffer_size"),
  lxw_protection = character(0),
  lxw_image_options = character(0),
  lxw_chart_options = character(0),
  lxw_row_col_options = character(0),
  lxw_header_footer_options = character(0),
  lxw_data_validation = character(0),
  # the Excel 2010 data bar extension: solid bars and the negative-value axis
  lxw_conditional_format = c("data_bar_2010", "bar_negative_color_same",
                             "bar_negative_border_color_same"),
  lxw_table_options = character(0),
  lxw_table_column = character(0),
  lxw_comment_options = character(0),
  lxw_filter_rule = character(0),
  lxw_rich_string_tuple = character(0),
  # "rarely required, set to 0" in chart.h: a pitch family and character set
  # are Windows font-matching hints, and baseline is super/subscript, which a
  # chart title has no use for
  lxw_chart_font = c("pitch_family", "charset", "baseline"),
  lxw_chart_line = character(0),
  lxw_chart_fill = character(0),
  lxw_chart_pattern = character(0),
  lxw_chart_point = character(0),
  lxw_chart_layout = character(0)
)

STRUCT_HEADER <- c(
  lxw_doc_properties = "workbook.h", lxw_workbook_options = "workbook.h",
  lxw_protection = "worksheet.h", lxw_image_options = "worksheet.h",
  lxw_chart_options = "worksheet.h", lxw_row_col_options = "worksheet.h",
  lxw_header_footer_options = "worksheet.h",
  lxw_data_validation = "worksheet.h", lxw_conditional_format = "worksheet.h",
  lxw_table_options = "worksheet.h", lxw_table_column = "worksheet.h",
  lxw_comment_options = "worksheet.h", lxw_filter_rule = "worksheet.h",
  lxw_rich_string_tuple = "worksheet.h", lxw_chart_font = "chart.h",
  lxw_chart_line = "chart.h", lxw_chart_fill = "chart.h",
  lxw_chart_pattern = "chart.h", lxw_chart_point = "chart.h",
  lxw_chart_layout = "chart.h"
)

writexl_c <- function() {
  for (p in c("../../src/write_xlsx.c", "../../../src/write_xlsx.c"))
    if (file.exists(p)) return(paste(readLines(p, warn = FALSE), collapse = "\n"))
  NULL
}

test_that("every documented option-struct field is either set or listed unused", {
  skip_if(is.null(writexl_c()), "source tree not available")
  src <- writexl_c()
  missed <- character(0)
  for (struct in names(UNUSED)) {
    fields <- struct_fields(STRUCT_HEADER[[struct]], struct)
    expect_false(identical(fields, NA_character_),
                 info = paste("struct not found in the header:", struct))
    if (is.null(fields)) next
    for (f in setdiff(fields, UNUSED[[struct]])) {
      # reached by assignment (`x.f =`, `x->f =`) or by address (`&x.f`, which
      # is how the datetime fields and the metadata strings are filled).  Not
      # by the field's name appearing as a payload key: `"rotation"` is a
      # string writexl reads from R, and says nothing about where it is put.
      if (!grepl(sprintf("[.>]%s\\b", f), src))
        missed <- c(missed, paste0(struct, "$", f))
    }
  }
  expect_equal(missed, character(0))
})

test_that("nothing is listed unused that writexl has since started setting", {
  # keeps UNUSED honest: a stale entry would hide a field going missing again
  skip_if(is.null(writexl_c()), "source tree not available")
  for (struct in names(UNUSED)) {
    fields <- struct_fields(STRUCT_HEADER[[struct]], struct)
    if (is.null(fields) || identical(fields, NA_character_)) next
    expect_equal(as.character(setdiff(UNUSED[[struct]], fields)),
                 character(0),
                 info = paste("UNUSED names a field", struct,
                              "no longer has"))
  }
})

test_that("the pattern enums writexl bridges are the sizes the headers declare", {
  # .CHART_PATTERN maps cell pattern names onto chart pattern numbers, two
  # enums that share names but not values.  A member added to either shifts
  # nothing by itself, but it does mean the bridge needs another look.
  fmt <- lxw_header("format.h")
  cht <- lxw_header("chart.h")
  skip_if(is.null(fmt) || is.null(cht), "source tree not available")
  # one is a plain `enum name {`, the other a `typedef enum name {`
  count <- function(txt, enum, prefix) {
    m <- regmatches(txt, regexpr(sprintf("enum %s \\{.*?\\n\\}", enum), txt))
    if (!length(m)) return(NA_integer_)
    length(regmatches(m, gregexpr(prefix, m))[[1]])
  }
  expect_equal(count(fmt, "lxw_format_patterns", "LXW_PATTERN_"), 19L)
  expect_equal(count(cht, "lxw_chart_pattern_type", "LXW_CHART_PATTERN_"), 49L)
  # every bridged number is inside the chart enum, and none is the cell one
  expect_true(all(.CHART_PATTERN < 49L))
  expect_equal(length(.CHART_PATTERN), 11L)
})

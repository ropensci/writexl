# writexl 1.5.4.9000

## New features

writexl now reaches the whole feature set of the libxlsxwriter it bundles.
Each entry below names the one or two functions to start from; the vignettes
carry the detail.

* **Cell content** — `xl_cell_general()` writes any combination of value,
  formula, hyperlink, format and comment, in mixed-type columns. It also
  carries array and dynamic array formulas, comments (`xl_comment()`) and rich
  strings, one cell in several fonts (`xl_rich_string()`). `xl_formula()` and
  `xl_hyperlink()` return these objects and stay backward compatible;
  `xl_hyperlink()` now writes a real URL hyperlink, with display text and a
  tooltip.

* **Formatting** — `xl_format()` and the group constructors `xl_font()`,
  `xl_fill()`, `xl_border()`, `xl_align()`, `xl_num_format()` and
  `xl_protection()` build reusable format objects that combine with `+`, and
  apply to a cell, a column, a row, a sheet or the workbook.

* **Conditional formatting** — `xl_sheet(conditional =)`, with `xl_cond_cell()`
  pairing a rule with a format and `xl_cond_scale()`, `xl_cond_bar()` and
  `xl_cond_icons()` for colour scales, data bars and icon sets.

* **Data validation** — `xl_sheet(validation = xl_validation(...))`: dropdown
  lists, numeric, date, time and text-length bounds, and custom formulas, with
  the input and error messages Excel shows (#43).

* **Autofilters** — `xl_sheet(filter = xl_filter(...))`. Excel does not apply a
  filter when a file is opened, so writexl also hides the rows the criteria
  exclude; without that the sheet looks filtered but shows every row.
  `xl_filter_keep()` exposes the same matching rule on its own. The rules
  reproduce Excel's, which were measured rather than assumed.

* **Worksheets and workbooks** — `xl_sheet()` carries column and row geometry
  (in Excel's units or in pixels), frozen and split panes, gridlines, tab state
  and the opening view, protection, outline display (`xl_outline()`), and the
  error indicators Excel shows on cells it believes are wrong. `xl_workbook()`
  and `xl_properties()` set document metadata — custom properties may now be
  `Date` or `POSIXct` — and the workbook-wide formatting defaults, including
  `hyperlink_format = NULL` for hyperlinks with no styling at all.

* **Page setup and printing** — `xl_sheet(page = xl_page_setup(...))`:
  orientation, paper size, margins, scaling and fit-to-pages, centring, headers
  and footers, print area, repeating heading rows and columns, manual page
  breaks, and the print options.

* **Tables** — `xl_sheet(table = xl_table(...))` and `xl_table_column()`: a
  named, styled range with banded rows, a filter dropdown, an optional total
  row, and per-column headers, formats and formulas.

* **Merged cells** — `xl_sheet(merge = xl_merge(...))`. A merged range holds one
  value, so `xl_merge()` carries its own; merging over cells the data frame
  filled keeps only that value, as it does in Excel.

* **Images** — `xl_sheet(image = xl_image(...))`, floating over the cells or
  placed inside one with `embed = TRUE`. The source may be a file path, a raw
  vector or an in-memory picture (a `raster`, colour matrix, RGB array or
  `nativeRaster`), so a plot never has to touch the disk. Also a tiled screen
  backdrop via `xl_sheet(background_image =)`, and images in printed headers and
  footers. Two arrangements libxlsxwriter miscounts — an embedded image
  alongside any other, and a header/footer or background image on a sheet before
  one with a floating image — are refused with the order that works.

* **Charts** — `xl_sheet(chart = xl_chart(...))` and `xl_chart_series()`, in all
  22 types Excel offers, with axes (`xl_chart_axis()`), the parts of a series
  (markers, data labels, trendlines, error bars and a format per point) and the
  chart's own furniture (legend, data table, plot and chart areas, manual
  layouts). `xl_chartsheet()` gives one chart a tab of its own.

  A series names its values and categories either as an A1 range
  (`"Data!B2:B10"`) or by column (`list(cols = "revenue")`), so a range follows
  the data when rows are added or a header is written, and a series that plots a
  column is named after that column's header. Series and titles are styled with
  the ordinary `xl_format()` groups — `xl_border()` becomes the line,
  `xl_fill()` the fill, `xl_font()` the title text — so one format object can
  style both a cell and a chart.

  Anything the chart cannot draw is refused by name rather than dropped
  silently, which is what Excel does with it: a format group no chart shape has,
  a value-axis option on a category axis, a doughnut hole on a bar chart, a data
  label in a position its chart type disallows. Tests read the function lists
  out of the bundled `chart.h`, so a function added upstream surfaces as a
  failure rather than as a gap.

* **A stand-in for missing values** via `na`, which writexl has always written
  as an empty cell (#76). `write_xlsx(df, na = "not measured")` sets it for a
  whole workbook, `xl_properties(na = )` does the same on a workbook object,
  and `xl_col_spec(na = )` and `xl_cell_general(na = )` narrow it to one column
  or one cell --- the innermost setting wins. It covers `NaN` as well as `NA`,
  and keeps its own type, so `na = 0` writes a number and leaves a numeric
  column numeric. The default, `na = NA`, is the empty cell as before.

* `xl_cell_general()` is a vector rather than a list, so a cell column can be
  assigned with `df[, j] <- cells` and not only with `df[[j]] <- cells`.
  `[<-.data.frame` reads a list value as a list of *columns*, which scattered
  one cell record per column and dropped the class; the records now ride in an
  attribute behind an integer index, the shape `factor` uses. `[`, `[[`, `c()`,
  `rep()` and `length()` are unchanged from a caller's point of view.

  `df[i, j] <- x` sets that cell's value. A formula needs `xl_formula()`: a
  cell column has no column-wide "these are all formulas", which is what
  writexl 1.5.4 carried on the column's class.

* Argument names are consistent across the new functions. Whatever a cell,
  label or box will show is `value`, whatever its type; a size in pixels says
  so (`width_pixels`); and a caption is `title`, with `title_format` and
  `title_layout` beside it. `as.character()` methods on `xl_cell_general()` and
  `xl_rich_string()` mean a cell built for a sheet can be reused anywhere a
  plain string is wanted.

* Columns of a type writexl cannot represent (`complex`, `raw`, a bare list
  column) now raise an error naming the column, rather than warning and leaving
  the cells empty.

* `write_xlsx()` now errors informatively when a data frame exceeds the xlsx
  column limit (16384) or row limit (1048576).

* Bundled libxlsxwriter updated to 1.2.4.

* See the "Getting started with writexl" vignette, and the five that follow it
  for formatting, worksheets and workbooks, charts and images, formulas and
  tables, and one runnable example of everything.

## Bug fixes

* Fix installation on systems without GNU make, by replacing a GNU-specific
  pattern rule in `src/Makevars` with a portable static library recipe
  (#97).

* Fix a `strcpy()` buffer overflow in the internal `C_set_tempdir()`; a tempdir
  path of 2048 bytes or more now errors informatively.

* `POSIXct` columns are no longer silently converted to UTC. When every datetime
  in the workbook shares one time zone, the zone is dropped and local
  wall-clock time is written; when they differ, all are converted to UTC with a
  warning. Code that relied on always getting the UTC instant will see shifted
  values. See the "Getting started with writexl" vignette.

* `Date` values before 1900-03-01 were written one day too late, because the
  `Date` writer did not compensate for Excel's phantom 1900-02-29. `Date` and
  `POSIXct` now write identical serial numbers.

* Sheet names are repaired properly: they are truncated to a genuine 31
  characters (previously 29, despite the warning saying 31), characters Excel
  forbids (`[ ] : * ? / \`) are replaced, edge apostrophes are stripped, and
  duplicates created by either repair are resolved. A sheet named `"2024/Q1"`
  previously produced a file Excel refused to open.

# writexl 1.5.4

* Fix LTO build for bundled libxlsxwriter

# writexl 1.5.3

* `write_xlsx()` now gives a warning if a column is of unsupported type
* Fix crash in `write_xlsx()` for corrupted data frames

# writexl 1.5.2

* Fix parallel make; cleanup after build

# writexl 1.5.0

* Update libxlsxwriter from b0c76b33

# writexl 1.4.2

* Bugfix for NA timestamps

# writexl 1.4.1

* Fix strict-prototypes warnings

# writexl 1.4.0

* Update libxlsxwriter to 1.0.3

# writexl 1.3.1

* Fix a unit test in R-devel for timezone attribute comparisons

# writexl 1.3

* `write_xlsx()` gains option `use_zip64` for 4GB+ file support
* libxlsxwriter error messages are printed to `REprintf` instead of `fprintf`
* Handle overly long or duplicate sheet names
* The help assistant only appears once per session

# writexl 1.2

* Update bundled libxlsxwriter 0.8.8
* `xl_formula()` and `xl_hyperlink()` now correctly support `NA`
* Oil clippy a bit

# writexl 1.1

* Update bundled libxlsxwriter 0.8.4
* Do not write blank xlsx strings for `NA` and `""` character values
* Coerce bit64 vectors to double with warning (xlsx does not have int64)

# writexl 1.0

* Save R `Date` types as proper datetime strings
* Update vendored libxlsxwriter to 0.7.6

# writexl 0.2

* Add support for lists in `write_xlsx()` to create xlsx with multiple sheets
* Automatically coerce columns of type `Date` and `hms` to strings

# writexl 0.1

* Initial CRAN release with clippy

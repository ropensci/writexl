# writexl 1.5.4.9000

## New features

* `xl_cell_general()` writes a cell with any combination of value, formula,
  hyperlink, format and comment, and supports mixed-type columns and recycling.
  `xl_formula()` and `xl_hyperlink()` now return `xl_cell_general` objects and
  remain backward compatible; `xl_hyperlink()` writes real URL hyperlinks, so
  display text and tooltips are available.

* Cell, worksheet, and workbook **formatting** is now supported. `xl_format()`
  and the group constructors `xl_font()`, `xl_fill()`, `xl_border()`,
  `xl_align()`, `xl_num_format()` and `xl_protection()` build reusable format
  objects that combine with `+`; `xl_sheet()` adds worksheet options (column and
  row geometry, frozen panes, gridlines, tab color, zoom, automatic column
  widths, autofilters, protection); `xl_workbook()` and `xl_properties()` set
  document metadata and the workbook-wide formatting defaults.

* Cell **comments** (notes) via `xl_cell_general(comment =)` and `xl_comment()`.

* **Array and dynamic array formulas** via `xl_cell_general(array =, dynamic =,
  array_range =)`.

* **Rich strings** — one cell whose text is split into differently formatted
  runs — via `xl_rich_string()` and `xl_rich_run()`.

* **Page setup and printing** via `xl_sheet(page = xl_page_setup(...))`:
  orientation, paper size, margins, scaling and fit-to-pages, centring, headers
  and footers, print area, repeating heading rows/columns, manual page breaks,
  and the print options.

* **Sheet tabs and the opening view** via `xl_sheet()`: which tab is active,
  selected, hidden or leftmost; the selected cell and scroll position; hiding
  zero values; right-to-left column order; and split panes alongside the
  existing frozen panes. The workbook-wide visibility rules (only one active
  sheet, no hidden-and-active sheet, at least one visible sheet) are checked
  before writing, naming the sheet at fault.

* **Data validation** via `xl_sheet(validation = xl_validation(...))`: dropdown
  lists, numeric, date, time and text-length bounds, and custom formulas, with
  the input and error messages Excel shows (#43).

* **Conditional formatting** via `xl_sheet(conditional = ...)`, in four
  flavours: `xl_cond_cell()` for a rule paired with a format, `xl_cond_scale()`
  for two- and three-colour scales, `xl_cond_bar()` for data bars and
  `xl_cond_icons()` for Excel's built-in icon sets.

* **Autofilter criteria** via `xl_sheet(filter = xl_filter(...))`. Excel does
  not apply a filter when a file is opened, so `xl_filter()` also hides the rows
  the criteria exclude; without that the sheet looks filtered but shows every
  row. writexl reproduces Excel's matching rules, which were measured rather
  than assumed. An exact value or a `list` matches the text a cell *displays*,
  so it matches the number `10` and the string `"10"` alike; every other
  criteria compares by type, and a wildcard turns `"=="` into one of those. Text
  matching is case-insensitive, `*` and `?` are wildcards, and blank covers an
  empty cell as well as an empty string. Mixed-type columns made with
  `xl_cell_general()` are matched cell by cell.

* **Worksheet tables** via `xl_sheet(table = xl_table(...))`: a named, styled
  range with banded rows, a filter dropdown, an optional total row, and
  per-column headers, formats and formulas from `xl_table_column()`. Column
  headers default to the data frame's column names, which is what keeps the
  table definition and the header cells in agreement — Excel refuses a file
  where they differ. Table names are made unique across the workbook. Any table
  turns off the memory-efficient row-streaming mode, which libxlsxwriter does
  not support alongside tables.

* `xl_filter_keep()` exposes the matching rule on its own: given a data frame
  and one or more `xl_filter()`s, it returns which rows Excel would leave
  visible, without writing anything. `xl_sheet(filter =)` uses it to decide
  what to hide, so the two cannot disagree.

* Column widths and row heights may be given in **pixels** ---
  `xl_col_spec(width_pixels =)` and `xl_row_spec(height_pixels =)` --- as well
  as in Excel's character units and points.

* **Outline display** via `xl_sheet(outline = xl_outline(...))`: whether the
  grouping symbols are shown, and which side of the detail rows and columns
  the summary sits on. Grouping itself still comes from `level`.

* **Error indicators** can be turned off with `xl_sheet(ignore_errors =)`, for
  example `list(number_stored_as_text = "A2:A99")` to stop Excel flagging
  numbers deliberately stored as text.

* **Merged cells** via `xl_sheet(merge = xl_merge(...))`. A merged range holds
  one value, so `xl_merge()` carries its own text; merging over cells the data
  frame filled keeps only the merged text, as it does in Excel.

* `xl_properties(hyperlink_format = NULL)` writes hyperlinks with no styling at
  all, which was previously impossible.

* Custom document properties may now be `Date` or `POSIXct`, and are written as
  real datetime properties rather than as text.

* Columns of a type writexl cannot represent (`complex`, `raw`, a bare list
  column) now raise an error naming the column, rather than warning and leaving
  the cells empty.

* `write_xlsx()` now errors informatively when a data frame exceeds the xlsx
  column limit (16384) or row limit (1048576).

* Bundled libxlsxwriter updated to 1.2.4.

* See the "Formatting and workbook properties" and "Writing Special Cell Types"
  vignettes.

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
  values. See the formatting vignette.

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

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

* **Images** via `xl_sheet(image = xl_image(...))`, floating over the cells or,
  with `embed = TRUE`, placed inside one. The image may be a file path, a raw
  vector, or an in-memory picture — a `raster`, colour matrix, RGB array or
  `nativeRaster` — so a plot never has to touch the disk. Scale, offset,
  position, alt text and a hyperlink are all supported, and the format is read
  from the file's own bytes rather than its extension. Also
  `xl_sheet(background =)` for a tiled screen watermark, and images in printed
  headers and footers via `xl_page_setup(header_image =, footer_image =)`.

  Two combinations are refused, because libxlsxwriter miscounts them and Excel
  repairs the file: an embedded image alongside any other image, and a
  header/footer or background image on a sheet before one with a floating
  image. Both errors name the sheets and the arrangement that works.

* **Worksheet tables** via `xl_sheet(table = xl_table(...))`: a named, styled
  range with banded rows, a filter dropdown, an optional total row, and
  per-column headers, formats and formulas from `xl_table_column()`. Column
  headers default to the data frame's column names, which is what keeps the
  table definition and the header cells in agreement — Excel refuses a file
  where they differ. Table names are made unique across the workbook. Any table
  turns off the memory-efficient row-streaming mode, which libxlsxwriter does
  not support alongside tables.

* **Charts** via `xl_sheet(chart = xl_chart(...))` and `xl_chart_series()`, in
  all 22 of the types Excel offers. A series names its values, categories and
  name either as an A1 range (`"Data!B2:B10"`) or by column
  (`list(cols = "revenue")`), so a range follows the data when rows are added
  or a header is written; a range may name another sheet, and one that reaches
  past the data is refused rather than plotted as blanks. A series name and a
  chart title are always literal text unless given as a spec, and
  `list(header = "revenue")` names a header cell. A series that plots a column
  is named after that column's header unless told otherwise, as it is when you
  chart a column with its header in Excel; `name = FALSE` leaves it unnamed.
  Placement uses the same vocabulary as an image (`at`, `scale`, `offset`,
  `position`, `description`, `decorative`).

  Series and titles are styled with the ordinary `xl_format()` groups:
  `xl_border()` becomes the line, `xl_fill()` the fill or pattern, and
  `xl_font()` the title text via `xl_chart(title_format =)`, so one format
  object can style both a cell and a chart. `xl_fill()` and `xl_border()` gained
  `transparency`, which charts honour and cells ignore. Anything the chart
  cannot draw is refused by name rather than dropped silently: border styles a
  chart line has no width for, format groups no chart shape has, a font on a
  series (a series is a shape, not text), and a fill or border on a title.

  **Axes** via `xl_chart(x_axis =, y_axis =)` and `xl_chart_axis()`, covering
  all 32 of libxlsxwriter's axis functions: the axis title and its font, tick
  labels and their number format and font, bounds, log scale, major and minor
  units and tick marks, tick-label position and alignment, category intervals,
  crossing, reversal, display units, gridlines and their styling, the axis line
  itself, and hiding the axis.

  Options that apply to one *kind* of axis are refused on the other, since
  Excel discards them silently. A scatter chart plots numbers against numbers,
  so both its axes are value axes; every other type has a category x axis and a
  value y axis, bar charts included --- Excel draws their categories up the
  side, but the axes keep their names. So `min`, `max`, `log_base`,
  `major_unit`, `minor_unit` and the display units are value-axis only, and
  `position`, `label_align`, `interval_unit` and `interval_tick` are
  category-axis only. Pie and doughnut charts have no axes at all. A test reads
  the applicability out of the bundled `chart.h` and checks every function is
  reached, so a function added upstream shows up as a failure rather than as a
  gap.

  **The parts of a series** --- `xl_chart_marker()` for the symbol at each
  point, `xl_chart_labels()` and `xl_chart_label()` for the numbers printed
  beside them, `xl_chart_trendline()` for a fit through the series,
  `xl_chart_error_bars()` for `x_error_bars` and `y_error_bars`, and an
  `xl_format()` per point in `points =` to colour one slice of a pie or one
  bar of a column chart. Custom labels give a single point its own text, its
  own styling, or `hide = TRUE`. A label holds exactly the parts named, so
  `show_percentage = TRUE` gives a percentage rather than Excel's
  "10, 11.8%"; naming none of them leaves Excel's default, the value.

  These carry three more restrictions libxlsxwriter documents and does not
  enforce: the positions a data label may take depend on the chart type (the
  whole table is in the header, and is checked), a moving average has no
  forecast, equation or R-squared, an intercept applies only to exponential,
  linear and polynomial fits, and an automatic marker cannot also be given a
  size or a format. A test reads the function list out of the bundled
  `chart.h`, so all 40 `chart_series_*()` functions are known to be reached.

  **The chart itself** --- `xl_chart_legend()` moves, styles or removes the
  legend and can leave a series out of the key; `xl_chart_table()` prints the
  plotted numbers in a grid beneath the chart; `plot_area_format` and
  `chart_area_format` style the panel and the surround; `title_layout`,
  `plot_area_layout` and `xl_chart_legend(layout =)` place those parts by hand;
  `show_blanks` says what an empty cell does to the plot, and
  `show_hidden_data` plots rows Excel would leave out.

  The six options that belong to one family of chart --- the doughnut hole, pie
  rotation, drop lines, high-low lines, up-down bars, and the bar gap and
  overlap --- are now reachable, and the feature matrix that has known about
  them since the first charts refuses them elsewhere. `drop_lines` and
  `high_low_lines` take `TRUE` or an `xl_format()` giving their line;
  `up_down_bars` takes `TRUE` or `list(up =, down =)`.

  Features that apply to only some chart types are checked against the chart's
  type throughout, because Excel discards them without a word. A scatter series
  must have `categories`: they are its x axis, and libxlsxwriter crashes
  without them.

* **Chartsheets** via `xl_chartsheet()`, a tab holding one chart and no cells.
  Give one to `write_xlsx()` in place of a data frame. Every range in its chart
  must name the sheet it plots, since a chartsheet has no cells of its own for
  a bare `list(cols = )` to resolve against. A chartsheet supports a subset of
  what a worksheet does --- of `xl_page_setup()` the orientation, paper,
  margins, header and footer; of `xl_sheet_view()` the four tab-state options
  --- and the rest are refused by name rather than dropped.

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

* Argument names are consistent across the new functions. Whatever a cell,
  label or box will show is `value`, whatever its type; a size in pixels says
  so (`width_pixels`); and a caption is `title`, with `title_format` and
  `title_layout` beside it. `as.character()` methods on `xl_cell_general()` and
  `xl_rich_string()` mean a cell built for a sheet can be reused anywhere a
  plain string is wanted.

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

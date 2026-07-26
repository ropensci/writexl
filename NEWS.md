# writexl 1.5.4.9000

## New features

* Cell, worksheet, and workbook **formatting** is now supported, built on a
  shared format engine:

  * `xl_format()` and the group constructors `xl_font()`, `xl_fill()`,
    `xl_border()`, `xl_align()`, `xl_num_format()` and `xl_protection()` build
    reusable format objects; `xl_color()` normalizes R color names / hex /
    integers. Formats combine property-by-property with `+`.

  * `xl_cell_general()` gains a `format` argument (a single format applied to
    every cell, or a list of per-cell formats). `xl_formula()`,
    `xl_hyperlink()` and `xl_hyperlink_cell()` also gain a `format` argument.

  * `xl_sheet()` wraps a data frame with worksheet options — column and row
    specifications (`xl_col_spec()`, `xl_row_spec()`, which are themselves
    `xl_format` subclasses), frozen panes, gridlines, tab color, zoom, default
    row height, automatic column widths (`auto_colwidth`), autofilters, and
    worksheet protection (with optional password and fine-grained options).
    Pass it anywhere `write_xlsx()` accepts a data frame.

  * `xl_workbook()` and `xl_properties()` set workbook-level document metadata
    (title, author, custom properties, ...), a few native settings (read-only
    recommended, window size, defined names), and the formatting defaults
    (default cell format, header style, hyperlink style, date/time formats),
    all as overridable `xl_format` objects.

* Cell **comments** (notes) are now supported. `xl_cell_general()` gains a
  `comment` argument: a character vector of text (the easy default), a single
  `xl_comment()` (recycled), or a per-cell list. `xl_comment()` exposes the full
  set of comment options (author, visibility, box size/position) and reuses the
  format engine for styling — only the fill background color and font
  name/size/family are supported, and any other format property triggers a
  warning. `xl_sheet()` gains `comment_author` and `show_comments` (comment
  defaults are worksheet-scoped); a comment's own `author` overrides the sheet
  default.

* Columns of a type writexl cannot represent (`complex`, `raw`, or a bare list
  column) now raise an **error** naming the offending column, its position and
  its type, instead of writing a warning and leaving the cells empty. Columns
  that can be written are unaffected, including classed numeric columns such as
  `difftime` (written from its underlying value) and the classes that are
  coerced on the way out (`factor`, `hms`, `POSIXlt`, `integer64`).

* Writing a cell that unlocks (`locked = FALSE`) or hides (`hidden = TRUE`) a
  cell on a worksheet that is not protected now warns, since the cell locking
  has no effect until the sheet is protected.

* See the new "Formatting and workbook properties" vignette.

## Bug fixes

* `POSIXct` columns are no longer silently converted to UTC. Excel has no
  concept of a time zone, so writexl now decides once per workbook:

  * if every datetime in the workbook shares one time zone, the zone is dropped
    and local wall-clock time is written -- `as.POSIXct("2025-12-01 01:00:00",
    tz = "Australia/Perth")` now appears in Excel as `2025-12-01 01:00:00`
    rather than `2025-11-30 17:00:00`. Daylight saving is handled per value.
    The default `datetime_format` also loses its `" UTC"` suffix in this case,
    since the value is no longer UTC;
  * if the time zones differ, all datetimes are converted to UTC (as before)
    and a warning is raised, because no single wall clock can represent them
    all. Convert them yourself beforehand if you want something else.

  Datetimes inside [xl_cell_general()] cells are included in this survey.
  Supplying your own `datetime_format` overrides the label either way. Code
  that relied on the previous behaviour of always writing the UTC instant will
  see shifted values when a non-UTC time zone is used throughout.

* `Date` columns holding dates before 1900-03-01 were written one day too late.
  Excel's 1900 date system wrongly treats 1900 as a leap year, and the `Date`
  writer did not compensate for that phantom 1900-02-29 (the `POSIXct` writer
  already did). As a result `1900-01-01` was written as `1900-01-02`, and
  `1900-02-28` landed on the nonexistent serial 60 so readers such as `readxl`
  returned `NA` with a warning about an "impossible 1900-02-29 datetime".
  `Date` and `POSIXct` columns now write identical serial numbers. Dates from
  1900-03-01 on, and all times written as `POSIXct`, are unaffected. Note
  that dates before 1900-01-01 still have no valid Excel serial number.

## Internal

* The header, hyperlink, and date/time formats that were previously hard-coded
  in C are now supplied from R as ordinary formats, so their defaults live in
  `xl_properties()` and can be overridden.

* All range addressing now goes through one resolver, which also accepts
  absolute references and whole-column / whole-row forms. As a result
  `xl_sheet(freeze =)` additionally accepts `"$A$2"`, and
  `xl_sheet(autofilter =)` additionally accepts `"$A$1:$D$51"`, `"B:D"` and
  `"2:10"`; every previously accepted spelling behaves exactly as before.

* Range-scoped worksheet features are applied through a single sheet-overlay
  stage that runs before rows are written (a requirement of the libxlsxwriter
  constant-memory mode). The autofilter is its first user.

* Whether the workbook is written in the libxlsxwriter constant-memory mode is
  now resolved from the features the workbook actually uses, rather than being
  hard-coded. It still always resolves to "on"; no feature yet requires it off.
  The `writexl.constant_memory` option that briefly existed to override this has
  been removed — the mode is not a user setting.

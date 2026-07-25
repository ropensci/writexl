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

* See the new "Formatting and workbook properties" vignette.

## Internal

* The header, hyperlink, and date/time formats that were previously hard-coded
  in C are now supplied from R as ordinary formats, so their defaults live in
  `xl_properties()` and can be overridden.

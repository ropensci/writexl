# writexl 2.0.0

This is a resubmission. The previous submission's pretest reported a reverse
dependency, BioMonTools, as a change to worse:

```
metvalgrpxl() ... Error: writexl cannot write column(s) of these types:
'X2' (column 2): list, 'X3' (column 3): list
```

## What changed in response

`xl_formula()` returns a per-cell object in 2.0.0 where 1.5.4 returned a
character vector with a class attached. BioMonTools assigned one into a data
frame column with `df[, j] <-`, which `[<-.data.frame` reads as a list of
columns rather than as one column, so the class was dropped and a bare list
column reached `write_xlsx()`.

Two changes address it:

* The cell object is now an atomic vector carrying its per-cell data in an
  attribute, rather than a list. `df[, j] <- cells` therefore assigns one
  column, as it did in 1.5.4.
* Assigning a string beginning with `=` into a cell column writes those
  characters, which is this version's documented behaviour but a quiet way to
  lose a sheet's formulas. It now warns and names the spelling that works.

BioMonTools 1.3.2, released upstream on 2026-08-04, also adopts the spelling
that works. Running its `metvalgrpxl()` example against BioMonTools 1.3.2 gives
an identical workbook under writexl 1.5.4 and under this version -- 62 formula
cells, all hyperlinks present, no warnings.

## Change of maintainer

The maintainer changes with this release, from Jeroen Ooms
(jeroenooms@gmail.com) to Bill Denney (wdenney@humanpredictions.com). Jeroen
Ooms remains an author. Confirmation from the outgoing maintainer is being sent
to CRAN separately.

## Why the major version

Most of this release is additive, but these change what an existing script
writes, so the version is 2.0.0 rather than 1.6.0:

* `POSIXct` columns are no longer silently converted to UTC.
* `Date` values before 1900-03-01 were written one day too late and now agree
  with `POSIXct`.
* Columns of a type the package cannot represent are an error rather than a
  warning with empty cells.
* Sheet-name repair truncates to a genuine 31 characters rather than 29.
* A formula is a property of a cell rather than of a column, so
  `df[i, j] <- "=SUM(A1:A2)"` writes text and warns, where 1.5.4 wrote a
  formula. This is the BioMonTools case above.

`xl_hyperlink(name = )` is deprecated in favour of `value`. `value` occupies the
argument position `name` used to, so positional calls are unaffected; supplying
`name` warns rather than failing. All of these are listed under "Breaking
changes" in NEWS.md.

## Test environments

* Windows 11, R 4.6.1 (local)
* GitHub Actions:
  * macOS (release, next)
  * Windows (4.1, 4.2, release, devel)
  * Ubuntu (oldrel-1, release, devel)

## R CMD check results

0 errors | 0 warnings | 0 notes

## Reverse dependencies

writexl has 87 strong reverse dependencies and 135 including Suggests.

All 135 were checked with revdepcheck against both the CRAN version and this
one. That run was made before the two changes described above, so it does not
cover them; the package flagged by it, BioMonTools, has been re-checked
individually against the current sources as described. No other package changed
its check result.

One package, ggpaintr, did not finish that run: its vignette hangs at "checking
running R code from vignettes", identically against both versions, so the hang
is a property of that package in the check environment rather than a change
here.

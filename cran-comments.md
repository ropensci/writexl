# writexl 2.0.0

This is a resubmission. The previous submission's pretest reported a reverse
dependency, BioMonTools, as a change to worse which has now been fixed by an
updated version of BioMonTools which works under both the current CRAN
version and this version of writexl.

## Change of maintainer

The maintainer changes with this release, from Jeroen Ooms
(jeroenooms@gmail.com) to Bill Denney (wdenney@humanpredictions.com). Jeroen
Ooms remains an author. He confirmed the change to CRAN by email during the
previous submission, and CRAN acknowledged it.

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
one. BioMonTools, has been re-checked individually against the current sources
as described. No other package changed its check result.

One package, ggpaintr, did not finish that run: its vignette hangs at "checking
running R code from vignettes", identically against both versions, so the hang
is a property of that package in the check environment rather than a change
here.

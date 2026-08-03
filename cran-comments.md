# writexl 2.0.0

## Change of maintainer

The maintainer changes with this release, from Jeroen Ooms
(jeroenooms@gmail.com) to Bill Denney (wdenney@humanpredictions.com). Jeroen
Ooms remains an author. Confirmation from the outgoing maintainer is being sent
to CRAN separately.

## Why the major version

Most of this release is additive, but four changes can alter what an existing
script writes, so the version is 2.0.0 rather than 1.6.0:

* `POSIXct` columns are no longer silently converted to UTC.
* `Date` values before 1900-03-01 were written one day too late and now agree
  with `POSIXct`.
* Columns of a type the package cannot represent are an error rather than a
  warning with empty cells.
* Sheet-name repair truncates to a genuine 31 characters rather than 29.

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

writexl has 87 strong reverse dependencies and 135 including Suggests. All were
checked with revdepcheck against both the CRAN version and this one, and no
package changed its check result.

One package, ggpaintr, did not finish: its vignette hangs at "checking running R
code from vignettes", identically against both versions, so the hang is a
property of that package in the check environment rather than a change here.

# writexl 2.0.1

This release fixes the WARN on the r-devel Fedora flavors, requested by CRAN
before 2026-08-21:

```
libxlsxwriter/src/worksheet.c:1979:9: warning: assignment discards 'const' qualifier from pointer target type [-Wdiscarded-qualifiers]
libxlsxwriter/src/worksheet.c:8420:18: warning: assignment discards 'const' qualifier from pointer target type [-Wdiscarded-qualifiers]
libxlsxwriter/src/worksheet.c:8425:18: warning: assignment discards 'const' qualifier from pointer target type [-Wdiscarded-qualifiers]
libxlsxwriter/src/worksheet.c:8436:26: warning: assignment discards 'const' qualifier from pointer target type [-Wdiscarded-qualifiers]
```

The pointers involved are never written through; they are now declared
`const` or the search result is tested directly. The warnings were reproduced
and confirmed gone by compiling the package's C sources with the C23
const-correct `<string.h>` semantics; the same change is being proposed to
upstream libxlsxwriter, which has the same code on its main branch.

There are no R-level changes.

## Reverse dependencies

This release changes no R-level interface or behavior; the only change
removes compiler warnings from the bundled C library. Reverse dependencies
were checked in full with revdepcheck for 2.0.0, published 2026-08-05.

# lintr configuration.  Written as .R rather than DCF so the reasons can sit
# next to what they explain.
#
# The aim is a gate that is green today, so anything it reports later is a real
# regression.  A linter left on with hundreds of standing violations reports
# nothing: it is red always, so it is read never.
#
# The linters switched off below are ones where writexl's own style differs
# from lintr's defaults, or where the linter cannot see enough here to be
# right.  Everything else in lintr's default set stays on and passes.

linters <- linters_with_defaults(

  # Verified false positives: lintr resolves neither the package namespace (so
  # every internal `.helper()` reads as undefined) nor the objects a test file
  # or vignette chunk gets from the package.  Loading the namespace first and
  # re-running gives no lints.  R CMD check's own "checking R code for possible
  # problems" does have the namespace, and is clean.
  object_usage_linter = NULL,

  # False positives: it reads a trailing comment naming a C function --
  # `# chart_axis_set_position()` -- or a units note -- `# "1234.5" = 6` -- as
  # commented-out code.  There is none in the package.
  commented_code_linter = NULL,

  # The lookup tables mirroring libxlsxwriter's enums are named for them:
  # .LXW_CHART_TYPE, .CHART_PATTERN, .AXIS_OPTION_KIND.  The upper case is what
  # marks them as the header's vocabulary rather than ordinary variables.
  object_name_linter = NULL,

  # Inherited style, applied consistently in the files that use it: `if(cond){`
  # with no space before the paren or brace, and paired assignments on one line
  # (`r1 <- ...; r2 <- ...`) where the pair reads as a single step.  Converting
  # the package to lintr's model would touch almost every line for no
  # behavioural gain.
  brace_linter = NULL,
  paren_body_linter = NULL,
  spaces_left_parentheses_linter = NULL,
  semicolon_linter = NULL,
  indentation_linter = NULL,

  # The R sources aim at 80 and mostly hold it; 100 is the ceiling the whole
  # package -- sources, tests and vignettes -- currently meets, so a report
  # here means a line that has genuinely got away.
  line_length_linter = line_length_linter(100L),

  # writexl supports the R versions its CI checks, the oldest being 4.1, and a
  # function that does not exist there is invisible on a modern machine until
  # CI says so.  `%||%` is excepted because the package defines its own in
  # xl_format.R rather than relying on base's, which arrived in 4.4.
  backport_linter = backport_linter("4.1.0", except = "%||%")
)

exclusions <- list(
  # An HTML <link> tag whose href is a URL: nothing to wrap.
  "R/write_xlsx.R" = list(line_length_linter = 15),
  # A markdown table row; breaking it would break the table.
  "vignettes/d-charts-images.Rmd" = list(line_length_linter = 101),
  # Vendored third-party C, never edited here.
  "src/libxlsxwriter"
)

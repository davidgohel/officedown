## Submission

This release fixes the errors shown on the CRAN check page (issues to be
fixed before 2026-10-22): the tests did not work with pandoc 3.12 and later.

## Test environments

- local macOS install (R 4.6.1, pandoc 3.12.1)
- GitHub Actions: Windows, macOS and Ubuntu (R release and devel), with
  pandoc 2.19, 3.8.3 and the latest release
- winbuilder (release and devel)

## R CMD check results

There were no ERROR, WARNING or NOTE.

On winbuilder, a test of the odd and even page headers is skipped: its
library contains a pre-release build of officer 0.7.7 (packaged
2026-10-14 16:42 UTC) that does not include a fix shipped in the officer 0.7.7
published on CRAN (packaged 2026-10-14 19:01 UTC).

## Reverse dependencies

There were no error or warning related to officedown. The plot hook was
rewritten in this version; it stays compatible with accessr, which modifies
it, and the accessr maintainer has been notified.

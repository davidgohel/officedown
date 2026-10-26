# officedown 0.4.4

## Issues

- The style of table captions made by flextable is now applied when its id
differs from its name, e.g. the style `caption` of the default
`reference_docx`, whose id is `Lgende` (#112). flextable writes an id
guessed from the style name; an unknown paragraph style id is now replaced by
the id of the style whose name gives this guess. Word was using the "Normal"
style instead.

## Changes

- The unused `external_img()` call kept in the plot hook for package accessr
is removed: accessr >= 1.1.4 no longer modifies the hook. Older versions of
accessr fail in `rmd2word()`. The test checking this line was failing under
covr (#154).

# officedown 0.4.3

## Issues

- `hook_plot_officedown.Rd` has no usage, and a fix was needed to avoid a NOTE 
(seen in CRAN Debian check).

# officedown 0.4.2

## Issues

- The new plot hook stays compatible with package accessr, whose
`rmd2word()` modifies it.
- Loading officedown in a document using another Word output format (e.g.
`word_document`) no longer produces a file that Word can not open: data frames
are printed as usual, and printing officer objects (such as `block_caption()`)
stops with a message asking to use `officedown::rdocx_document` (#121).
- Cross-references to sections (`\@ref(section-id)`) work again with pandoc
3.12 and later, where Word displayed "Error! Reference source not found."
instead of the section number. Section identifiers can now also contain `_`
or `.`.
- Headers and footers of the `reference_docx` now appear on every page,
including after a landscape or multi-column section (#127), and no longer
disappear from even pages.
- Landscape orientation set with `page_size` (e.g. `orient: "landscape"`) is
now applied even when `page_margins` is not set (#145).
- Bullet and numbered lists now use the list style set with `ul.style` or
`ol.style`, even when the `reference_docx` contains another list style with a
similar name (e.g. `new-ul` and `new-ul-old`).
- Data frames no longer disappear from HTML documents and other non Word
outputs when officedown is loaded, for example in a Quarto project.
- Cross-references and captions written in Chinese or other non Latin scripts
now work on computers that do not use UTF-8 by default (#37).

## Changes

- Figure sizes set with `out.width` and `out.height` can use any unit
(e.g. `"10cm"`, `"5in"`, `"80%"`), and `out.extra` adds other image
attributes. A percentage is now relative to the width of the page text, as in
other R Markdown formats, and no longer to the figure size. Figures wider than
the page text are reduced to fit it.
- The documentation now explains how to cross-reference tables (chunk option
`tab.id`) and figures (chunk label or `fig.id`).
- officedown now requires officer 0.7.7.

# officedown 0.4.1

## Changes

- Default page size is now set to exact A4 format for Word documents
- prepare paragraph styles management for officer > 0.6.8

# officedown 0.4.0

## new feature

- plots now accept chunk options `out.width` and `out.height`
if expressed in percent. #144

## issues

- fix `get_fun()` function
- fix BugReports links for CRAN check

# officedown 0.3.3

## issues

- fix manual links for CRAN check

# officedown 0.3.3

## issues

- drop unreproducible test failure for CRAN on linux 

# officedown 0.3.1

## issues

- Compliance of pandoc version comparison in `knit_print.dml()` function 
following CRAN e-mail.

# officedown 0.3.0

## new feature

- data.frame in Word now support chunk options `tab.alt.title` 
and `tab.alt.description`, used as alternative text for tables.
- Document sections are set to NULL by default.
- PowerPoint support for `fpar()` and `block_list()`.

## issues

- fix usage of bookmark id containing underscore
- update default values according to new default values of 
`officer::opts_current_table()`.
- dml background color `bg` is now read in R Markdown.
- require officer `>= 0.5.1` to avoid page breaking between sections

## Internals

- simplify/drop code with officer 0.4.5
- add tests

# officedown 0.2.4

## new feature

* emit a knitr chunk option `is_rdocx_document` set to `TRUE` when starting 
an `rdocx_document`. This will help with programing as it can be used to 
detect if in an `rdocx_document` context or not.

## issues

* bookmark only numbers
* capture `fig.cap.fp_text` and not `fig.fp_text`
* `topcaption` is now set to TRUE by default for tables as it is 
set by default in officer.


# officedown 0.2.3

## new feature

* support for knitr chunk options `tab.topcaption` and `tab.lp`.
* support for knitr chunk options `tab.cap.tnd`, `tab.cap.tns` and `tab.cap.fp_text` 
that let define prefix for auto-number table captions. Same with 
`fig.cap.tnd`, `fig.cap.tns` and `fig.cap.fp_text` for plots.

## Issues

* fix issue missing Word headers and footers.

# officedown 0.2.2

## new feature

* Added the control of page dimensions; it is now possible to define the default 
section of a document. 

## Issues

* set `number_sections` to FALSE when `bookdown::markdown_document2` is 
used to avoid sections numbered twice.

## new feature

* support for knitr chunk option `fig.alt` and `fig.topcaption`.

# officedown 0.2.1

## Issues

* fix section issue with margin sizes (now totally handled by officer)
* fix rendering with writing to intermediates_dir
* fix regression with cross-references

# officedown 0.2.0

## Changes 

* Changed RStudio menu label 

## Issues

* fix issue with PowerPoint when no `reference_doc` is provided.

# officedown 0.0.7

* Added a `NEWS.md` file to track changes to the package.
* Big refactoring


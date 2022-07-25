# officedown 0.2.5

## issues

* fix usage of bookmark id containing underscore

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


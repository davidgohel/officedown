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


# sections ----

#' @title section
#'
#' @description Add a section as a block output in an R
#' Markdown document. A section affects preceding paragraphs
#' or tables.
#'
#' @details
#' A section starts at the end of the previous section (or the beginning of
#' the document if no preceding section exists), and stops where the section is declared.
#' @param w,h width and height in inches of the section page. This will
#' be ignored if the default section (of the `reference_docx` file)
#' already has width and height.
#' @noRd
#' @rdname sections
#' @name sections
#' @importFrom officer block_section prop_section
block_section_continuous <- function( ){
  block_section(prop_section(type = "continuous"))
}

#' @noRd
#' @importFrom officer page_size
#' @param break_page break page type. It defines how the contents of the section will be
#' placed relative to the previous section. Available types are "evenPage"
#' (begins on the next, "nextPage" (begins on the following page), "oddPage"
#' (begins on the next odd-numbered page).
block_section_landscape <- function( w = 11906 / 1440, h = 16838 / 1440, break_page = "oddPage" ){
  block_section(prop_section(
    page_size = page_size(width = w, height = h, orient = "landscape"),
    type = break_page))
}

#' @noRd
block_section_portrait <- function( w = 16838 / 1440, h = 11906 / 1440, break_page = "oddPage"){
  block_section(prop_section(
    page_size = page_size(width = w, height = h, orient = "portrait"),
    type = break_page))
}

#' @noRd
#' @param widths columns widths in inches. If 3 values, 3 columns
#' will be produced.
#' @param space space in inches between columns.
#' @param sep if TRUE a line is separating columns.
#' @importFrom officer section_columns
block_section_columns <- function(widths = c(2.5,2.5), space = .25, sep = FALSE){
  block_section(prop_section(
    section_columns = section_columns(widths = widths, space = space, sep = sep),
    type = "continuous"))
}


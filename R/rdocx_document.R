# utils ----
#' @importFrom utils getAnywhere getFromNamespace
get_fun <- function(x){
  if( grepl("::", x, fixed = TRUE) ){
    coumpounds <- strsplit(x, split = "::", x, fixed = TRUE)[[1]]
    z <- getFromNamespace(coumpounds[2], ns = coumpounds[1] )
  } else {
    z <- getAnywhere(x)
    if(length(z$objs) < 1){
      stop("could not find any function named ", shQuote(z$name), " in loaded namespaces or in the search path. If the package is installed, specify name with `packagename::function_name`.")
    }
  }
  z
}

#' @export
#' @title Convert to an MS Word document
#' @description Format for converting from R Markdown to an MS Word
#' document. The function comes also with improved output options.
#' \code{rdocx_document2} also supports cross reference based on the syntax of
#' the bookdown package.
#' @param mapstyles a named list of style to be replaced in the generated
#' document. \code{list("Date"="Author")} will result in a document where
#' all paragraphs styled with stylename "Date" will be styled with
#' stylename "Author".
#' @param base_format a scalar character, format to be used as a base document for
#' officedown. default to \link[rmarkdown]{word_document} but
#' can also be word_document2 from bookdown
#' @param ... arguments used by \link[rmarkdown]{word_document}
#' @examples
#' skeleton <- system.file(package = "officedown",
#'   "rmarkdown/templates/word/skeleton/skeleton.Rmd")
#' file.copy(skeleton, to = "officedown.Rmd")
#' library(rmarkdown)
#' render("officedown.Rmd", output_file = "officedown.docx")
#' @importFrom officer change_styles
rdocx_document <- function(mapstyles, base_format = "rmarkdown::word_document", ...) {

  base_format_fun <- get_fun(base_format)
  output_formats <- base_format_fun(...)

  if( missing(mapstyles) )
    mapstyles <- list()

  output_formats$pre_processor = function(metadata, input_file, runtime, knit_meta, files_dir, output_dir){
    md <- readLines(input_file)
    md <- chunk_macro(md)
    md <- block_macro(md)
    writeLines(md, input_file)
  }

  output_formats$post_processor <- function(metadata, input_file, output_file, clean, verbose) {
    x <- officer::read_docx(output_file)
    x <- process_images(x)
    x <- process_links(x)
    x <- process_embedded_docx(x)
    x <- process_chunk_style(x)
    x <- process_sections(x)
    x <- process_par_settings(x)
    x <- change_styles(x, mapstyles = mapstyles)

    print(x, target = output_file)
    output_file
  }
  output_formats$bookdown_output_format = 'docx'
  output_formats

}


#' @rdname rdocx_document
#' @importFrom bookdown markdown_document2
#' @export
rdocx_document2 <- function(...) {
  markdown_document2(..., base_format = rdocx_document)
}

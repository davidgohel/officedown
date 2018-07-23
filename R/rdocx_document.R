#' @export
#' @title Convert to an MS Word document
#' @description Format for converting from R Markdown to an MS Word
#' document. The function comes also with improved output options.
#' @param mapstyles a named list of style to be replaced in the generated
#' document. \code{list("Date"="Author")} will result in a document where
#' all paragraphs styled with stylename "Date" will be styled with
#' stylename "Author".
#' @param ... arguments used by \link[rmarkdown]{word_document}
#' @examples
#' skeleton <- system.file(package = "worded",
#'   "rmarkdown/templates/worded/skeleton/skeleton.Rmd")
#' file.copy(skeleton, to = "worded.Rmd")
#' library(rmarkdown)
#' render("worded.Rmd", output_file = "worded.docx")
#' @importFrom officer change_styles
rdocx_document <- function(mapstyles, ...) {

  output_formats <- rmarkdown::word_document(...)

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

  output_formats
}



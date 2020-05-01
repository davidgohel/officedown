#' @export
#' @title Advanced R Markdown PowerPoint Format
#' @description Format for converting from R Markdown to an MS PowerPoint
#' document. \code{rpptx_document2} also supports cross reference based on the
#' syntax of the bookdown package.
#' @param base_format a scalar character, format to be used as a base document for
#' officedown. default to [powerpoint_presentation][rmarkdown::powerpoint_presentation] but
#' can also be powerpoint_presentation2 from bookdown
#' @param ... arguments used by [powerpoint_presentation][rmarkdown::powerpoint_presentation]
#' @examples
#' library(rmarkdown)
#' if(require("ggplot2")){
#'   skeleton <- system.file(package = "officedown",
#'     "rmarkdown/templates/powerpoint/skeleton/skeleton.Rmd")
#'   rmd_file <- tempfile(fileext = ".Rmd")
#'   file.copy(skeleton, to = rmd_file)
#'
#'   pptx_file_1 <- tempfile(fileext = ".pptx")
#'   render(rmd_file, output_file = pptx_file_1)
#' }
rpptx_document <- function(base_format = "rmarkdown::powerpoint_presentation", ...) {

  base_format <- get_fun(base_format)
  output_formats <- base_format(...)

  output_formats$post_processor <- function(metadata, input_file, output_file, clean, verbose) {
    x <- officer::read_pptx(output_file)

    # iterate over slides - not very nice but no better solution for now
    for(slide_i in seq_len(length(x)) ){
      x <- on_slide(x, index = slide_i)
      slide <- x$slide$get_slide(x$cursor)

      blips <- xml_find_all(slide$get(), "//a:blip[@r:embed]")
      is_hacked <- grepl( "\\.png$", xml_attr(blips, "embed") )
      blips <- blips[is_hacked]
      imgs <- xml_attr(blips, "embed")[is_hacked]
      for( img in imgs ) {
        slide$reference_img(src = img, dir_name = file.path(x$package_dir, "ppt/media"))
        if(clean) file.remove(img)
      }
      rel_df <- slide$rel_df()
      rids <- rel_df$id[match( imgs, rel_df$ext_src)]
      for( rid in seq_along(rids) ){
        xml_attr(blips[rid], "r:embed") <- rids[rid]
      }
    }

    print(x, target = output_file)
    output_file
  }

  output_formats
}


#' @rdname rpptx_document
#' @importFrom bookdown markdown_document2
#' @export
rpptx_document2 <- function(...) {
  rpptx_document(..., base_format = rpptx_document)
}

#' @importFrom officer get_reference_value read_pptx
get_pptx_uncached <- function() {
  ref_pptx <- read_pptx(get_reference_value(format = "pptx"))
  ref_pptx
}

#' @noRd
#' @importFrom memoise memoise
get_reference_pptx <- memoise(get_pptx_uncached)



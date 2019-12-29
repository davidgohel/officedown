#' @export
#' @title Convert to an MS PowerPoint document
#' @description Format for converting from R Markdown to an MS PowerPoint
#' document. \code{rpptx_document2} also supports cross reference based on the
#' syntax of the bookdown package.
#' @param ... arguments used by \link[rmarkdown]{powerpoint_presentation}
rpptx_document <- function(...) {

  output_formats <- rmarkdown::powerpoint_presentation(...)

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


#' @rdname rdocx_document
#' @importFrom bookdown markdown_document2
#' @export
rdocx_document2 <- function(...) {
  rpptx_document(..., base_format = rpptx_document)
}

#' @importFrom knitr opts_chunk
#' @title knitr hook for figure caption autonumbering
#' @description The function allows you to add a hook when executing
#' knitr to allow to turn figure captions into auto numbered figure
#' captions.
#' @noRd
plot_word_fig_caption <- function(x, options) {

  if(!is.character(options$fig.cap)) options$fig.cap <- NULL
  if(is.null(options$fig.id))
    fig.id <- options$label
  else fig.id <- options$fig.id
  if(!is.logical(options$fig.topcaption)) options$fig.topcaption <- FALSE
  
  bc <- block_caption(label =  options$fig.cap, style = options$fig.cap.style,
                      autonum = run_autonum(
                        seq_id = gsub(":$", "", options$fig.lp),
                        pre_label = options$fig.cap.pre,
                        post_label = options$fig.cap.sep,
                        bkm = fig.id, bkm_all = FALSE
                      ))
  cap_str <- to_wml(bc, knitting = TRUE)

  fig.width <- opts_current$get("fig.width")
  if(is.null(fig.width)) fig.width <- 5
  fig.height <- opts_current$get("fig.height")
  if(is.null(fig.height)) fig.height <- 5

  img <- external_img(src = x[1], width = fig.width, height = fig.height)

  doc <- get_reference_rdocx()
  si <- styles_info(doc)
  fig.style_id <- style_id(opts_current$get("fig.style"), type = "paragraph", si)

  if(length(fig.style_id) != 1 ){
    warning("paragraph style for plots ", shQuote(opts_current$get("fig.style")),
            " has not been found in the reference_docx document.",
            " Style 'Normal' will be used instead.",
            call. = FALSE)
    fig.style_id <- style_id("Normal", type = "paragraph", si)

  }
  ooxml <- paste0("<w:p><w:pPr><w:jc w:val=\"%s\"/><w:pStyle w:val=\"%s\"/></w:pPr>",
         to_wml(img),
         "</w:p>"
         )
  ooxml <- sprintf(ooxml, opts_current$get("fig.align"), fig.style_id)
  img_wml <- paste("```{=openxml}", ooxml, "```", sep = "\n")

  if (options$fig.topcaption)
    paste("", cap_str, img_wml, sep = "\n\n")
  else
    paste("", img_wml, cap_str, sep = "\n\n")
}

knitr_opts_current <- function(x, default = FALSE){
  x <- knitr::opts_current$get(x)
  if(is.null(x)) x <- default
  x
}

#' @importFrom knitr opts_chunk
#' @title knitr hook for figure caption autonumbering
#' @description The function allows you to add a hook when executing
#' knitr to allow to turn figure captions into auto numbered figure
#' captions.
#' @noRd
plot_word_fig_caption <- function(x, options) {

  if (grepl("^(ftp|ftps|http|https)://", x[1])) {
    stop("Images in 'rdocx_document' must be local files accessible without an internet connection:\n",
         shQuote(x[1]), call. = FALSE)
  }

  if(!is.character(options$fig.cap)) options$fig.cap <- NULL
  if(!is.character(options$fig.alt)) options$fig.alt <- NULL
  if(is.null(options$fig.id))
    fig.id <- options$label
  else fig.id <- options$fig.id
  if(!is.logical(options$fig.topcaption)) options$fig.topcaption <- FALSE

  tnd <- knitr_opts_current("fig.cap.tnd", default = 0)
  tns <- knitr_opts_current("fig.cap.tns", default = "-")
  fig.fp_text <- knitr_opts_current("fig.cap.fp_text", default = fp_text_lite(bold = TRUE))


  bc <- block_caption(label =  options$fig.cap, style = options$fig.cap.style,
                      autonum = run_autonum(
                        seq_id = gsub(":$", "", options$fig.lp),
                        pre_label = options$fig.cap.pre,
                        post_label = options$fig.cap.sep,
                        bkm = fig.id, bkm_all = FALSE,
                        tnd = tnd, tns = tns,
                        prop = fig.fp_text
                      ))
  cap_str <- to_wml(bc, knitting = TRUE)

  # physical size of plots
  fig.width <- opts_current$get("fig.width")
  if(is.null(fig.width)) fig.width <- 5
  fig.height <- opts_current$get("fig.height")
  if(is.null(fig.height)) fig.height <- 5

  # out.width and out.height in percent
  fig.out.width <- opts_current$get("out.width")
  has_fig_out_width <- !is.null(fig.out.width)
  is_pct_width <- has_fig_out_width && grepl("%", fig.out.width)
  if (is_pct_width) {
    fig.out.width <- gsub("%", "", fig.out.width, fixed = TRUE)
    fig.out.width <- as.numeric(fig.out.width) / 100
    fig.out.height <- fig.out.width
  } else {
    fig.out.height <- opts_current$get("out.height")
    has_fig_out_height <- !is.null(fig.out.height)
    is_pct_height <- !is_pct_width && has_fig_out_height && grepl("%", fig.out.height)
    if (is_pct_height) {
      fig.out.height <- gsub("%", "", fig.out.height, fixed = TRUE)
      fig.out.height <- as.numeric(fig.out.height) / 100
      fig.out.width <- fig.out.height
    }
  }

  if (!has_fig_out_width && !has_fig_out_height) {
    fig.out.width <- 1
    fig.out.height <- 1
  }

  fig.width <- fig.width * fig.out.width
  fig.height <- fig.height * fig.out.height

  img <- external_img(src = x[1], width = fig.width, height = fig.height, alt = options$fig.alt)

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
  ooxml <- "<w:p><w:pPr><w:jc w:val=\"%s\"/><w:pStyle w:val=\"%s\"/></w:pPr>"
  ooxml <- sprintf(ooxml, opts_current$get("fig.align"), fig.style_id)
  ooxml <- paste0(ooxml,
         to_wml(img),
         "</w:p>"
         )
  img_wml <- paste("```{=openxml}", ooxml, "```", sep = "\n")

  if (options$fig.topcaption)
    paste("", cap_str, img_wml, sep = "\n\n")
  else
    paste("", img_wml, cap_str, sep = "\n\n")
}

knitr_opts_current <- function(x, default = FALSE){
  x <- knitr::opts_current$get(x)
  if(is.null(x)) x <- default
  x
}

#' @importFrom knitr opts_chunk opts_knit
#' @title knitr hook for figure caption autonumbering
#' @description
#' Plot hook for Word documents that generates auto-numbered figure captions
#' with custom styling. The hook uses Pandoc markdown syntax for images
#' combined with OOXML for captions and paragraph formatting.
#'
#' The hook delegates image rendering to Pandoc while providing enhanced
#' caption management with auto-numbering, cross-references, and custom
#' Word paragraph styles.
#'
#' @param x Character string containing the path to the saved plot file.
#' @param options Named list of chunk options.
#'
#' @section Supported chunk options:
#' - `fig.cap`: Figure caption text. Used to generate auto-numbered captions
#'   with cross-reference support.
#' - `fig.cap.style`: Paragraph style for the caption (default from options).
#' - `fig.cap.pre`, `fig.cap.sep`: Caption prefix and separator for
#'   auto-numbering (e.g., "Figure" and ": ").
#' - `fig.id`: Cross-reference identifier (defaults to chunk label).
#' - `fig.topcaption`: Logical; if TRUE, caption appears above the image.
#' - `out.width`, `out.height`: Output dimensions. Can use units like "5in",
#'   "80%", "10cm". Passed directly to Pandoc.
#' - `out.extra`: Additional Pandoc image attributes.
#' - `fig.alt`: Alternative text for the image.
#' - `fig.align`: Image alignment. Must be one of "default", "left", "center",
#'   or "right". Default value "default" is converted to "center".
#' - `fig.style`: Word paragraph style name for the image paragraph
#'   (default: "Normal").
#'
#' @section Global options:
#' - `fig.cap.tnd`: Caption numbering depth (default: 0).
#' - `fig.cap.tns`: Caption numbering separator (default: "-").
#' - `fig.cap.fp_text`: Text formatting properties for caption numbers.
#'
#' @return
#' A character string containing markdown with embedded OOXML code that
#' combines the auto-numbered caption and the Pandoc image syntax.
#'
#' @examples
#' \dontrun{
#' # Chunk with auto-numbered caption and custom size
#' ```{r, fig.cap="Sales over time", out.width="80%"}
#' plot(sales_data)
#' ```
#'
#' # Chunk with alignment and custom style
#' ```{r, fig.cap="Key findings", fig.align="center", fig.style="ImageCenter"}
#' plot(results)
#' ```
#'
#' # Caption above image with cross-reference
#' ```{r sales-plot, fig.cap="Q4 Sales", fig.topcaption=TRUE, fig.id="sales"}
#' plot(sales_data)
#' ```
#' }
#' @rdname hook_plot_officedown
#' @name hook_plot_officedown
NULL

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

  # Build Pandoc markdown image with attributes
  base <- opts_knit$get('base.url')
  if (is.null(base)) base <- ''

  # Collect image attributes (width, height, out.extra)
  attrs <- c()
  if (!is.null(options$out.width)) {
    attrs <- c(attrs, sprintf('width=%s', options$out.width))
  }
  if (!is.null(options$out.height)) {
    attrs <- c(attrs, sprintf('height=%s', options$out.height))
  }
  if (!is.null(options$out.extra)) {
    attrs <- c(attrs, options$out.extra)
  }

  # Build attribute string
  attr_str <- ''
  if (length(attrs) > 0) {
    attr_str <- paste0('{', paste(attrs, collapse = ' '), '}')
  }

  # Generate Pandoc markdown: ![alt](path){attributes}
  alt_text <- if (!is.null(options$fig.alt)) options$fig.alt else ''
  img_markdown <- sprintf('![%s](%s%s)%s', alt_text, base, x[1], attr_str)

  # Add fig.align via inline R code with fp_par
  fig.align <- opts_current$get("fig.align") %||% "center"
  valid_aligns <- c("default", "left", "right", "center")
  if (!fig.align %in% valid_aligns) {
    warning("fig.align must be one of ",
            paste(shQuote(valid_aligns), collapse = ", "),
            ". Using 'center' instead.", call. = FALSE)
    fig.align <- "center"
  }
  if (fig.align == "default") {
    fig.align <- "center"
  }
  fig.style <- opts_current$get("fig.style") %||% "Normal"

  par_sty_wml <- to_wml(fp_par_lite(text.align = fig.align, word_style = fig.style))
  img_markdown <- paste0(img_markdown, " `", par_sty_wml, "`{=openxml}")
#
  if (options$fig.topcaption)
    paste("", cap_str, img_markdown, sep = "\n\n")
  else
    paste("", img_markdown, cap_str, sep = "\n\n")
}

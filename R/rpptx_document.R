#' @export
#' @title Advanced R Markdown PowerPoint Format
#' @description Format for converting from R Markdown to an MS PowerPoint
#' document.
#'
#' The function will allow you to specify the destination of your chunks in the output
#' PowerPoint file. In this case, you must specify the `layout` and `master` for the
#' layout you want to use, as well as the `ph` argument, which will allow you to specify
#' the placeholder to be generated to place the result. Use the officer package to help
#' you choose the identfiers to use.
#'
#' This function also support Vector graphics output in an editable format (using package
#' `rvg`). Wrap you R plot commands with function `dml` to use this graphic capability.
#' @param base_format a scalar character, format to be used as a base document for
#' officedown. default to [powerpoint_presentation][rmarkdown::powerpoint_presentation] but
#' can also be powerpoint_presentation2 from bookdown
#' @param layout default slide layout name to use
#' @param master default master layout name where \code{layout} is located
#' @param tcf default conditional formatting settings
#' defined by [officer::table_conditional_formatting()]
#' @param ... arguments used by [powerpoint_presentation][rmarkdown::powerpoint_presentation]
#' @return R Markdown output format to pass to [render][rmarkdown::render]
#' @examples
#' library(rmarkdown)
#' run_ok <- pandoc_available() && pandoc_version() > numeric_version("2.4")
#'
#' if(run_ok){
#'   example <- system.file(package = "officedown",
#'     "examples/minimal_powerpoint.Rmd")
#'   rmd_file <- tempfile(fileext = ".Rmd")
#'   file.copy(example, to = rmd_file)
#'
#'   pptx_file_1 <- tempfile(fileext = ".pptx")
#'   render(rmd_file, output_file = pptx_file_1)
#' }
rpptx_document <- function(base_format = "rmarkdown::powerpoint_presentation",
                           layout = "Title and Content",
                           master = "Office Theme",
                           tcf = list(),
                           ...) {


  args <- list(...)
  if(is.null(args$reference_doc)){
    reference_doc <- officer:::get_default_pandoc_data_file(format = "pptx")
  } else reference_doc <- args$reference_doc
  new_reference_doc <- tempfile(fileext = ".pptx")
  officer::annotate_base(reference_doc, output_file = new_reference_doc)
  args$reference_doc <- new_reference_doc


  base_format <- get_fun(base_format)
  output_formats <- do.call(base_format, args)

  tcf <- modifyList(tables_default_values$conditional, tcf)

  output_formats$knitr$opts_chunk <- append(
    output_formats$knitr$opts_chunk,
    list(layout = layout, master = master,
         first_row = tcf$first_row,
         first_column = tcf$first_column,
         last_row = tcf$last_row,
         last_column = tcf$last_column,
         no_hband = tcf$no_hband,
         no_vband = tcf$no_vband
         )
  )

  output_formats$post_processor <- function(metadata, input_file, output_file, clean, verbose) {
    x <- officer::read_pptx(output_file)
    forget(get_reference_pptx)
    print(x, target = output_file)
    output_file
  }

  output_formats
}

#' @importFrom officer get_reference_value read_pptx
get_pptx_uncached <- function() {
  ref_pptx <- read_pptx(get_reference_value(format = "pptx"))
  ref_pptx
}

#' @noRd
#' @importFrom memoise memoise
get_reference_pptx <- memoise(get_pptx_uncached)


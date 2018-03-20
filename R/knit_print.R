#' @importFrom rmarkdown pandoc_version
#' @importFrom knitr knit_print asis_output opts_knit
check_all <- function(){
  if (is.null(opts_knit$get("rmarkdown.pandoc.to")))
    stop("`knit_print.ooxml` needs to be used as a renderer for ",
         "an rmarkdown R code chunk (render by rmarkdown)", call. = FALSE)

  if (!(pandoc_version() >= 2))
    stop("pandoc version >= 2.0 required for ooxml rendering in docx", call. = FALSE)

  if (!(opts_knit$get("rmarkdown.pandoc.to") == "docx"))
    stop("unsupported format for ooxml rendering:", opts_knit$get("rmarkdown.pandoc.to"), call. = FALSE)
  invisible()
}

#' @export
format.ooxml <- function(x, type ="docx", ...){

  check_all()
  paste("```{=openxml}", as.character(x), "```", sep = "\n")

}

#' @export
format.ooxml_chunk <- function(x, type ="docx", ...){

  check_all()
  paste("`", as.character(x), "`{=openxml}", sep = "")

}

#' @export
knit_print.ooxml_chunk <- function(x, ...){

  knit_print( asis_output(
    format(x)
  ) )
}

#' @export
knit_print.ooxml <- function(x, ...){

  knit_print( asis_output(
    format(x)
  ) )
}


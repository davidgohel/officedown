#' @export
#' @importFrom officer to_wml to_pml
#' @importFrom knitr knit_print asis_output opts_knit
knit_print.run <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to")))
    knit_print( asis_output(
      paste("`", to_wml(x), "`{=openxml}", sep = "")
    ) )
  else if(grepl( "pptx", opts_knit$get("rmarkdown.pandoc.to"))){
    knit_print( asis_output(
      paste("`", to_pml(x), "`{=openxml}", sep = "")
    ) )
  } else knit_print( asis_output("") )
}


#' @export
knit_print.fp_par <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to"))){
    knit_print( asis_output(
      paste("`", to_wml(x), "`{=openxml}", sep = "")
    ) )
  } else if(grepl( "pptx", opts_knit$get("rmarkdown.pandoc.to"))){
    knit_print( asis_output(
      paste("`", to_pml(x), "`{=openxml}", sep = "")
    ) )
  } else knit_print( asis_output("") )
}

#' @export
knit_print.block <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to"))){
    knit_print( asis_output(
      paste("```{=openxml}", to_wml(x), "```", sep = "\n")
    ) )
  } else if(grepl( "pptx", opts_knit$get("rmarkdown.pandoc.to"))){
    knit_print( asis_output(
      paste("```{=openxml}", to_pml(x), "```", sep = "\n")
    ) )
  } else knit_print( asis_output("") )
}

#' @export
#' @title Force block printing while knitting
#' @description When used in a loop, blocks output anything
#' because `knit_print` method is not called. In that situation,
#' use the function to force printing. Also you should tell the chunk
#' to use results 'as-is' (by adding `results='asis'` to your chunk header).
#' @param x a block object, result of a block function from officer package
#' @param ... unused arguments
#' @section Word example:
#' Copy example located here:
#' `system.file(package = "officedown", "examples", "word_loop.Rmd")`
knit_print_block <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to"))){
    cat(paste("```{=openxml}", to_wml(x), "```\n\n", sep = "\n"))
  } else cat("")
}

#' @export
#' @title Force run printing while knitting
#' @description When used in a loop, runs output anything
#' because `knit_print` method is not called. In that situation,
#' use the function to force printing. Also you should tell the chunk
#' to use results 'as-is' (by adding `results='asis'` to your chunk header).
#' @param x a run object, result of a run function from officer package
#' @param ... unused arguments
knit_print_run <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to"))){
    cat(paste("`", to_wml(x), "`{=openxml}", sep = ""))
  } else if(grepl( "pptx", opts_knit$get("rmarkdown.pandoc.to"))){
    cat(paste("`", to_pml(x), "`{=openxml}", sep = ""))
  } else cat("")
}


#' @export
#' @importFrom officer to_wml
#' @importFrom knitr knit_print asis_output opts_knit
knit_print.run <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to")))
    knit_print( asis_output(
      paste("`", to_wml(x), "`{=openxml}", sep = "")
    ) )
  else knit_print( asis_output("") )
}

#' @export
print.run <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to")))
    cat(paste("`", to_wml(x), "`{=openxml}", sep = ""))
  else cat("")
}

#' @export
knit_print.fp_par <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to")))
    knit_print( asis_output(
      paste("`", to_wml(x), "`{=openxml}", sep = "")
    ) )
  else knit_print( asis_output("") )
}

#' @export
knit_print.block <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to")))
    knit_print( asis_output(
      paste("```{=openxml}", to_wml(x), "```", sep = "\n")
    ) )
  else knit_print( asis_output("") )
}

#' @export
print.block <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to")))
    cat(paste("```{=openxml}", to_wml(x), "```\n\n", sep = "\n"))
  else cat("")
}

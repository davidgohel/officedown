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
#' @title Force Block Printing while Knitting
#' @description When used in a loop, calls to blocks do not generate
#' output because `knit_print` method is not called.
#' Use the function to force printing. Also you should tell the chunk
#' to use results 'as-is' (by adding `results='asis'` to your chunk header).
#' @param x a block object, result of a block function from officer package
#' @param ... unused arguments
#' @return None. the function only print XML code.
#' @family functions that force printing
#' @examples
#' library(rmarkdown)
#' rmd_file_src <- system.file(
#'   package = "officedown", "examples", "word_loop.Rmd")
#' rmd_file_des <- tempfile(fileext = ".Rmd")
#' if(pandoc_available()){
#'
#'   file.copy(rmd_file_src, to = rmd_file_des)
#'   docx_file_1 <- tempfile(fileext = ".docx")
#'   render(rmd_file_des, output_file = docx_file_1, quiet = TRUE)
#'
#'   if(file.exists(docx_file_1)){
#'     message("file ", docx_file_1, " has been written.")
#'   }
#' }
knit_print_block <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to"))){
    cat(paste("```{=openxml}", to_wml(x), "```\n\n", sep = "\n"))
  } else cat("")
}

#' @export
#' @title Force Run Printing while Knitting
#' @description When used in a loop, runs do not outputs
#' because `knit_print` method is not called.
#' Use the function to force printing. Also you should tell the chunk
#' to use results 'as-is' (by adding `results='asis'` to your chunk header).
#' @param x a run object, result of a run function from officer package
#' @param ... unused arguments
#' @family functions that force printing
#' @return None. the function only print XML code.
#' @examples
#' library(rmarkdown)
#' rmd_file_src <- system.file(
#'   package = "officedown", "examples", "word_loop.Rmd")
#' rmd_file_des <- tempfile(fileext = ".Rmd")
#' if(pandoc_available()){
#'
#'   file.copy(rmd_file_src, to = rmd_file_des)
#'   docx_file_1 <- tempfile(fileext = ".docx")
#'   render(rmd_file_des, output_file = docx_file_1, quiet = TRUE)
#'
#'   if(file.exists(docx_file_1)){
#'     message("file ", docx_file_1, " has been written.")
#'   }
#' }
knit_print_run <- function(x, ...){
  if(grepl( "docx", opts_knit$get("rmarkdown.pandoc.to"))){
    cat(paste("`", to_wml(x), "`{=openxml}", sep = ""))
  } else if(grepl( "pptx", opts_knit$get("rmarkdown.pandoc.to"))){
    cat(paste("`", to_pml(x), "`{=openxml}", sep = ""))
  } else cat("")
}


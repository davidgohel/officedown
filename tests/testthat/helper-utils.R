render_rmd <- function(rmd_file, output_file) {
  unlink(output_file, force = TRUE)
  sucess <- FALSE
  tryCatch(
    {
      rmarkdown::render(rmd_file,
        output_file = output_file,
        envir = new.env(),
        quiet = TRUE
      )
      sucess <- TRUE
    },
    warning = function(e) {
    },
    error = function(e) {
    }
  )
  sucess
}
get_docx_xml <- function(x) {
  redoc <- read_docx(x)
  xml2::xml_child(docx_body_xml(redoc))
}

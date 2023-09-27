library(rmarkdown)
library(officedown)

if (pandoc_available() && pandoc_version() >= numeric_version("2.0")) {
  # minimal example -----
  example <- system.file(
    package = "officedown",
    "examples/minimal_word.Rmd"
  )
  rmd_file <- tempfile(fileext = ".Rmd")
  file.copy(example, to = rmd_file)

  docx_file <- tempfile(fileext = ".docx")
  render(rmd_file, output_file = docx_file, quiet = TRUE)
}

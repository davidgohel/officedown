library(rmarkdown)


skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

source("utils.R")

str <- c(
  "#' ---",
  "#' output: officedown::rdocx_document",
  "#' ---",
  "",
  "mtcars",
  ""
)
filename <- tempfile(fileext = ".R")
writeLines(str, filename, useBytes = TRUE)

test_that("rendering R file", {
  docx_file <- tempfile(fileext = ".docx")
  expect_true(render_rmd(filename, output_file = docx_file))
})

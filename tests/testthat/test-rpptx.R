library(xml2)
library(officer)
library(rmarkdown)

skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

source("utils.R")


test_that("visual testing tables", {
  testthat::skip_if_not_installed("doconv")
  testthat::skip_if_not(doconv::msoffice_available())
  library(doconv)
  pptx_file <- tempfile(fileext = ".pptx")
  render_rmd("rmd/pptx.Rmd", output_file = pptx_file)
  expect_snapshot_doc(x = pptx_file, name = "pptx-example", engine = "testthat")
})


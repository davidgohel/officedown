library(xml2)
library(officer)
library(rmarkdown)

skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

source("utils.R")


test_that("reading captions", {

  docx_file <- tempfile(fileext = ".docx")
  render_rmd("rmd/plot-basic.Rmd", output_file = docx_file)

  node_body <- get_docx_xml(docx_file)

  expect_equal(
    xml_text(xml_child(node_body, "/w:p[w:bookmarkStart/@w:name='boxplot']")),
    "Figure SEQ fig \\* Arabic: A boxplot"
  )
  expect_equal(
    xml_text(xml_child(node_body, "/w:p[w:bookmarkStart/@w:name='barplot']")),
    "Figure SEQ fig \\* Arabic: What a barplot"
  )
})

test_that("visual testing plots", {
  testthat::skip_if_not_installed("doconv")
  testthat::skip_if_not(doconv::msoffice_available())
  library(doconv)
  docx_file <- tempfile(fileext = ".docx")
  render_rmd("rmd/plot-basic.Rmd", output_file = docx_file)

  expect_snapshot_doc(x = docx_file, name = "plot-basic", engine = "testthat")
})

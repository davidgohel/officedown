library(xml2)
library(officer)
library(rmarkdown)

skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

docx_file <- tempfile(fileext = ".docx")
render_rmd("rmd/tables-basic.Rmd", output_file = docx_file)

test_that("reading captions", {
  node_body <- get_docx_xml(docx_file)

  expect_equal(
    xml_text(xml_child(node_body, "/w:p[w:bookmarkStart/@w:name='mtcars']")),
    "Table SEQ tab \\* Arabic: caption 1"
  )
  expect_equal(
    xml_text(xml_child(node_body, "/w:p[w:bookmarkStart/@w:name='cars']")),
    "Table SEQ tab \\* Arabic: cars"
  )
})


test_that("visual testing tables", {
  testthat::skip_if_not_installed("doconv")
  testthat::skip_if_not(doconv::msoffice_available())
  library(doconv)
  expect_snapshot_doc(x = docx_file, name = "docx-tables-basic", engine = "testthat")
})


library(xml2)
library(officer)
library(rmarkdown)

skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

docx_file <- tempfile(fileext = ".docx")
render_rmd("rmd/par-settings.Rmd", output_file = docx_file)

test_that("scanning settings", {
  node_body <- get_docx_xml(docx_file)
  spacing_node <- xml_child(node_body, "/w:p/w:pPr[w:pStyle[@w:val='Normal']]/w:spacing")
  expect_equal(xml_attr(spacing_node, "after"), "400")
  expect_equal(xml_attr(spacing_node, "before"), "2400")
  expect_equal(xml_attr(spacing_node, "line"), "240")
})

docx_file <- tempfile(fileext = ".docx")
render_rmd("rmd/change-styles.Rmd", output_file = docx_file)

test_that("change styles", {
  node_body <- get_docx_xml(docx_file)
  normal_node <- xml_child(node_body, "/w:p[2]/w:pPr[w:pStyle[@w:val='Normal']]")
  expect_false(inherits(normal_node, "xml_missing"))
  normal_node <- xml_child(node_body, "/w:p[3]/w:pPr[w:pStyle[@w:val='Normal']]")
  expect_false(inherits(normal_node, "xml_missing"))
})

library(xml2)
library(officer)
library(rmarkdown)

skip_on_cran()

skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

source("utils.R")

docx_file <- tempfile(fileext = ".docx")
render_rmd("rmd/sections.Rmd", output_file = docx_file)

test_that("reading sections properties", {
  node_body <- get_docx_xml(docx_file)

  continuous_sec_node <- xml_find_first(node_body, "w:p[w:pPr/w:pStyle/@w:val='Titre']/following-sibling::w:p")
  expect_false(
    inherits(
      xml_child(continuous_sec_node, "w:pPr/w:sectPr/w:type[@w:val='continuous']"),
      "xml_missing"
    )
  )

  next_node <- xml_find_first(node_body, "w:p[w:pPrw/sectPr/w:type[@w:val='continuous']]")
  next_node <- xml_find_first(node_body, "w:p[w:pPr/sectPr/w:type[@w:val='continuous']]")
  next_node <- xml_find_first(node_body, "//w:p[w:pPr/w:sectPr/w:cols[@w:num='2']]/preceding-sibling::w:p[1]")

  expect_false(
    inherits(xml_child(next_node, "w:r/w:br"), "xml_missing")
  )
  expect_equal(
    xml_text(next_node),
    officedown:::str_lorem
  )

  x <- read_docx(docx_file)
  x <- cursor_end(x)
  x <- cursor_backward(x)
  x <- cursor_backward(x)
  sect_node <- docx_current_block_xml(x)
  pr_nodes <- xml_child(sect_node, "w:pPr/w:sectPr")
  part_references <- grep("(header|footer)Reference", xml_name(xml_children(pr_nodes)), value = TRUE)
  expect_equal(sum(part_references %in% "footerReference"), 3L)
  expect_equal(sum(part_references %in% "headerReference"), 3L)
})

test_that("visual testing sections properties", {
  testthat::skip_if_not_installed("doconv")
  testthat::skip_if_not(doconv::msoffice_available())
  library(doconv)
  expect_snapshot_doc(x = docx_file, name = "docx-sections", engine = "testthat")
})


docx_file <- tempfile(fileext = ".docx")
render_rmd("rmd/long-text.Rmd", output_file = docx_file)

test_that("visual testing sections inheritance", {
  testthat::skip_if_not_installed("doconv")
  testthat::skip_if_not(doconv::msoffice_available())
  library(doconv)
  expect_snapshot_doc(x = docx_file, name = "all-sections", engine = "testthat")
})

library(xml2)
library(officer)
library(rmarkdown)

skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

source("utils.R")

docx_file <- tempfile(fileext = ".docx")
render_rmd("rmd/lists.Rmd", output_file = docx_file)

test_that("reading lists", {
  node_body <- get_docx_xml(docx_file)

  li1 <- xml_child(node_body, "w:bookmarkStart[@w:name='ul']/following-sibling::w:p[2]/w:pPr/w:numPr/w:ilvl")
  li2 <- xml_child(node_body, "w:bookmarkStart[@w:name='ul']/following-sibling::w:p[3]/w:pPr/w:numPr/w:ilvl")
  li3 <- xml_child(node_body, "w:bookmarkStart[@w:name='ul']/following-sibling::w:p[4]/w:pPr/w:numPr/w:ilvl")
  li4 <- xml_child(node_body, "w:bookmarkStart[@w:name='ul']/following-sibling::w:p[5]/w:pPr/w:numPr/w:ilvl")
  li5 <- xml_child(node_body, "w:bookmarkStart[@w:name='ul']/following-sibling::w:p[6]/w:pPr/w:numPr/w:ilvl")
  li6 <- xml_child(node_body, "w:bookmarkStart[@w:name='ul']/following-sibling::w:p[7]/w:pPr/w:numPr/w:ilvl")

  expect_equal(xml_attr(li1, "val"), "0")
  expect_equal(xml_attr(li2, "val"), "1")
  expect_equal(xml_attr(li3, "val"), "2")
  expect_equal(xml_attr(li4, "val"), "2")
  expect_equal(xml_attr(li5, "val"), "1")
  expect_equal(xml_attr(li6, "val"), "0")


  dir_tmp <- tempfile()
  unpack_folder(docx_file, dir_tmp)
  num_xml <- file.path(dir_tmp, "word", "numbering.xml")
  x <- read_xml(num_xml)

  num_id <- xml_child(node_body, "w:bookmarkStart[@w:name='ul']/following-sibling::w:p[2]/w:pPr/w:numPr/w:numId")
  num_id <- xml_attr(num_id, "val")

  abstractnum <- xml_find_first(x, sprintf("//w:num[@w:numId='%s']/w:abstractNumId", num_id))
  abstractnum <- xml_attr(abstractnum, "val")

  stylelink <- xml_find_first(x, sprintf("//w:abstractNum[@w:abstractNumId='%s']/w:styleLink", abstractnum))
  expect_equal(xml_attr(stylelink, "val"), "new-ul")

  num_id <- xml_child(node_body, "w:bookmarkStart[@w:name='ol']/following-sibling::w:p[2]/w:pPr/w:numPr/w:numId")
  num_id <- xml_attr(num_id, "val")

  abstractnum <- xml_find_first(x, sprintf("//w:num[@w:numId='%s']/w:abstractNumId", num_id))
  abstractnum <- xml_attr(abstractnum, "val")

  stylelink <- xml_find_first(x, sprintf("//w:abstractNum[@w:abstractNumId='%s']/w:styleLink", abstractnum))
  expect_equal(xml_attr(stylelink, "val"), "new-ol")
})


test_that("visual testing list", {
  testthat::skip_if_not_installed("doconv")
  testthat::skip_if_not(doconv::msoffice_available())
  library(doconv)
  expect_snapshot_doc(x = docx_file, name = "lists", engine = "testthat")
})


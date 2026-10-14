library(xml2)
library(officer)
library(rmarkdown)

skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

docx_file <- tempfile(fileext = ".docx")
render_rmd("rmd/lists.Rmd", output_file = docx_file)

# Return the `w:abstractNum/w:styleLink` value associated with a list
# paragraph, by following numId -> abstractNumId in numbering.xml.
list_style_link <- function(list_par, numbering) {
  num_id <- xml_attr(xml_child(list_par, "w:pPr/w:numPr/w:numId"), "val")
  abstract_num_id <- xml_attr(
    xml_find_first(
      numbering,
      sprintf("//w:num[@w:numId='%s']/w:abstractNumId", num_id)
    ),
    "val"
  )
  xml_attr(
    xml_find_first(
      numbering,
      sprintf(
        "//w:abstractNum[@w:abstractNumId='%s']/w:styleLink",
        abstract_num_id
      )
    ),
    "val"
  )
}

test_that("reading lists", {
  node_body <- get_docx_xml(docx_file)

  # list paragraphs are identified by their numbering properties rather
  # than by the heading bookmarks, whose names and positions depend on the
  # pandoc version (pandoc >= 3.12 prefixes them with `_`).
  list_pars <- xml_find_all(node_body, "w:p[w:pPr/w:numPr]")
  expect_length(list_pars, 12)

  ul_pars <- list_pars[1:6]
  ol_pars <- list_pars[7:12]

  ilvl <- xml_attr(xml_find_first(ul_pars, "w:pPr/w:numPr/w:ilvl"), "val")
  expect_equal(ilvl, c("0", "1", "2", "2", "1", "0"))
  ilvl <- xml_attr(xml_find_first(ol_pars, "w:pPr/w:numPr/w:ilvl"), "val")
  expect_equal(ilvl, c("0", "1", "2", "2", "1", "0"))

  dir_tmp <- tempfile()
  unpack_folder(docx_file, dir_tmp)
  numbering <- read_xml(file.path(dir_tmp, "word", "numbering.xml"))

  expect_equal(list_style_link(ul_pars[[1]], numbering), "new-ul")
  expect_equal(list_style_link(ol_pars[[1]], numbering), "new-ol")
})


test_that("list styles are matched by their exact name", {
  # the reference_docx defines a list style named 'new-ul-old' before the
  # 'new-ul' list style; bullet lists must use 'new-ul'
  rmd_file <- tempfile(fileext = ".Rmd")
  rmd <- readLines("rmd/lists.Rmd")
  rmd <- sub(
    "reference_docx: template.docx",
    paste0(
      "reference_docx: ",
      normalizePath("rmd/template-similar-list-styles.docx")
    ),
    rmd,
    fixed = TRUE
  )
  writeLines(rmd, rmd_file)
  out_file <- tempfile(fileext = ".docx")
  expect_true(render_rmd(rmd_file, output_file = out_file))

  node_body <- get_docx_xml(out_file)
  list_pars <- xml_find_all(node_body, "w:p[w:pPr/w:numPr]")

  dir_tmp <- tempfile()
  unpack_folder(out_file, dir_tmp)
  numbering <- read_xml(file.path(dir_tmp, "word", "numbering.xml"))

  expect_equal(list_style_link(list_pars[[1]], numbering), "new-ul")
  expect_equal(list_style_link(list_pars[[7]], numbering), "new-ol")
})


test_that("visual testing list", {
  testthat::skip_if_not_installed("doconv")
  testthat::skip_if_not(doconv::msoffice_available())
  library(doconv)
  expect_snapshot_doc(x = docx_file, name = "lists", engine = "testthat")
})

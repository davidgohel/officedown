library(xml2)
library(officer)
library(rmarkdown)

skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

docx_file <- tempfile(fileext = ".docx")
render_rmd("rmd/plot-basic.Rmd", output_file = docx_file)

test_that("reading captions", {

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

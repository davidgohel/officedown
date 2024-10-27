library(xml2)
library(officer)
library(rmarkdown)

skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

pptx_file <- tempfile(fileext = ".pptx")
render_rmd("rmd/pptx.Rmd", output_file = pptx_file)


test_that("find text in PowerPoint file", {
  doc <- read_pptx(pptx_file)
  all_doc_sum <- pptx_summary(doc)

  doc_sum <- all_doc_sum[all_doc_sum$slide_id == 3,]
  expect_contains(doc_sum$text, "temperature")
  expect_contains(doc_sum$text, "pressure")

  doc_sum <- all_doc_sum[all_doc_sum$slide_id == 4,]
  expect_contains(doc_sum$text, "temperature")
  expect_contains(doc_sum$text, "pressure")
})

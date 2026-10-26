library(xml2)
library(officer)
library(rmarkdown)

skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

test_that("paragraph style ids guessed from style names are fixed", {
  # in the default reference_docx, the style named 'caption' has the id
  # 'Lgende'; flextable writes ids guessed from names, here 'caption' (#112)
  raw_par <- function(style_id, txt) {
    sprintf(
      "<w:p><w:pPr><w:pStyle w:val=\"%s\"/></w:pPr><w:r><w:t>%s</w:t></w:r></w:p>",
      style_id,
      txt
    )
  }
  rmd_file <- tempfile(fileext = ".Rmd")
  writeLines(
    c(
      "---",
      "output: officedown::rdocx_document",
      "---",
      "",
      "```{=openxml}",
      raw_par("caption", "guessed id"),
      raw_par("TableCaption", "existing id"),
      raw_par("NoSuchStyle", "unknown id"),
      "```"
    ),
    rmd_file
  )
  out_file <- tempfile(fileext = ".docx")
  expect_true(render_rmd(rmd_file, output_file = out_file))

  node_body <- get_docx_xml(out_file)
  style_of <- function(txt) {
    xml_attr(
      xml_find_first(
        node_body,
        sprintf("//w:p[w:r/w:t='%s']/w:pPr/w:pStyle", txt)
      ),
      "val"
    )
  }
  expect_equal(style_of("guessed id"), "Lgende")
  expect_equal(style_of("existing id"), "TableCaption")
  expect_equal(style_of("unknown id"), "NoSuchStyle")
})

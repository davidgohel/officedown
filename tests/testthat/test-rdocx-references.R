library(xml2)
library(officer)
library(rmarkdown)

skip_if_not(rmarkdown::pandoc_available())
skip_if_not(pandoc_version() >= numeric_version("2"))

test_that("section references target existing bookmarks", {
  rmd_file <- tempfile(fileext = ".Rmd")
  writeLines(
    c(
      "---",
      "output: officedown::rdocx_document",
      "---",
      "",
      "# Underscore {#intro_part}",
      "",
      "# Dot {#sec.dot}",
      "",
      "# Dash {#sec-dash}",
      "",
      "See \\@ref(intro_part), \\@ref(sec.dot) and \\@ref(sec-dash)."
    ),
    rmd_file
  )
  docx_file <- tempfile(fileext = ".docx")
  expect_true(render_rmd(rmd_file, output_file = docx_file))

  node_body <- get_docx_xml(docx_file)
  # no reference is left as plain text
  expect_false(any(grepl("@ref(", xml_text(node_body), fixed = TRUE)))

  instr <- xml_text(xml_find_all(node_body, "//w:instrText"))
  ref_targets <- sub("^\\s*REF\\s+(\\S+).*$", "\\1", grep("REF ", instr, value = TRUE))
  expect_length(ref_targets, 3)

  bookmarks <- xml_attr(xml_find_all(node_body, "//w:bookmarkStart"), "name")
  expect_true(all(ref_targets %in% bookmarks))
})

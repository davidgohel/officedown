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

test_that("plot size follows fig.width, fig.height and out options", {
  ext_png <- tempfile(fileext = ".png")
  png(ext_png, width = 300, height = 150)
  plot(1)
  dev.off()

  rmd_file <- tempfile(fileext = ".Rmd")
  writeLines(
    c(
      "---",
      "output: officedown::rdocx_document",
      "---",
      "",
      "```{r small, fig.width=3, fig.height=2}",
      "plot(1)",
      "```",
      "",
      "```{r two, fig.width=c(3, 5), fig.height=c(2, 3)}",
      "plot(1)",
      "plot(2)",
      "```",
      "",
      "```{r cm, out.width='10cm'}",
      "plot(1)",
      "```",
      "",
      "```{r ext, fig.width=6, fig.height=4}",
      sprintf("knitr::include_graphics('%s')", ext_png),
      "```"
    ),
    rmd_file
  )
  out_file <- tempfile(fileext = ".docx")
  expect_true(render_rmd(rmd_file, output_file = out_file))

  extents <- xml_find_all(get_docx_xml(out_file), "//wp:extent")
  width_in <- as.numeric(xml_attr(extents, "cx")) / 914400
  height_in <- as.numeric(xml_attr(extents, "cy")) / 914400

  # fig.width and fig.height are used, also per plot
  expect_equal(width_in[1:3], c(3, 3, 5), tolerance = 0.01)
  expect_equal(height_in[1:3], c(2, 2, 3), tolerance = 0.01)
  # out.width is used when set
  expect_equal(width_in[4], 10 / 2.54, tolerance = 0.01)
  # an image from include_graphics() keeps its own size, not fig.width
  expect_false(isTRUE(all.equal(width_in[5], 6, tolerance = 0.01)))
})

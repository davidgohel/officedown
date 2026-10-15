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
  rmd_dir <- tempfile()
  dir.create(rmd_dir)
  png(file.path(rmd_dir, "ext.png"), width = 300, height = 150)
  plot(1)
  dev.off()

  rmd_file <- file.path(rmd_dir, "plots.Rmd")
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
      "knitr::include_graphics('ext.png')",
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

test_that("plot hook keeps the external_img() line used by accessr", {
  # accessr::rmd2word() replaces the top-level line of the plot hook that
  # calls external_img(); it must exist exactly once
  plot_body <- body(rdocx_document()$knitr$knit_hooks$plot)
  found <- vapply(
    seq_along(plot_body),
    function(i) !is.na(pmatch("external_img", as.character(plot_body[[i]]))),
    logical(1)
  )
  expect_equal(sum(found), 1L)
})
